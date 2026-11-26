# -*- coding: utf-8 -*-
"""
Leitor de Folhas de Corte a Laser -> WK Radar (Entrada da Produção)
-------------------------------------------------------------------
Lê fotos de folhas de controle de corte a laser (com anotações/rasuras à caneta)
usando a API do Google Gemini e digita automaticamente os dados na tela de
Entrada da Produção do ERP WK Radar.
"""

import base64
import ctypes
import io
import json
import os
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import winsound

import pyautogui
import pyperclip
from dotenv import load_dotenv
from google import genai
from PIL import Image, ImageGrab, ImageOps
from pydantic import BaseModel, Field, field_validator
from pydantic.json_schema import SkipJsonSchema

# =============================================================================
# CONFIGURAÇÕES (ajuste aqui)
# =============================================================================
# Modelo de visão mais rápido da API para leitura de tabelas/documentos:
MODELO_GEMINI = "gemini-3.5-flash-lite"
# Nível de raciocínio da IA: "minimal", "low" (mais rápido), "medium" ou "high".
NIVEL_RACIOCINIO = "low"
# Se o modelo principal estiver sobrecarregado (erro 503), tenta estes em ordem:
MODELOS_RESERVA = ["gemini-4-hydrogen", "gemini-3.8-flash"]
# Tempo máximo (segundos) esperando cada modelo antes de tentar o próximo.
TEMPO_LIMITE_IA = 60

CONTAGEM_SEGUNDOS = 10          # tempo total da calibração (5s + 5s)
CONTAGEM_ETAPA_SEGUNDOS = 5     # segundos por etapa (1/2 Operação, 2/2 Produto)
SETAS_OPERACAO_PADRAO = 1       # toques na seta p/ baixo no campo Operação
ESPERA_APOS_PRODUTO = 1.0       # tempo (s) após Enter do Produto p/ consulta no banco
ENTERS_APOS_QTD_PADRAO = 18     # Enters finais até OK + confirmação do OK
ESPERA_APOS_OK = 1.5            # tempo (s) para salvar e limpar o formulário
VALOR_UNITARIO_PADRAO = "0,00"
LOCAL_DESTINO_PADRAO = "23"

COR_ALTERADO = "#fff3cd"        # fundo amarelo claro das linhas com rasura à caneta
COR_LANCADO = "#d4edda"         # fundo verde claro das linhas já lançadas (✅)
COR_POPUP = "#f8d7da"           # fundo vermelho claro quando abre pop-up no WK Radar
TAMANHO_MAX_IMAGEM = 2400       # px (maior lado) antes de enviar para a API

# Códigos de tecla virtual do Windows (atalhos globais via user32.GetAsyncKeyState)
VK_ESCAPE = 0x1B                # Esc  -> Parar digitação imediatamente
VK_F12 = 0x7B                   # F12  -> Pausar / Retomar digitação
VK_BACK = 0x08                  # Backspace
VK_SHIFT = 0x10                 # Shift
VK_END = 0x23                   # End
VK_HOME = 0x24                  # Home
VK_DELETE = 0x2E                # Delete
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002

INSTRUCAO_RODAPE = (
    "Clique em '🎯 Calibrar Posições' 1 vez (salvo em config.json). "
    "Depois deixe o WK Radar aberto atrás e clique em '▶ Digitar no WK Radar'."
)

pyautogui.FAILSAFE = True       # mouse no canto superior esquerdo = aborta

# =============================================================================
# CAMINHOS / .ENV
# =============================================================================
if getattr(sys, "frozen", False):          # rodando como .exe (PyInstaller)
    PASTA_BASE = os.path.dirname(sys.executable)
else:
    PASTA_BASE = os.path.dirname(os.path.abspath(__file__))

load_dotenv(os.path.join(PASTA_BASE, ".env"))
CAMINHO_CONFIG = os.path.join(PASTA_BASE, "config.json")


# =============================================================================
# SCHEMA (Structured Outputs com Pydantic)
# =============================================================================
class ItemCorte(BaseModel):
    codigo: str = Field(
        description="Código da peça impresso na folha (sem letras de prefixo no início "
                    "antes do número, mantendo pontos, hífens, barras, espaços internos e "
                    "letras no meio ou no final do código)."
    )
    valor_unitario: str = Field(
        default=VALOR_UNITARIO_PADRAO,
        description="Preço / valor unitário da peça na coluna 'Preço' (valores com 'R$'). "
                    "O ponto (.) no preço impresso funciona como vírgula decimal (ex.: "
                    "'R$ 199.906' = '199,90'; 'R$ 161.673' = '161,67'; 'R$ 10.97' = '10,97'). "
                    "Retorne apenas o número com vírgula e 2 casas decimais (sem 'R$'). "
                    "Se não houver preço na folha, use '0,00'."
    )
    local_destino: str = Field(
        default=LOCAL_DESTINO_PADRAO,
        description="Local de destino no estoque. Use sempre o valor padrão '23', "
                    "a menos que outro destino esteja explicitamente indicado."
    )
    quantidade_final: int = Field(
        description="Quantidade final da peça. Se a quantidade impressa estiver "
                    "riscada/rasurada, priorize sempre o número escrito à caneta."
    )
    alterado_a_mao: bool = Field(
        description="True se houve risco, rasura ou anotação manual na quantidade."
    )
    observacao: str = Field(
        description="Nota breve sobre rasuras (ex.: 'Riscado 15, escrito 12 - queimou'). "
                    "Vazio se não houver nada a observar."
    )
    lancado: SkipJsonSchema[bool] = Field(default=False, exclude=True)
    erro_popup: SkipJsonSchema[bool] = Field(default=False, exclude=True)
    resolvida_manual: SkipJsonSchema[bool] = Field(default=False, exclude=True)

    @field_validator("valor_unitario", mode="before")
    @classmethod
    def _fallback_valor_unitario(cls, v):
        if v is None or (isinstance(v, str) and not v.strip()):
            return VALOR_UNITARIO_PADRAO
        return str(v)


class ResultadoLeitura(BaseModel):
    itens: list[ItemCorte] = Field(description="Lista de peças lidas nas folhas.")


PROMPT_SISTEMA = """\
Você é um assistente especialista em ler folhas de controle de corte a laser de uma \
indústria metalúrgica brasileira. As folhas são fotografadas com celular e contêm uma \
tabela impressa (com colunas como #, IMG, Referência, Dobra, Preço, Dimensões, \
Peças/Pedido, Tempo, Peso/KG), muitas vezes com anotações, marca-texto, clipes/post-its \
e rasuras feitas à caneta pelo operador.

REGRAS OBRIGATÓRIAS:
1. CÓDIGO ("codigo"): fica na coluna "Referência" (ex.: "3.24.031", "L3.12.084D", \
"P3.27.133"). Ignore qualquer letra no início antes do código numérico (prefixos), \
retornando sem essas letras iniciais (ex.: "L3.12.084D" -> "3.12.084D", \
"P3.27.133" -> "3.27.133"), mas mantenha SEMPRE letras no meio ou no final do código \
(como o "D" final em "3.12.084D"), além de pontos (.), hífens (-), barras (/), \
underscores e zeros à esquerda. NÃO use o número da primeira coluna "#" como código.
2. PREÇO / VALOR UNITÁRIO ("valor_unitario"):
   - Fica na coluna "Preço" (os valores começam com "R$", mesmo que o título da coluna \
esteja coberto por um clipe ou post-it).
   - ATENÇÃO CRÍTICA: nessa folha, o PONTO (.) no preço funciona como VÍRGULA DECIMAL \
(o software imprime 2 ou 3 casas decimais após o ponto). Portanto, NÃO é separador de \
milhar!
   - Exemplos reais:
     * "R$ 199.906" significa 199 reais e 90 centavos -> retorne "199,90" (NUNCA 199 mil!).
     * "R$ 161.673" -> retorne "161,67".
     * "R$ 270.999" -> retorne "270,99".
     * "R$ 3.927"   -> retorne "3,92".
     * "R$ 10.97"   -> retorne "10,97".
     * "R$ 1.65"    -> retorne "1,65".
   - Sempre remova o "R$", troque o ponto por vírgula e mantenha APENAS as 2 primeiras \
casas decimais (descarte a 3ª casa decimal sem arredondar para cima).
   - Se o preço estiver riscado e corrigido à caneta, use o valor escrito à caneta.
   - Só retorne "0,00" se a folha realmente não tiver coluna/valor de preço para a peça.
3. QUANTIDADE ("quantidade_final") e RASURAS:
   - A quantidade fica nas colunas "Peças" / "Pedido" (não confunda com a coluna "#" \
nem com "Peso / KG" ou dimensões).
   - Quando a quantidade original impressa estiver riscada, rasurada, circulada com \
correção ou sobrescrita à caneta, SEMPRE priorize o número escrito à caneta \
(geralmente peças que queimaram/perderam no corte). Nesse caso, marque \
alterado_a_mao = true e descreva em "observacao" (ex.: "Riscado 15, escrito 12 - queimou").
   - Traços ou vistos de conferência ("/") na coluna "Dobra" (SIM/NÃO) NÃO são rasuras \
de quantidade. Se a quantidade não foi alterada à mão, marque alterado_a_mao = false.
   - Se a correção à caneta indicar zero, ou a linha inteira estiver riscada \
(peça cancelada), use quantidade_final = 0 e explique na observação.
4. LOCAL DE DESTINO ("local_destino"): retorne sempre "23" como padrão.
5. VÁRIAS FOLHAS: você pode receber 2 ou mais imagens do mesmo lote. Se o MESMO \
código de peça aparecer repetido (na mesma folha ou entre folhas diferentes), \
CONSOLIDE em um único item somando as quantidades finais. Mantenha \
alterado_a_mao = true se QUALQUER uma das ocorrências tiver rasura, e mencione a \
soma na observação (ex.: "somado: folha 1 = 5, folha 2 = 3").
6. Se algum dígito estiver ilegível ou duvidoso, faça a melhor leitura possível e \
registre a dúvida na observação.
7. Não invente peças. Ignore cabeçalhos, totais, datas, assinaturas e campos que \
não sejam linhas de peças.
8. Mantenha a ordem em que as peças aparecem nas folhas.
9. Deixe "observacao" como string vazia quando não houver nada relevante.
"""


# =============================================================================
# FUNÇÕES DE APOIO
# =============================================================================
def _tecla_pressionada_global(vk_code: int) -> bool:
    """Verifica no Windows se a tecla física está pressionada (mesmo em outra janela)."""
    try:
        return (ctypes.windll.user32.GetAsyncKeyState(vk_code) & 0x8000) != 0
    except Exception:
        return False


def obter_hwnd_ativo() -> int:
    """Retorna o identificador (HWND) da janela atualmente ativa no Windows."""
    try:
        return int(ctypes.windll.user32.GetForegroundWindow())
    except Exception:
        return 0


def obter_titulo_janela(hwnd: int) -> str:
    """Lê o título da janela ativa via GetWindowTextW."""
    if not hwnd:
        return "(sem título)"
    try:
        user32 = ctypes.windll.user32
        tamanho = user32.GetWindowTextLengthW(hwnd)
        if tamanho <= 0:
            return "(sem título)"
        buffer = ctypes.create_unicode_buffer(tamanho + 1)
        user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
        return buffer.value.strip() or "(sem título)"
    except Exception:
        return "(sem título)"


def verificar_popup(hwnd_wk: int) -> bool:
    """Retorna True sempre que outra janela/pop-up assumir o foco."""
    return obter_hwnd_ativo() != hwnd_wk


def janela_valida(hwnd: int) -> bool:
    """Retorna True se `hwnd` corresponde a uma janela ainda aberta no Windows."""
    if not hwnd:
        return False
    try:
        return bool(ctypes.windll.user32.IsWindow(int(hwnd)))
    except Exception:
        return False


def focar_janela(hwnd: int) -> bool:
    """Traz a janela identificada por `hwnd` para o primeiro plano."""
    if not hwnd:
        return False
    try:
        return bool(ctypes.windll.user32.SetForegroundWindow(int(hwnd)))
    except Exception:
        return False


def _normalizar_ponto(valor) -> tuple[int, int] | None:
    """Converte [x, y], (x, y) ou objeto com .x/.y em tupla (int, int)."""
    if valor is None:
        return None
    if hasattr(valor, "x") and hasattr(valor, "y"):
        try:
            return (int(valor.x), int(valor.y))
        except (TypeError, ValueError):
            return None
    if isinstance(valor, (list, tuple)) and len(valor) == 2:
        try:
            return (int(valor[0]), int(valor[1]))
        except (TypeError, ValueError):
            return None
    return None


def carregar_config(caminho: str | None = None) -> dict:
    """Carrega e valida config.json da pasta do aplicativo; retorna {} se ausente/inválido."""
    caminho_alvo = caminho or CAMINHO_CONFIG
    if not os.path.isfile(caminho_alvo):
        return {}
    try:
        with open(caminho_alvo, "r", encoding="utf-8") as arq:
            dados = json.load(arq)
    except Exception:
        return {}
    if not isinstance(dados, dict):
        return {}

    cfg: dict = {}
    pos_op = _normalizar_ponto(dados.get("pos_operacao"))
    pos_prod = _normalizar_ponto(dados.get("pos_produto"))
    if pos_op is not None:
        cfg["pos_operacao"] = pos_op
    if pos_prod is not None:
        cfg["pos_produto"] = pos_prod

    if "setas_operacao" in dados:
        try:
            cfg["setas_operacao"] = max(0, min(int(dados["setas_operacao"]), 50))
        except (TypeError, ValueError):
            pass

    if "enters_apos_qtd" in dados:
        try:
            cfg["enters_apos_qtd"] = max(1, min(int(dados["enters_apos_qtd"]), 50))
        except (TypeError, ValueError):
            pass

    if "espera_apos_produto" in dados:
        try:
            val_prod = str(dados["espera_apos_produto"]).strip().replace(",", ".")
            cfg["espera_apos_produto"] = max(0.1, min(float(val_prod), 30.0))
        except (TypeError, ValueError):
            pass

    if "espera_apos_ok" in dados:
        try:
            val_ok = str(dados["espera_apos_ok"]).strip().replace(",", ".")
            cfg["espera_apos_ok"] = max(0.1, min(float(val_ok), 30.0))
        except (TypeError, ValueError):
            pass

    return cfg


def salvar_config(pos_operacao, pos_produto, setas_operacao: int,
                  enters_apos_qtd: int, espera_apos_ok: float,
                  espera_apos_produto: float = ESPERA_APOS_PRODUTO,
                  caminho: str | None = None) -> None:
    """Salva calibração de posições e parâmetros do robô em config.json."""
    caminho_alvo = caminho or CAMINHO_CONFIG
    op = _normalizar_ponto(pos_operacao)
    prod = _normalizar_ponto(pos_produto)
    dados = {
        "pos_operacao": [op[0], op[1]] if op is not None else None,
        "pos_produto": [prod[0], prod[1]] if prod is not None else None,
        "setas_operacao": int(setas_operacao),
        "espera_apos_produto": round(float(espera_apos_produto), 2),
        "enters_apos_qtd": int(enters_apos_qtd),
        "espera_apos_ok": round(float(espera_apos_ok), 2),
    }
    pasta = os.path.dirname(os.path.abspath(caminho_alvo))
    if pasta:
        os.makedirs(pasta, exist_ok=True)
    with open(caminho_alvo, "w", encoding="utf-8") as arq:
        json.dump(dados, arq, indent=2, ensure_ascii=False)


def _enviar_tecla_win(vk_code: int, estendida: bool = False) -> None:
    """Envia tecla via keybd_event com scan code real e flag estendida (imune ao NumLock)."""
    user32 = ctypes.windll.user32
    scan = int(user32.MapVirtualKeyW(vk_code, 0))
    flags_down = KEYEVENTF_EXTENDEDKEY if estendida else 0
    flags_up = flags_down | KEYEVENTF_KEYUP
    user32.keybd_event(vk_code, scan, flags_down, 0)
    time.sleep(0.015)
    user32.keybd_event(vk_code, scan, flags_up, 0)


def _enviar_shift_home_win() -> None:
    """Envia Shift+Home via keybd_event com scan codes reais (imune ao NumLock)."""
    user32 = ctypes.windll.user32
    scan_shift = int(user32.MapVirtualKeyW(VK_SHIFT, 0))
    scan_home = int(user32.MapVirtualKeyW(VK_HOME, 0))
    user32.keybd_event(VK_SHIFT, scan_shift, 0, 0)
    time.sleep(0.015)
    user32.keybd_event(VK_HOME, scan_home, KEYEVENTF_EXTENDEDKEY, 0)
    time.sleep(0.015)
    user32.keybd_event(VK_HOME, scan_home, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
    time.sleep(0.015)
    user32.keybd_event(VK_SHIFT, scan_shift, KEYEVENTF_KEYUP, 0)


def digitar_campo_seguro(texto, limpar: bool = True) -> None:
    """Limpa o campo (se `limpar=True`) e digita `texto` direto, usando Ctrl+V apenas como fallback."""
    texto_limpo = "" if texto is None else str(texto).strip()
    if limpar:
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.06)
        _enviar_tecla_win(VK_BACK)
        time.sleep(0.05)
        _enviar_tecla_win(VK_END, estendida=True)
        time.sleep(0.05)
        _enviar_shift_home_win()
        time.sleep(0.05)
        _enviar_tecla_win(VK_BACK)
        time.sleep(0.05)
        for _ in range(8):
            _enviar_tecla_win(VK_BACK)
            time.sleep(0.02)
        for _ in range(8):
            _enviar_tecla_win(VK_DELETE, estendida=True)
            time.sleep(0.02)
        time.sleep(0.05)

    if texto_limpo:
        if re.fullmatch(r"[A-Za-z0-9,.\-]+", texto_limpo):
            pyautogui.write(texto_limpo, interval=0.04)
        else:
            pyperclip.copy(texto_limpo)
            time.sleep(0.15)
            pyautogui.hotkey("ctrl", "v")
    time.sleep(0.2)


def limpar_e_digitar_destino(valor) -> None:
    """Limpa completamente o campo Local de Destino e digita `valor` sem duplicar."""
    digitar_campo_seguro(valor, limpar=True)


def _formatar_valor_unitario(texto: str) -> str:
    """Converte preços como 'R$ 199.906', '199.906', '199,906' ou '10.97' para '199,90' e '10,97'.

    Na folha de corte, o ponto funciona como vírgula decimal (com 2 ou 3 casas).
    Mantém apenas 2 casas decimais (truncando a 3ª casa, sem arredondar para cima).
    """
    s = (texto or "").strip()
    if not s:
        return VALOR_UNITARIO_PADRAO
    s = s.replace("R$", "").replace("r$", "").strip()
    if not s:
        return VALOR_UNITARIO_PADRAO
    # Se vier no formato "199.906,00" (IA interpretou ponto como milhar e adicionou ,00),
    # desfaz para "199.906" quando houver apenas 1 ponto seguido de 3 dígitos e ",00":
    if s.endswith(",00") and s.count(".") == 1 and s.count(",") == 1:
        parte_ant = s[:-3]
        partes_ponto = parte_ant.split(".")
        if len(partes_ponto) == 2 and partes_ponto[0].isdigit() and len(partes_ponto[1]) == 3 and partes_ponto[1].isdigit():
            s = parte_ant
    # Agora trata '.' como ',' (separador decimal único)
    s_norm = s.replace(".", ",")
    if s_norm.count(",") == 1:
        inteiro, dec = s_norm.split(",")
        inteiro_limpo = "".join(c for c in inteiro if c.isdigit())
        dec_limpo = "".join(c for c in dec if c.isdigit())
        if inteiro_limpo:
            dec_2 = (dec_limpo + "00")[:2]
            return f"{int(inteiro_limpo)},{dec_2}"
    elif s_norm.isdigit():
        return f"{int(s_norm)},00"
    return s


def normalizar_codigo(codigo: str) -> str:
    """Remove letras no início do código (e espaço/hífen logo após), preservando meio/fim."""
    codigo = (codigo or "").strip()
    if not re.search(r"\d", codigo):
        return codigo
    return re.sub(r"^[A-Za-z]+[\s\-]*", "", codigo)


def _normalizar_item(item: ItemCorte) -> ItemCorte:
    """Garante formatação correta de valor_unitario ('199,90') e padrão '23' para local_destino."""
    val = _formatar_valor_unitario(item.valor_unitario)
    dest = (item.local_destino or "").strip() or LOCAL_DESTINO_PADRAO
    if val != item.valor_unitario or dest != item.local_destino:
        return item.model_copy(update={"valor_unitario": val, "local_destino": dest})
    return item


def consolidar_itens_extraidos(itens: list[ItemCorte]) -> list[ItemCorte]:
    """Limpa prefixos de letras nos códigos extraídos e consolida duplicatas em ordem."""
    grupos: dict[str, ItemCorte] = {}
    for item in itens:
        item_norm = _normalizar_item(item)
        codigo_limpo = normalizar_codigo(item_norm.codigo)
        item_limpo = (
            item_norm
            if codigo_limpo == item_norm.codigo
            else item_norm.model_copy(update={"codigo": codigo_limpo})
        )
        existente = grupos.get(codigo_limpo)
        if existente is None:
            grupos[codigo_limpo] = item_limpo
            continue

        obs_partes = [
            p.strip()
            for texto in (existente.observacao, item_limpo.observacao)
            for p in (texto or "").split(" | ")
            if p.strip()
        ]
        obs_unidas = " | ".join(dict.fromkeys(obs_partes))
        valor_unit = existente.valor_unitario
        if valor_unit == VALOR_UNITARIO_PADRAO and item_limpo.valor_unitario != VALOR_UNITARIO_PADRAO:
            valor_unit = item_limpo.valor_unitario

        grupos[codigo_limpo] = existente.model_copy(
            update={
                "quantidade_final": existente.quantidade_final + item_limpo.quantidade_final,
                "alterado_a_mao": existente.alterado_a_mao or item_limpo.alterado_a_mao,
                "observacao": obs_unidas,
                "valor_unitario": valor_unit,
            }
        )
    return list(grupos.values())


def imagem_para_base64(img: Image.Image) -> str:
    """Corrige rotação EXIF, reduz tamanho e converte para JPEG base64."""
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail((TAMANHO_MAX_IMAGEM, TAMANHO_MAX_IMAGEM), Image.LANCZOS)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=90)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def _erro_de_sobrecarga(erro: Exception) -> bool:
    """True se o erro indica servidor sobrecarregado/indisponível (tentar outro modelo)."""
    texto = f"{type(erro).__name__} {erro}".lower()
    return any(s in texto for s in ("503", "high demand", "unavailable", "overloaded",
                                    "internalservererror", "429", "resource_exhausted",
                                    "timeout", "timed out"))


def ler_folhas_com_gemini(imagens: list[Image.Image], aviso=None) -> list[ItemCorte]:
    """Envia todas as imagens em UMA requisição e devolve a lista de itens.

    `aviso` (opcional) recebe mensagens de status, ex.: troca para modelo reserva.
    """
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key or api_key == "sua_chave_aqui":
        raise RuntimeError(
            "GEMINI_API_KEY não configurada.\n\n"
            f"Edite o arquivo .env em:\n{PASTA_BASE}"
        )

    client = genai.Client(api_key=api_key)

    conteudo = []
    for img in imagens:
        conteudo.append({
            "type": "image",
            "data": imagem_para_base64(img),
            "mime_type": "image/jpeg",
        })
    conteudo.append({
        "type": "text",
        "text": (
            f"Leia as {len(imagens)} folha(s) de corte a laser enviadas e extraia "
            "todas as peças com seus campos completos para Entrada da Produção, "
            "consolidando códigos repetidos e priorizando anotações à caneta."
        ),
    })

    modelos = [MODELO_GEMINI] + [m for m in MODELOS_RESERVA if m != MODELO_GEMINI]
    ultimo_erro = None
    for n, modelo in enumerate(modelos):
        if n > 0 and aviso:
            aviso(f"⏳ {modelos[n - 1]} está sobrecarregado. Tentando {modelo}...")
        try:
            interaction = client.interactions.create(
                model=modelo,
                system_instruction=PROMPT_SISTEMA,
                input=conteudo,
                generation_config={"thinking_level": NIVEL_RACIOCINIO},
                response_format={
                    "type": "text",
                    "mime_type": "application/json",
                    "schema": ResultadoLeitura.model_json_schema(),
                },
                timeout=TEMPO_LIMITE_IA,
            )
        except Exception as e:
            if _erro_de_sobrecarga(e):
                ultimo_erro = e
                if n < len(modelos) - 1:
                    continue
                raise RuntimeError(
                    "Os servidores do Google Gemini estão sobrecarregados agora "
                    "(isso é temporário e não é problema do programa).\n\n"
                    "Aguarde alguns minutos e tente de novo.\n\n"
                    f"Detalhe técnico: {e}"
                ) from e
            raise
        texto = interaction.output_text
        if not texto:
            raise RuntimeError("A IA não retornou nenhum conteúdo. Tente novamente.")
        itens = ResultadoLeitura.model_validate_json(texto).itens
        return consolidar_itens_extraidos(itens)

    raise RuntimeError(f"Todos os modelos estão indisponíveis agora: {ultimo_erro}")


# =============================================================================
# INTERFACE GRÁFICA
# =============================================================================
class App:
    # Mapeamento das colunas editáveis direto na célula
    COLUNAS_EDITAVEIS = {
        "#1": "codigo",
        "#2": "valor_unitario",
        "#3": "local_destino",
        "#4": "quantidade_final",
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Leitor de Corte a Laser → WK Radar (Entrada da Produção)")
        self.root.geometry("1180x720")
        self.root.minsize(980, 560)

        self.itens: list[ItemCorte] = []
        self.ocupado = False
        self.calibrando = False
        self.digitando = False
        self.pausado = False
        self.parar_solicitado = False
        self._esc_anterior = False
        self._f12_anterior = False
        self.idx_pausado_popup: int | None = None

        cfg = carregar_config()
        self.pos_operacao: tuple[int, int] | None = cfg.get("pos_operacao")
        self.pos_produto: tuple[int, int] | None = cfg.get("pos_produto")
        self.hwnd_wk: int = 0

        # Variáveis do Painel de Calibração
        setas_ini = cfg.get("setas_operacao", SETAS_OPERACAO_PADRAO)
        espera_prod_ini = cfg.get("espera_apos_produto", ESPERA_APOS_PRODUTO)
        enters_ini = cfg.get("enters_apos_qtd", ENTERS_APOS_QTD_PADRAO)
        espera_ini = cfg.get("espera_apos_ok", ESPERA_APOS_OK)
        self.var_setas_operacao = tk.StringVar(value=str(setas_ini))
        self.var_espera_apos_produto = tk.StringVar(value=f"{float(espera_prod_ini):.1f}")
        self.var_enters_apos_qtd = tk.StringVar(value=str(enters_ini))
        self.var_espera_apos_ok = tk.StringVar(value=f"{float(espera_ini):.1f}")

        self._montar_interface()
        self._atualizar_indicador_calibracao()

        # Ctrl+V na janela principal (não afeta caixas de edição)
        self.root.bind("<Control-v>", self._atalho_colar)
        self.root.bind("<Control-V>", self._atalho_colar)

        # Monitor global de teclado (F12 = Pausar/Retomar, Esc = Parar)
        self._monitorar_atalhos_globais()

    # ------------------------------------------------------------------ layout
    def _montar_interface(self):
        estilo = ttk.Style()
        try:
            estilo.theme_use("vista")
        except tk.TclError:
            pass
        estilo.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
        estilo.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        estilo.configure("TButton", font=("Segoe UI", 10), padding=6)
        estilo.configure("Status.TLabel", font=("Segoe UI", 10, "bold"))
        estilo.configure("Instrucao.TLabel", font=("Segoe UI", 10, "bold"),
                         foreground="#8a4b00")

        # --- barra superior (arquivos e lista)
        topo = ttk.Frame(self.root, padding=(10, 10, 10, 4))
        topo.pack(fill="x")

        self.btn_abrir = ttk.Button(topo, text="📂 Abrir Foto(s) da Folha",
                                    command=self.abrir_fotos)
        self.btn_colar = ttk.Button(topo, text="📋 Colar Imagem (Ctrl+V)",
                                    command=self.colar_imagem)
        self.btn_limpar = ttk.Button(topo, text="🗑️ Limpar Lista",
                                     command=self.limpar_lista)
        self.btn_editar = ttk.Button(topo, text="✏️ Editar Item",
                                     command=self._editar_selecionada_dialogo)
        self.btn_desmarcar = ttk.Button(
            topo, text="🔄 Desmarcar Lançados (✅)",
            command=self.desmarcar_todos_digitados)

        self.btn_abrir.pack(side="left", padx=(0, 5))
        self.btn_colar.pack(side="left", padx=5)
        self.btn_limpar.pack(side="left", padx=5)
        self.btn_editar.pack(side="left", padx=5)
        self.btn_desmarcar.pack(side="left", padx=5)

        # --- dica de edição e atalhos
        ttk.Label(
            self.root,
            text=("Duplo clique na linha: editar campos (Código, Valor Unit., Destino, Qtd Final)  •  "
                  "Atalhos globais na digitação: F12 = Pausar/Retomar  |  Esc = Parar  |  "
                  "Mouse no canto sup. esq. = Abortar"),
            foreground="#005fb8", padding=(10, 2, 10, 0),
        ).pack(fill="x")

        # --- tabela central
        meio = ttk.Frame(self.root, padding=(10, 6, 10, 6))
        meio.pack(fill="both", expand=True)

        colunas = ("codigo", "valor_unit", "destino", "qtd", "manual", "obs", "status")
        self.tree = ttk.Treeview(meio, columns=colunas, show="headings",
                                 selectmode="extended")
        self.tree.heading("codigo", text="Código (Produto)")
        self.tree.heading("valor_unit", text="Valor Unit.")
        self.tree.heading("destino", text="Destino")
        self.tree.heading("qtd", text="Qtd Final")
        self.tree.heading("manual", text="Anotação Manual?")
        self.tree.heading("obs", text="Observação")
        self.tree.heading("status", text="Status")

        self.tree.column("codigo", width=200, anchor="w")
        self.tree.column("valor_unit", width=95, anchor="center")
        self.tree.column("destino", width=75, anchor="center")
        self.tree.column("qtd", width=85, anchor="center")
        self.tree.column("manual", width=130, anchor="center")
        self.tree.column("obs", width=355, anchor="w")
        self.tree.column("status", width=190, anchor="center")

        self.tree.tag_configure("alterado", background=COR_ALTERADO)
        self.tree.tag_configure("lancado", background=COR_LANCADO)
        self.tree.tag_configure("popup", background=COR_POPUP)

        scroll = ttk.Scrollbar(meio, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self._duplo_clique)
        self.tree.bind("<Delete>", self.remover_selecionados)
        self.tree.bind("<Return>", lambda _e: self.editar_qtd_selecionada())
        self.tree.bind("<KP_Enter>", lambda _e: self.editar_qtd_selecionada())
        self.tree.bind("<F2>", lambda _e: self._editar_selecionada_dialogo())
        self.tree.bind("<Key>", self._tecla_na_tabela)
        self.tree.bind("<Button-3>", self._menu_contexto)
        self.tree.bind("<MouseWheel>", lambda _e: self._fechar_editor(salvar=True), add="+")

        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="▶️ Digitar a partir desta linha",
                              command=self.iniciar_digitacao_da_selecionada)
        self.menu.add_command(label="✅ Marcar / desmarcar como lançado",
                              command=self.alternar_digitado_selecionados)
        self.menu.add_command(label="🔄 Desmarcar todos os lançados",
                              command=self.desmarcar_todos_digitados)
        self.menu.add_separator()
        self.menu.add_command(label="📝 Editar item completo (Código, Valor, Destino, Qtd)...",
                              command=self._editar_selecionada_dialogo)
        self.menu.add_command(label="✏️ Editar Quantidade Final",
                              command=self.editar_qtd_selecionada)
        self.menu.add_command(label="💲 Editar Valor Unitário",
                              command=lambda: self._editar_selecionada("#2"))
        self.menu.add_command(label="📦 Editar Local de Destino",
                              command=lambda: self._editar_selecionada("#3"))
        self.menu.add_command(label="🔤 Editar Código (Produto)",
                              command=lambda: self._editar_selecionada("#1"))
        self.menu.add_separator()
        self.menu.add_command(label="➕ Somar 1 na Qtd", command=lambda: self._ajustar_qtd(+1))
        self.menu.add_command(label="➖ Subtrair 1 na Qtd", command=lambda: self._ajustar_qtd(-1))
        self.menu.add_separator()
        self.menu.add_command(label="🗑️ Remover linha", command=self.remover_selecionados)

        self._editor = None  # caixinha de edição direta na célula

        # --- painel inferior: calibração + botões de iniciar/pausar/parar
        barra_robo = ttk.LabelFrame(
            self.root, text="Calibração e Digitação (Inclusão — WK Radar)",
            padding=(10, 6, 10, 8))
        barra_robo.pack(fill="x", padx=10, pady=(2, 4))

        linha_cal = ttk.Frame(barra_robo)
        linha_cal.pack(fill="x", pady=(0, 6))

        ttk.Label(linha_cal, text="Setas p/ baixo (Operação):").pack(side="left")
        self.spn_setas_operacao = ttk.Spinbox(
            linha_cal, from_=0, to=20, increment=1, width=5,
            textvariable=self.var_setas_operacao, font=("Segoe UI", 10))
        self.spn_setas_operacao.pack(side="left", padx=(4, 10))

        ttk.Label(linha_cal, text="Espera após Produto (s):").pack(side="left")
        self.spn_espera_apos_produto = ttk.Spinbox(
            linha_cal, from_=0.2, to=15.0, increment=0.1, width=6,
            textvariable=self.var_espera_apos_produto, font=("Segoe UI", 10))
        self.spn_espera_apos_produto.pack(side="left", padx=(4, 10))

        ttk.Label(linha_cal, text="Enters após Quantidade (até OK):").pack(side="left")
        self.spn_enters_apos_qtd = ttk.Spinbox(
            linha_cal, from_=1, to=50, increment=1, width=5,
            textvariable=self.var_enters_apos_qtd, font=("Segoe UI", 10))
        self.spn_enters_apos_qtd.pack(side="left", padx=(4, 10))

        ttk.Label(linha_cal, text="Espera após OK (s):").pack(side="left")
        self.spn_espera_apos_ok = ttk.Spinbox(
            linha_cal, from_=0.2, to=15.0, increment=0.1, width=6,
            textvariable=self.var_espera_apos_ok, font=("Segoe UI", 10))
        self.spn_espera_apos_ok.pack(side="left", padx=(4, 4))

        linha_acoes = ttk.Frame(barra_robo)
        linha_acoes.pack(fill="x")

        self.btn_pausar = ttk.Button(
            linha_acoes, text="⏸️ Pausar (F12)",
            command=lambda: self.alternar_pausa(via_atalho=False), state="disabled")
        self.btn_parar = ttk.Button(
            linha_acoes, text="⏹️ Parar (Esc)",
            command=self.parar_digitacao, state="disabled")
        self.btn_calibrar = ttk.Button(
            linha_acoes, text="🎯 Calibrar Posições",
            command=self.iniciar_calibracao,
        )
        self.lbl_status_calibracao = tk.Label(
            linha_acoes, text="⚠️ Não calibrado", fg="#d97706",
            font=("Segoe UI", 10, "bold"),
        )
        self.btn_pausar.pack(side="left", padx=(0, 6))
        self.btn_parar.pack(side="left", padx=6)
        self.btn_calibrar.pack(side="left", padx=(10, 6))
        self.lbl_status_calibracao.pack(side="left", padx=4)

        self.btn_digitar = ttk.Button(
            linha_acoes, text="▶ Digitar no WK Radar",
            command=self.iniciar_digitacao)
        self.btn_digitar_sel = ttk.Button(
            linha_acoes, text="▶️ Da Linha Selecionada",
            command=self.iniciar_digitacao_da_selecionada)
        self.btn_digitar.pack(side="right", padx=(6, 0))
        self.btn_digitar_sel.pack(side="right", padx=6)

        # --- rodapé com instrução fixa + status
        rodape = ttk.Frame(self.root, padding=(10, 2, 10, 10))
        rodape.pack(fill="x")

        faixa_instrucao = tk.Frame(rodape, bg="#fff3cd", bd=1, relief="solid",
                                   padx=8, pady=4)
        faixa_instrucao.pack(fill="x", pady=(0, 6))
        tk.Label(
            faixa_instrucao,
            text=f"📌 Instrução: {INSTRUCAO_RODAPE}",
            bg="#fff3cd", fg="#664d03", font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).pack(fill="x")

        linha_retomada = ttk.Frame(rodape)
        linha_retomada.pack(fill="x", pady=(0, 6))
        self.btn_continuar_proxima = ttk.Button(
            linha_retomada, text="▶ Continuar da Próxima Peça",
            command=self.continuar_proxima_peca, state="disabled")
        self.btn_repetir_pausada = ttk.Button(
            linha_retomada, text="🔁 Repetir Peça Pausada",
            command=self.repetir_peca_pausada, state="disabled")
        self.btn_continuar_proxima.pack(side="left", padx=(0, 8))
        self.btn_repetir_pausada.pack(side="left")

        linha_status = ttk.Frame(rodape)
        linha_status.pack(fill="x")
        self.var_status = tk.StringVar(value=INSTRUCAO_RODAPE)
        self.var_total = tk.StringVar(value="0 itens")
        ttk.Label(linha_status, textvariable=self.var_status,
                  style="Status.TLabel").pack(side="left")
        ttk.Label(linha_status, textvariable=self.var_total).pack(side="right")

    # ----------------------------------------------------------------- helpers
    def _set_status(self, texto: str):
        self.var_status.set(texto)

    def _atualizar_indicador_calibracao(self):
        if self.pos_operacao is not None and self.pos_produto is not None:
            self.lbl_status_calibracao.configure(text="✅ Calibrado", fg="#198754")
        else:
            self.lbl_status_calibracao.configure(text="⚠️ Não calibrado", fg="#d97706")

    def _salvar_config_atual(self):
        salvar_config(
            self.pos_operacao,
            self.pos_produto,
            self._ler_setas_operacao(),
            self._ler_enters_apos_qtd(),
            self._ler_espera_apos_ok(),
            espera_apos_produto=self._ler_espera_apos_produto(),
        )

    def _garantir_calibrado(self) -> bool:
        if self.pos_operacao is None or self.pos_produto is None:
            messagebox.showwarning(
                "Calibração necessária",
                "As posições ainda não foram calibradas.\n\n"
                "Clique primeiro em '🎯 Calibrar Posições' para salvar os campos Operação e Produto.",
            )
            return False
        return True

    def _restaurar_janela_app(self):
        try:
            self.root.deiconify()
            self.root.state("normal")
            self.root.lift()
            self.root.focus_force()
        except tk.TclError:
            pass

    def _obter_hwnd_app(self) -> int:
        try:
            wid = int(self.root.winfo_id())
            anc = int(ctypes.windll.user32.GetAncestor(wid, 2))
            return anc or wid
        except Exception:
            try:
                return int(self.root.winfo_id())
            except Exception:
                return 0

    def _definir_botoes_retomada(self, habilitar: bool):
        estado = "normal" if habilitar else "disabled"
        self.btn_continuar_proxima.configure(state=estado)
        self.btn_repetir_pausada.configure(state=estado)

    def _limpar_estado_popup(self):
        self.idx_pausado_popup = None
        self._definir_botoes_retomada(False)

    def _set_ocupado(self, ocupado: bool, modo: str = ""):
        self.ocupado = ocupado
        if ocupado:
            self._fechar_editor(salvar=True)
        estado = "disabled" if ocupado else "normal"
        for btn in (self.btn_abrir, self.btn_colar, self.btn_limpar,
                    self.btn_editar, self.btn_calibrar, self.btn_digitar,
                    self.btn_digitar_sel, self.btn_desmarcar):
            btn.configure(state=estado)

        em_digitacao = ocupado and modo == "digitacao"
        em_calibracao = ocupado and modo == "calibracao"
        self.btn_pausar.configure(
            state="normal" if em_digitacao else "disabled",
            text="▶️ Retomar (F12)" if (em_digitacao and self.pausado) else "⏸️ Pausar (F12)",
        )
        self.btn_parar.configure(
            state="normal" if (em_digitacao or em_calibracao) else "disabled"
        )

        pode_retomar_popup = (
            not ocupado
            and self.idx_pausado_popup is not None
            and 0 <= self.idx_pausado_popup < len(self.itens)
        )
        self._definir_botoes_retomada(pode_retomar_popup)

    def _ler_setas_operacao(self) -> int:
        try:
            valor = int(self.var_setas_operacao.get().strip())
            return max(0, min(valor, 50))
        except ValueError:
            self.var_setas_operacao.set(str(SETAS_OPERACAO_PADRAO))
            return SETAS_OPERACAO_PADRAO

    def _ler_enters_apos_qtd(self) -> int:
        try:
            valor = int(self.var_enters_apos_qtd.get().strip())
            return max(1, min(valor, 50))
        except ValueError:
            self.var_enters_apos_qtd.set(str(ENTERS_APOS_QTD_PADRAO))
            return ENTERS_APOS_QTD_PADRAO

    def _ler_espera_apos_produto(self) -> float:
        try:
            valor = float(self.var_espera_apos_produto.get().strip().replace(",", "."))
            return max(0.1, min(valor, 30.0))
        except ValueError:
            self.var_espera_apos_produto.set(f"{ESPERA_APOS_PRODUTO:.1f}")
            return ESPERA_APOS_PRODUTO

    def _ler_espera_apos_ok(self) -> float:
        try:
            valor = float(self.var_espera_apos_ok.get().strip().replace(",", "."))
            return max(0.1, min(valor, 30.0))
        except ValueError:
            self.var_espera_apos_ok.set(f"{ESPERA_APOS_OK:.1f}")
            return ESPERA_APOS_OK

    def _status_item_texto(self, item: ItemCorte) -> str:
        if item.erro_popup:
            return "⚠️ Falta de material / Pop-up"
        if item.lancado:
            return "✅ Resolvida manual" if item.resolvida_manual else "✅ Lançado"
        if item.quantidade_final <= 0:
            return "⚪ Ignorado (0)"
        return "⏳ Pendente"

    def _tags_item(self, item: ItemCorte) -> tuple[str, ...]:
        if item.erro_popup:
            return ("popup",)
        if item.lancado:
            return ("lancado",)
        if item.alterado_a_mao:
            return ("alterado",)
        return ()

    def _valores_linha(self, item: ItemCorte) -> tuple:
        return (
            item.codigo,
            item.valor_unitario,
            item.local_destino,
            item.quantidade_final,
            "✏️ Sim" if item.alterado_a_mao else "Não",
            item.observacao,
            self._status_item_texto(item),
        )

    def atualizar_tabela(self):
        self._fechar_editor(salvar=False)
        self.tree.delete(*self.tree.get_children())
        for i, item in enumerate(self.itens):
            self.tree.insert("", "end", iid=str(i), tags=self._tags_item(item),
                             values=self._valores_linha(item))
        self._atualizar_resumo()

    def _atualizar_resumo(self):
        n_alt = sum(1 for it in self.itens if it.alterado_a_mao)
        n_lanc = sum(1 for it in self.itens if it.lancado)
        soma_pecas = sum(max(0, it.quantidade_final) for it in self.itens)
        self.var_total.set(
            f"{len(self.itens)} itens ({soma_pecas} peças)  •  "
            f"{n_alt} manual  •  {n_lanc} lançados (✅)"
        )

    def _atualizar_linha_tabela(self, idx: int):
        item = self.itens[idx]
        iid = str(idx)
        if iid in self.tree.get_children():
            self.tree.item(iid, tags=self._tags_item(item),
                           values=self._valores_linha(item))
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
        self._atualizar_resumo()

    def _marcar_lancado(self, idx: int, valor: bool = True, resolvida_manual: bool = False):
        if 0 <= idx < len(self.itens):
            self.itens[idx] = self.itens[idx].model_copy(update={
                "lancado": valor,
                "erro_popup": False,
                "resolvida_manual": resolvida_manual if valor else False,
            })
            self._atualizar_linha_tabela(idx)

    def _marcar_popup_item(self, idx: int):
        if 0 <= idx < len(self.itens):
            self.itens[idx] = self.itens[idx].model_copy(update={
                "lancado": False,
                "erro_popup": True,
                "resolvida_manual": False,
            })
            self._atualizar_linha_tabela(idx)

    def alternar_digitado_selecionados(self):
        if self.ocupado:
            return
        sel = [int(i) for i in self.tree.selection()]
        if not sel:
            iid = self._selecionada()
            if iid is not None:
                sel = [int(iid)]
        if not sel:
            return
        novo_estado = not all(self.itens[idx].lancado for idx in sel)
        for idx in sel:
            self.itens[idx] = self.itens[idx].model_copy(update={
                "lancado": novo_estado,
                "erro_popup": False,
                "resolvida_manual": False,
            })
        self.atualizar_tabela()
        for idx in sel:
            self.tree.selection_add(str(idx))

    def desmarcar_todos_digitados(self):
        if self.ocupado or not self.itens:
            return
        self._limpar_estado_popup()
        if not any(it.lancado or it.erro_popup for it in self.itens):
            self._set_status("Nenhum item está marcado como lançado.")
            return
        self.itens = [
            it.model_copy(update={
                "lancado": False,
                "erro_popup": False,
                "resolvida_manual": False,
            })
            for it in self.itens
        ]
        self.atualizar_tabela()
        self._set_status("🔄 Marcações de lançado (✅) limpas.")

    # --------------------------------------------------------------- leitura
    def abrir_fotos(self):
        if self.ocupado:
            return
        caminhos = filedialog.askopenfilenames(
            title="Selecione uma ou mais fotos das folhas de corte",
            filetypes=[("Imagens", "*.jpg *.jpeg *.png *.bmp *.webp *.tif *.tiff"),
                       ("Todos os arquivos", "*.*")],
        )
        if not caminhos:
            return
        try:
            imagens = [Image.open(c) for c in caminhos]
            for img in imagens:
                img.load()
        except Exception as e:
            messagebox.showerror("Erro ao abrir imagem", str(e))
            return
        self._processar(imagens, acumular=False)

    def _atalho_colar(self, event=None):
        # Ctrl+V dentro de uma caixa de texto cola texto normalmente
        if event is not None and isinstance(
            event.widget, (tk.Entry, ttk.Entry, ttk.Spinbox, ttk.Combobox)
        ):
            return None
        self.colar_imagem()
        return "break"

    def colar_imagem(self):
        if self.ocupado:
            return
        try:
            conteudo = ImageGrab.grabclipboard()
        except Exception as e:
            messagebox.showerror("Erro ao ler a área de transferência", str(e))
            return

        imagens: list[Image.Image] = []
        if isinstance(conteudo, Image.Image):
            imagens = [conteudo]
        elif isinstance(conteudo, list):  # arquivos copiados no Explorer
            for caminho in conteudo:
                try:
                    img = Image.open(caminho)
                    img.load()
                    imagens.append(img)
                except Exception:
                    pass

        if not imagens:
            messagebox.showwarning(
                "Nada para colar",
                "Não há imagem na área de transferência.\n\n"
                "No WhatsApp Web: clique com o botão direito na foto → "
                "'Copiar imagem' e depois use Ctrl+V aqui.",
            )
            return
        self._processar(imagens, acumular=True)

    def _processar(self, imagens: list[Image.Image], acumular: bool):
        self._set_ocupado(True, modo="ia")
        self._set_status(f"⏳ Lendo {len(imagens)} imagem(ns) com a IA... aguarde.")

        def tarefa():
            try:
                novos = ler_folhas_com_gemini(
                    imagens, aviso=lambda m: self.root.after(0, self._set_status, m))
                self.root.after(0, self._leitura_ok, novos, acumular)
            except Exception as e:
                self.root.after(0, self._leitura_erro, e)

        threading.Thread(target=tarefa, daemon=True).start()

    def _leitura_ok(self, novos: list[ItemCorte], acumular: bool):
        if acumular:
            self.itens.extend(novos)
        else:
            self._limpar_estado_popup()
            self.itens = list(novos)
        self.atualizar_tabela()
        self._set_ocupado(False)
        modo = "adicionados ao final" if acumular else "carregados"
        self._set_status(f"✅ {len(novos)} itens {modo}. Confira as linhas amarelas!")

    def _leitura_erro(self, erro: Exception):
        self._set_ocupado(False)
        self._set_status("❌ Erro na leitura.")
        msg = str(erro)
        if "APIConnectionError" in type(erro).__name__ or "getaddrinfo" in msg:
            msg = ("Sem conexão com a internet (ou ela caiu durante a leitura).\n\n"
                   "Verifique a internet e tente de novo.\n\n"
                   f"Detalhe técnico: {erro}")
        messagebox.showerror("Erro na leitura da IA", msg)

    # ---------------------------------------------------------------- edição
    def limpar_lista(self):
        if self.ocupado or not self.itens:
            return
        if messagebox.askyesno("Limpar Lista", "Apagar todos os itens da lista?"):
            self.itens.clear()
            self._limpar_estado_popup()
            self.atualizar_tabela()
            self._set_status("Lista limpa. Pronto para um novo lote.")

    def remover_selecionados(self, _event=None):
        if self.ocupado:
            return
        selecionados = sorted((int(i) for i in self.tree.selection()), reverse=True)
        if not selecionados:
            return
        if self.idx_pausado_popup is not None:
            if self.idx_pausado_popup in selecionados:
                self._limpar_estado_popup()
            else:
                self.idx_pausado_popup -= sum(1 for i in selecionados if i < self.idx_pausado_popup)
        for idx in selecionados:
            del self.itens[idx]
        self.atualizar_tabela()
        self._set_ocupado(False)
        self._set_status(f"{len(selecionados)} linha(s) removida(s).")

    # ------------------------------------------------- edição direta e diálogo
    def _selecionada(self):
        """iid da linha em foco/selecionada (ou None)."""
        foco = self.tree.focus()
        if foco and foco in self.tree.selection():
            return foco
        sel = self.tree.selection()
        return sel[0] if sel else (foco or None)

    def _duplo_clique(self, event):
        if self.ocupado:
            return
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        coluna = self.tree.identify_column(event.x)
        if not iid:
            return
        if coluna == "#7":                  # Status: alterna ✅ Lançado
            idx = int(iid)
            self._marcar_lancado(idx, not self.itens[idx].lancado)
        else:
            # Abre a janela de edição completa focando o campo clicado
            self._abrir_dialogo(iid, coluna_foco=coluna)
        return "break"

    def editar_qtd_selecionada(self):
        self._editar_selecionada("#4")

    def _editar_selecionada(self, coluna: str):
        if self.ocupado:
            return
        iid = self._selecionada()
        if not iid:
            messagebox.showinfo("Editar", "Selecione uma linha da tabela primeiro.")
            return
        self._editar_celula(iid, coluna)

    def _editar_selecionada_dialogo(self):
        if self.ocupado:
            return
        iid = self._selecionada()
        if not iid:
            messagebox.showinfo("Editar", "Selecione uma linha da tabela primeiro.")
            return
        self._abrir_dialogo(iid)

    def _tecla_na_tabela(self, event):
        """Digitar um número com a linha selecionada já começa a editar a Qtd (#4)."""
        if self.ocupado:
            return None
        if event.char and event.char.isdigit():
            iid = self._selecionada()
            if iid:
                self._editar_celula(iid, "#4", texto_inicial=event.char)
                return "break"
        elif event.keysym in ("plus", "KP_Add"):
            self._ajustar_qtd(+1)
            return "break"
        elif event.keysym in ("minus", "KP_Subtract"):
            self._ajustar_qtd(-1)
            return "break"
        return None

    def _menu_contexto(self, event):
        if self.ocupado:
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if iid not in self.tree.selection():
            self.tree.selection_set(iid)
        self.tree.focus(iid)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def _registrar_correcao(self, idx: int, antigo: ItemCorte, novo: ItemCorte):
        """Atualiza o item e anota a correção manual na observação."""
        novo = _normalizar_item(novo)
        notas = []
        if novo.quantidade_final != antigo.quantidade_final:
            notas.append(f"qtd corrigida: {antigo.quantidade_final} → {novo.quantidade_final}")
        if novo.codigo != antigo.codigo:
            notas.append(f"código corrigido: {antigo.codigo} → {novo.codigo}")
        if novo.valor_unitario != antigo.valor_unitario:
            notas.append(f"valor: {antigo.valor_unitario} → {novo.valor_unitario}")
        if novo.local_destino != antigo.local_destino:
            notas.append(f"destino: {antigo.local_destino} → {novo.local_destino}")
        if notas:
            obs = novo.observacao
            obs = (obs + " | " if obs else "") + "; ".join(notas)
            novo = novo.model_copy(update={"observacao": obs})
        self.itens[idx] = novo
        self.atualizar_tabela()
        iid = str(idx)
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self.tree.see(iid)
        self.tree.focus_set()
        if notas:
            self._set_status(f"✏️ Linha {idx + 1}: " + "; ".join(notas))

    def _ajustar_qtd(self, delta: int):
        if self.ocupado:
            return
        iid = self._selecionada()
        if not iid:
            return
        idx = int(iid)
        antigo = self.itens[idx]
        nova = max(0, antigo.quantidade_final + delta)
        if nova != antigo.quantidade_final:
            self._registrar_correcao(
                idx, antigo, antigo.model_copy(update={"quantidade_final": nova}))

    def _editar_celula(self, iid: str, coluna: str, texto_inicial: str | None = None):
        """Abre uma caixinha de edição em cima da célula (#1, #2, #3 ou #4)."""
        self._fechar_editor(salvar=True)
        campo = self.COLUNAS_EDITAVEIS.get(coluna)
        if not campo or iid not in self.tree.get_children():
            return
        self.tree.see(iid)
        self.tree.update_idletasks()
        bbox = self.tree.bbox(iid, coluna)
        if not bbox:
            return
        x, y, w, h = bbox
        idx = int(iid)
        e_qtd = campo == "quantidade_final"
        item = self.itens[idx]

        if e_qtd:
            vcmd = (self.root.register(lambda p: p == "" or p.isdigit()), "%P")
            ent = tk.Entry(self.tree, font=("Segoe UI", 11, "bold"), justify="center",
                           relief="solid", bd=2, bg="#e8f4ff",
                           validate="key", validatecommand=vcmd)
        else:
            ent = tk.Entry(self.tree, font=("Segoe UI", 10), relief="solid", bd=2,
                           bg="#e8f4ff")

        valor = str(getattr(item, campo))
        if texto_inicial is not None:
            ent.insert(0, texto_inicial)
            ent.icursor("end")
        else:
            ent.insert(0, valor)
            ent.select_range(0, "end")
        ent.place(x=x, y=y, width=max(w, 60), height=h)
        ent.focus_set()
        self._editor = (ent, idx, campo)
        self._set_status("✏️ Digite o novo valor e aperte Enter (Esc cancela).")

        ent.bind("<Return>", lambda _e: (self._fechar_editor(salvar=True), "break")[1])
        ent.bind("<KP_Enter>", lambda _e: (self._fechar_editor(salvar=True), "break")[1])
        ent.bind("<Tab>", lambda _e: (self._fechar_editor(salvar=True), "break")[1])
        ent.bind("<Escape>", lambda _e: (self._fechar_editor(salvar=False),
                                         self.tree.focus_set(), "break")[2])

        def _ao_perder_foco(_e=None):
            if self._editor and self._editor[0] is ent and self.root.focus_get() is not ent:
                self._fechar_editor(salvar=True)

        ent.after(80, lambda: ent.winfo_exists() and ent.bind("<FocusOut>", _ao_perder_foco))

    def _fechar_editor(self, salvar: bool):
        if not self._editor:
            return
        ent, idx, campo = self._editor
        self._editor = None
        try:
            texto = ent.get().strip()
            ent.destroy()
        except tk.TclError:
            return
        if not salvar or idx >= len(self.itens):
            self._set_status("Edição cancelada.")
            return
        antigo = self.itens[idx]
        if campo == "quantidade_final":
            if not texto.isdigit():
                self.root.bell()
                self._set_status("Quantidade inválida — mantido o valor anterior.")
                return
            novo = antigo.model_copy(update={"quantidade_final": int(texto)})
        elif campo == "codigo":
            if not texto:
                self.root.bell()
                self._set_status("O código não pode ficar vazio — mantido o anterior.")
                return
            novo = antigo.model_copy(update={"codigo": texto})
        elif campo == "valor_unitario":
            novo = antigo.model_copy(
                update={"valor_unitario": texto or VALOR_UNITARIO_PADRAO})
        elif campo == "local_destino":
            novo = antigo.model_copy(
                update={"local_destino": texto or LOCAL_DESTINO_PADRAO})
        else:
            return
        if novo != antigo:
            self._registrar_correcao(idx, antigo, novo)
        else:
            self.tree.focus_set()

    # ------------------------------------------------- janela de edição completa
    def _abrir_dialogo(self, iid: str, coluna_foco: str = "#4"):
        self._fechar_editor(salvar=True)
        idx = int(iid)
        item = self.itens[idx]

        janela = tk.Toplevel(self.root)
        janela.title("Editar item — Entrada da Produção")
        janela.transient(self.root)
        janela.resizable(False, False)

        frm = ttk.Frame(janela, padding=15)
        frm.pack(fill="both", expand=True)

        var_cod = tk.StringVar(value=item.codigo)
        var_val = tk.StringVar(value=item.valor_unitario)
        var_dst = tk.StringVar(value=item.local_destino)
        var_qtd = tk.StringVar(value=str(item.quantidade_final))
        var_alt = tk.BooleanVar(value=item.alterado_a_mao)
        var_obs = tk.StringVar(value=item.observacao)
        var_dig = tk.BooleanVar(value=item.lancado)

        ttk.Label(frm, text="Código (Produto):").grid(row=0, column=0, sticky="w", pady=4)
        ent_cod = ttk.Entry(frm, textvariable=var_cod, width=42)
        ent_cod.grid(row=0, column=1, pady=4, sticky="w")

        ttk.Label(frm, text="Valor Unitário:").grid(row=1, column=0, sticky="w", pady=4)
        ent_val = ttk.Entry(frm, textvariable=var_val, width=16)
        ent_val.grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Local de Destino:").grid(row=2, column=0, sticky="w", pady=4)
        ent_dst = ttk.Entry(frm, textvariable=var_dst, width=16)
        ent_dst.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Quantidade Final:").grid(row=3, column=0, sticky="w", pady=4)
        ent_qtd = ttk.Entry(frm, textvariable=var_qtd, width=16)
        ent_qtd.grid(row=3, column=1, sticky="w", pady=4)

        ttk.Checkbutton(frm, text="Anotação manual (rasura à caneta)",
                        variable=var_alt).grid(row=4, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Observação:").grid(row=5, column=0, sticky="w", pady=4)
        ent_obs = ttk.Entry(frm, textvariable=var_obs, width=42)
        ent_obs.grid(row=5, column=1, pady=4, sticky="w")

        ttk.Checkbutton(frm, text="Já lançado no WK Radar (✅)",
                        variable=var_dig).grid(row=6, column=1, sticky="w", pady=4)

        def salvar(_e=None):
            codigo = var_cod.get().strip()
            if not codigo:
                messagebox.showwarning("Inválido", "O código não pode ficar vazio.",
                                       parent=janela)
                ent_cod.focus_set()
                return
            try:
                qtd = int(var_qtd.get().strip())
            except ValueError:
                messagebox.showwarning("Inválido", "A quantidade deve ser um número inteiro.",
                                       parent=janela)
                ent_qtd.focus_set()
                return
            valor_unit = var_val.get().strip() or VALOR_UNITARIO_PADRAO
            destino = var_dst.get().strip() or LOCAL_DESTINO_PADRAO
            marcado_lancado = var_dig.get()
            novo = ItemCorte(
                codigo=codigo,
                valor_unitario=valor_unit,
                local_destino=destino,
                quantidade_final=qtd,
                alterado_a_mao=var_alt.get(),
                observacao=var_obs.get().strip(),
                lancado=marcado_lancado,
                erro_popup=False,
                resolvida_manual=item.resolvida_manual if marcado_lancado else False,
            )
            janela.destroy()
            self._registrar_correcao(idx, item, novo)

        botoes = ttk.Frame(frm)
        botoes.grid(row=7, column=0, columnspan=2, pady=(12, 0), sticky="e")
        ttk.Button(botoes, text="Salvar", command=salvar).pack(side="right", padx=4)
        ttk.Button(botoes, text="Cancelar", command=janela.destroy).pack(side="right")

        janela.bind("<Return>", salvar)
        janela.bind("<Escape>", lambda _e: janela.destroy())

        # centraliza sobre a janela principal
        janela.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - janela.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - janela.winfo_height()) // 2
        janela.geometry(f"+{max(x, 0)}+{max(y, 0)}")

        try:
            janela.wait_visibility()
            janela.grab_set()
        except tk.TclError:
            pass

        mapa_foco = {
            "#1": ent_cod,
            "#2": ent_val,
            "#3": ent_dst,
            "#4": ent_qtd,
            "#6": ent_obs,
        }
        alvo = mapa_foco.get(coluna_foco, ent_qtd)
        alvo.focus_set()
        alvo.select_range(0, "end")

    # ------------------------------------ atalhos globais (F12 pausa / Esc para)
    def _monitorar_atalhos_globais(self):
        esc_atual = _tecla_pressionada_global(VK_ESCAPE)
        f12_atual = _tecla_pressionada_global(VK_F12)
        if (self.digitando or self.calibrando) and esc_atual and not self._esc_anterior:
            self.parar_digitacao()
        elif self.digitando and f12_atual and not self._f12_anterior:
            self.alternar_pausa(via_atalho=True)
        self._esc_anterior = esc_atual
        self._f12_anterior = f12_atual
        self.root.after(30, self._monitorar_atalhos_globais)

    def alternar_pausa(self, via_atalho: bool = False):
        if not self.digitando:
            return
        self.pausado = not self.pausado
        self.btn_pausar.configure(
            text="▶️ Retomar (F12)" if self.pausado else "⏸️ Pausar (F12)"
        )
        if self.pausado:
            self._set_status("⏸️ PAUSADO — aperte F12 para retomar ou Esc para parar.")
        else:
            if via_atalho:
                self._set_status("▶️ Retomando digitação...")
            else:
                self._set_status("▶️ Retomando digitação no WK Radar...")

    def parar_digitacao(self):
        if not (self.digitando or self.calibrando):
            return
        self.parar_solicitado = True
        self.pausado = False

    # ----------------------------------------------------------- calibração 2 pontos
    def iniciar_calibracao(self):
        """Executa a contagem de 2 etapas (5s Operação + 5s Produto) e salva em config.json."""
        if self.ocupado:
            return
        self._fechar_editor(salvar=True)
        self.calibrando = True
        self.pausado = False
        self.parar_solicitado = False
        self._esc_anterior = _tecla_pressionada_global(VK_ESCAPE)
        self._set_ocupado(True, modo="calibracao")
        self._contagem_calibracao(1, CONTAGEM_ETAPA_SEGUNDOS)

    def _contagem_calibracao(self, etapa: int, restante: int,
                             pos_operacao=None, hwnd_wk: int = 0):
        if not self.calibrando:
            return
        if self.parar_solicitado:
            self._fim_calibracao("🛑 Calibração cancelada pelo usuário (Esc).")
            return

        if etapa == 1:
            if restante > 0:
                self._set_status(
                    f"⏱️ 1/2: Deixe o mouse parado sobre o campo OPERAÇÃO ({restante}s)..."
                )
                self.root.after(1000, self._contagem_calibracao, 1, restante - 1, None, 0)
                return

            pos_operacao = _normalizar_ponto(pyautogui.position())
            hwnd_wk = obter_hwnd_ativo()
            try:
                winsound.Beep(800, 150)
            except Exception:
                pass
            self._contagem_calibracao(2, CONTAGEM_ETAPA_SEGUNDOS, pos_operacao, hwnd_wk)
            return

        if restante > 0:
            self._set_status(
                f"⏱️ 2/2: Agora deixe o mouse parado sobre a 1ª caixinha do campo PRODUTO ({restante}s)..."
            )
            self.root.after(1000, self._contagem_calibracao, 2, restante - 1, pos_operacao, hwnd_wk)
            return

        pos_produto = _normalizar_ponto(pyautogui.position())
        try:
            winsound.Beep(800, 150)
        except Exception:
            pass
        self.pos_operacao = pos_operacao
        self.pos_produto = pos_produto
        self.hwnd_wk = hwnd_wk
        self._salvar_config_atual()
        self._atualizar_indicador_calibracao()
        self._fim_calibracao(
            f"✅ Calibração salva em config.json! "
            f"(Operação {pos_operacao[0]},{pos_operacao[1]} • "
            f"Produto {pos_produto[0]},{pos_produto[1]})"
        )

    def _fim_calibracao(self, mensagem: str):
        self.calibrando = False
        self.pausado = False
        self.parar_solicitado = False
        self._set_ocupado(False)
        self._set_status(mensagem)
        self._restaurar_janela_app()

    # ------------------------------------------------------------ digitação
    def iniciar_digitacao(self):
        """Digita todos os itens pendentes (não lançados com ✅ e com qtd > 0)."""
        if self.ocupado:
            return
        self._fechar_editor(salvar=True)
        if not self._garantir_calibrado():
            return
        com_qtd = [(idx, it) for idx, it in enumerate(self.itens) if it.quantidade_final > 0]
        if not com_qtd:
            messagebox.showwarning("Nada para digitar",
                                   "Não há itens com quantidade maior que zero.")
            return

        pendentes = [(idx, it) for idx, it in com_qtd if not it.lancado]
        if not pendentes:
            if not messagebox.askyesno(
                "Tudo já lançado",
                "Todos os itens já estão marcados como lançados (✅).\n\n"
                "Deseja desmarcar todos e digitar tudo novamente?",
            ):
                return
            self.desmarcar_todos_digitados()
            pendentes = [(idx, it) for idx, it in enumerate(self.itens)
                         if it.quantidade_final > 0]

        self._iniciar_fila_digitacao(pendentes)

    def iniciar_digitacao_da_selecionada(self):
        """Digita a partir da linha selecionada até o final da lista."""
        if self.ocupado:
            return
        self._fechar_editor(salvar=True)
        if not self._garantir_calibrado():
            return
        iid = self._selecionada()
        if iid is None:
            messagebox.showinfo(
                "Digitar da Linha Selecionada",
                "Selecione na tabela a linha a partir da qual deseja começar a digitar.",
            )
            return
        idx_inicio = int(iid)
        fila = [(idx, self.itens[idx]) for idx in range(idx_inicio, len(self.itens))
                if self.itens[idx].quantidade_final > 0]
        if not fila:
            messagebox.showwarning(
                "Nada para digitar",
                "Não há itens com quantidade maior que zero a partir da linha selecionada.",
            )
            return
        self._iniciar_fila_digitacao(fila)

    def _iniciar_fila_digitacao(self, fila: list[tuple[int, ItemCorte]]):
        if not self._garantir_calibrado():
            return
        setas_operacao = self._ler_setas_operacao()
        espera_apos_produto = self._ler_espera_apos_produto()
        enters_apos_qtd = self._ler_enters_apos_qtd()
        espera_apos_ok = self._ler_espera_apos_ok()
        self._salvar_config_atual()

        self._limpar_estado_popup()
        self.digitando = True
        self.pausado = False
        self.parar_solicitado = False
        self._esc_anterior = _tecla_pressionada_global(VK_ESCAPE)
        self._f12_anterior = _tecla_pressionada_global(VK_F12)
        self._set_ocupado(True, modo="digitacao")
        copia_fila = [(idx, it.model_copy()) for idx, it in fila]
        self._set_status(
            f"⌨️ Digitando no WK Radar (Operação {self.pos_operacao[0]},{self.pos_operacao[1]} • "
            f"Produto {self.pos_produto[0]},{self.pos_produto[1]})... [F12 pausa • Esc para]"
        )
        threading.Thread(
            target=self._digitar,
            args=(copia_fila, self.pos_operacao, self.pos_produto, self.hwnd_wk,
                  setas_operacao, enters_apos_qtd, espera_apos_ok, espera_apos_produto),
            daemon=True,
        ).start()

    def _aguardar_pausa_ou_parada(self) -> bool:
        """Retorna True se deve parar imediatamente."""
        while self.pausado and not self.parar_solicitado:
            time.sleep(0.04)
        return self.parar_solicitado

    def _dormir_interrompivel(self, segundos: float) -> bool:
        """Dorme `segundos` respeitando pausa (F12) e parada (Esc). True se parou."""
        restante = segundos
        passo = 0.03
        while restante > 0:
            if self._aguardar_pausa_ou_parada():
                return True
            t = min(passo, restante)
            time.sleep(t)
            restante -= t
        return self._aguardar_pausa_ou_parada()

    def _digitar(self, fila: list[tuple[int, ItemCorte]],
                 pos_operacao, pos_produto, hwnd_wk: int,
                 setas_operacao: int, enters_apos_qtd: int, espera_apos_ok: float,
                 espera_apos_produto: float = ESPERA_APOS_PRODUTO):
        total = len(fila)
        digitados_agora = 0

        def checar_popup(idx_atual: int, codigo_atual: str) -> bool:
            if verificar_popup(hwnd_wk):
                titulo_popup = obter_titulo_janela(obter_hwnd_ativo())
                try:
                    winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
                except Exception:
                    pass
                self.root.after(0, self._pausar_por_popup, idx_atual, codigo_atual, titulo_popup)
                return True
            return False

        try:
            # Minimiza o aplicativo e garante o foco no WK Radar antes de iniciar
            self.root.after(0, self.root.iconify)
            if self._dormir_interrompivel(0.5):
                self.root.after(
                    0, self._fim_digitacao,
                    "🛑 Digitação cancelada pelo usuário (Esc)."
                )
                return

            if janela_valida(hwnd_wk):
                focar_janela(hwnd_wk)
                if self._dormir_interrompivel(0.25):
                    self.root.after(
                        0, self._fim_digitacao,
                        "🛑 Digitação cancelada pelo usuário (Esc)."
                    )
                    return
            else:
                hwnd_wk = obter_hwnd_ativo()
                if not hwnd_wk or hwnd_wk == self._obter_hwnd_app():
                    pyautogui.click(pos_operacao)
                    if self._dormir_interrompivel(0.2):
                        self.root.after(
                            0, self._fim_digitacao,
                            "🛑 Digitação cancelada pelo usuário (Esc)."
                        )
                        return
                    hwnd_wk = obter_hwnd_ativo()
                self.hwnd_wk = hwnd_wk

            for pos, (idx, item) in enumerate(fila, start=1):
                if item.quantidade_final <= 0:
                    continue
                if self._aguardar_pausa_ou_parada():
                    self.root.after(
                        0, self._fim_digitacao,
                        f"🛑 Digitação parada pelo usuário ({digitados_agora}/{total} lançados)."
                    )
                    return

                self.root.after(
                    0, self._set_status,
                    f"⌨️ Digitando {pos}/{total} (linha {idx + 1}): {item.codigo} "
                    f"(qtd {item.quantidade_final})  [F12 pausa • Esc para]"
                )

                # 1. Selecionar "Entrada da Produção" com reset pela tecla Home
                if checar_popup(idx, item.codigo):
                    return
                pyautogui.click(pos_operacao)
                if self._dormir_interrompivel(0.25):
                    break

                if checar_popup(idx, item.codigo):
                    return
                pyautogui.press("home")
                if self._dormir_interrompivel(0.15):
                    break

                parou_nas_setas = False
                for s in range(setas_operacao):
                    if checar_popup(idx, item.codigo):
                        return
                    pyautogui.press("down")
                    if s < setas_operacao - 1:
                        if self._dormir_interrompivel(0.1):
                            parou_nas_setas = True
                            break
                if parou_nas_setas:
                    break
                if setas_operacao > 0 and self._dormir_interrompivel(0.15):
                    break

                if checar_popup(idx, item.codigo):
                    return
                pyautogui.press("enter")
                if self._dormir_interrompivel(0.35):
                    break

                # 2. Pular direto para o campo Produto (ignorando campos dinâmicos no meio)
                if checar_popup(idx, item.codigo):
                    return
                pyautogui.click(pos_produto)
                if self._dormir_interrompivel(0.35):
                    break

                if checar_popup(idx, item.codigo):
                    return
                digitar_campo_seguro(item.codigo, limpar=True)
                if self._aguardar_pausa_ou_parada():
                    break

                # 3. Fluxo por Enter a partir de Produto
                # Enter 1: sai de Produto / passa por Tipo de Estoque (consulta no banco)
                if checar_popup(idx, item.codigo):
                    return
                pyautogui.press("enter")
                if self._dormir_interrompivel(espera_apos_produto):
                    break
                if checar_popup(idx, item.codigo):
                    return

                # Enter 2: vai para Valor Unitário
                if checar_popup(idx, item.codigo):
                    return
                pyautogui.press("enter")
                if self._dormir_interrompivel(0.25):
                    break

                if checar_popup(idx, item.codigo):
                    return
                valor_unit = (item.valor_unitario or "").strip() or VALOR_UNITARIO_PADRAO
                digitar_campo_seguro(valor_unit, limpar=True)
                if self._aguardar_pausa_ou_parada():
                    break

                # Enter 3: vai para Local de Destino (aguarda 1.5s, limpa e digita destino)
                if checar_popup(idx, item.codigo):
                    return
                pyautogui.press("enter")
                if self._dormir_interrompivel(1.5):
                    break

                if checar_popup(idx, item.codigo):
                    return
                destino_val = (item.local_destino or "").strip() or LOCAL_DESTINO_PADRAO
                digitar_campo_seguro(destino_val, limpar=True)
                if self._aguardar_pausa_ou_parada():
                    break

                # Enter 4: sai de Local de Destino / campo intermediário
                if checar_popup(idx, item.codigo):
                    return
                pyautogui.press("enter")
                if self._dormir_interrompivel(0.25):
                    break

                # Enter 5: vai para Quantidade
                if checar_popup(idx, item.codigo):
                    return
                pyautogui.press("enter")
                if self._dormir_interrompivel(0.25):
                    break

                if checar_popup(idx, item.codigo):
                    return
                digitar_campo_seguro(str(item.quantidade_final), limpar=True)
                if self._aguardar_pausa_ou_parada():
                    break

                # Enters finais até o OK
                parou_nos_enters = False
                for _ in range(enters_apos_qtd):
                    if checar_popup(idx, item.codigo):
                        return
                    pyautogui.press("enter")
                    if self._dormir_interrompivel(0.12):
                        parou_nos_enters = True
                        break
                if parou_nos_enters:
                    break

                # 15. Monitoramento ativo pós-OK (checa pop-up a cada 0.1s durante espera_apos_ok)
                restante_ok = espera_apos_ok
                while restante_ok > 0:
                    if self._aguardar_pausa_ou_parada():
                        break
                    if checar_popup(idx, item.codigo):
                        return
                    passo_ok = min(0.1, restante_ok)
                    time.sleep(passo_ok)
                    restante_ok -= passo_ok
                if self.parar_solicitado:
                    break
                if checar_popup(idx, item.codigo):
                    return

                digitados_agora += 1
                self.root.after(0, self._marcar_lancado, idx, True, False)

            if self.parar_solicitado:
                self.root.after(
                    0, self._fim_digitacao,
                    f"🛑 Digitação parada pelo usuário ({digitados_agora}/{total} lançados)."
                )
                return
        except pyautogui.FailSafeException:
            self.root.after(
                0, self._fim_digitacao,
                f"🛑 Digitação abortada pelo mouse no canto ({digitados_agora}/{total} lançados)."
            )
            return
        except Exception as e:
            self.root.after(0, self._fim_digitacao, f"❌ Erro na digitação: {e}")
            return
        self.root.after(
            0, self._fim_digitacao,
            f"✅ Concluído! {digitados_agora} itens lançados no WK Radar."
        )

    def _pausar_por_popup(self, idx: int, codigo: str, titulo_popup: str):
        self.digitando = False
        self.pausado = False
        self.parar_solicitado = False
        self.idx_pausado_popup = idx
        self._marcar_popup_item(idx)
        self._set_ocupado(False)
        self._set_status(
            f"⚠️ PAUSADO POR POP-UP na peça '{codigo}' — Janela: {titulo_popup}"
        )
        self._restaurar_janela_app()

    def continuar_proxima_peca(self):
        """Marca a peça pausada como resolvida manualmente e retoma pela próxima."""
        if self.ocupado or self.idx_pausado_popup is None:
            return
        if not self._garantir_calibrado():
            return
        if not (0 <= self.idx_pausado_popup < len(self.itens)):
            self._limpar_estado_popup()
            return
        idx_pausado = self.idx_pausado_popup
        self._marcar_lancado(idx_pausado, True, resolvida_manual=True)
        self._limpar_estado_popup()
        fila = [
            (i, self.itens[i])
            for i in range(idx_pausado + 1, len(self.itens))
            if self.itens[i].quantidade_final > 0 and not self.itens[i].lancado
        ]
        if not fila:
            self._set_status("✅ Peça marcada como resolvida manual. Não há mais peças pendentes.")
            return
        self._iniciar_fila_digitacao(fila)

    def repetir_peca_pausada(self):
        """Retoma a digitação começando novamente pela peça onde ocorreu o pop-up."""
        if self.ocupado or self.idx_pausado_popup is None:
            return
        if not self._garantir_calibrado():
            return
        if not (0 <= self.idx_pausado_popup < len(self.itens)):
            self._limpar_estado_popup()
            return
        idx_pausado = self.idx_pausado_popup
        self._marcar_lancado(idx_pausado, False, resolvida_manual=False)
        self._limpar_estado_popup()
        fila = [
            (i, self.itens[i])
            for i in range(idx_pausado, len(self.itens))
            if self.itens[i].quantidade_final > 0 and not self.itens[i].lancado
        ]
        if not fila:
            self._set_status("Não há peças pendentes a partir da linha pausada.")
            return
        self._iniciar_fila_digitacao(fila)

    def _fim_digitacao(self, mensagem: str):
        self.digitando = False
        self.pausado = False
        self.parar_solicitado = False
        self._set_ocupado(False)
        self._set_status(mensagem)
        self._restaurar_janela_app()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
