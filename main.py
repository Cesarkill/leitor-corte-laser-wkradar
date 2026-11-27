# -*- coding: utf-8 -*-
"""
Leitor de Folhas de Corte a Laser -> WK Radar (Entrada da Produção)
-------------------------------------------------------------------
Lê fotos de folhas de controle de corte a laser (com anotações/rasuras à caneta)
usando a API do Google Gemini e digita automaticamente os dados na tela de
Entrada da Produção do ERP WK Radar.
"""

import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
import ctypes
import ctypes.wintypes
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import io
import json
import os
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import uuid
import winsound

import pdfplumber
try:
    import pypdfium2
except ImportError:
    pypdfium2 = None
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
FATOR_DOBRA_PADRAO = Decimal("1.2")  # acréscimo R$/kg quando Dobra == "SIM"
ESPERA_APOS_PRODUTO = 1.0       # tempo (s) após Enter do Produto p/ consulta no banco
ENTERS_APOS_QTD_PADRAO = 18     # Enters finais até OK + confirmação do OK
ESPERA_APOS_OK = 1.5            # tempo (s) para salvar e limpar o formulário
VALOR_UNITARIO_PADRAO = "0,00"
LOCAL_DESTINO_PADRAO = "23"

COR_DOBRA_SIM = "#e7f3ff"       # fundo azul claro suave para itens com Dobra == SIM
COR_ALTERADO = "#fff3cd"        # fundo amarelo claro para anotações manuais
COR_LANCADO = "#d4edda"         # fundo verde claro das linhas já lançadas (✅)
COR_POPUP = "#f8d7da"           # fundo vermelho claro quando abre pop-up no WK Radar
COR_NAO_CADASTRADO = "#ffe5b4"  # fundo laranja claro quando produto não cadastrado
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
class ItemBrutoGemini(BaseModel):
    codigo: str = Field(
        description="Código da peça na coluna 'Referência' (sem letras iniciais de "
                    "prefixo antes dos números, mantendo pontos, hífens e letras finais)."
    )
    dobra: str = Field(
        default="NÃO",
        description="Conteúdo da coluna 'Dobra': retorne 'SIM' apenas quando explícito; "
                    "se vier vazio, em branco ou NÃO, retorne 'NÃO'."
    )
    preco_base: str = Field(
        default="0,00",
        description="Preço bruto da coluna 'Preço' sem o 'R$' (ex.: '43.77', '16.754'). "
                    "NÃO faça cálculos nem some peso."
    )
    peso: str = Field(
        default="0",
        description="Peso bruto da coluna 'Peso / KG' (ex.: '3.92', '0.71'). Use '0' se vazio."
    )
    quantidade: int = Field(
        description="Quantidade inteira da peça ('Qtd Peças' ou fallback 'Pedido'). "
                    "Se houver número corrigido à caneta, priorize-o."
    )


class ResultadoLeituraBruta(BaseModel):
    itens: list[ItemBrutoGemini] = Field(
        description="Cada linha válida da folha como objeto separado, sem agrupar."
    )


class ItemCorte(BaseModel):
    uid: SkipJsonSchema[str] = Field(
        default_factory=lambda: uuid.uuid4().hex, exclude=True
    )
    codigo: str = Field(
        description="Código da peça impresso na folha (sem letras de prefixo no início "
                    "antes do número, mantendo pontos, hífens, barras, espaços internos e "
                    "letras no meio ou no final do código)."
    )
    dobra: str = Field(default="NÃO")
    preco_base: Decimal = Field(default=Decimal("0"))
    peso: str = Field(default="0")
    valor_unitario: str = Field(
        default=VALOR_UNITARIO_PADRAO,
        description="Valor unitário final (com acréscimo de dobra quando aplicável)."
    )
    valor_manual: SkipJsonSchema[bool] = Field(default=False, exclude=True)
    local_destino: str = Field(
        default=LOCAL_DESTINO_PADRAO,
        description="Local de destino no estoque. Use sempre o valor padrão '23'."
    )
    quantidade_final: int = Field(
        description="Quantidade final da peça."
    )
    alterado_a_mao: bool = Field(default=False)
    observacao: str = Field(default="")
    lancado: SkipJsonSchema[bool] = Field(default=False, exclude=True)
    erro_popup: SkipJsonSchema[bool] = Field(default=False, exclude=True)
    nao_cadastrado: SkipJsonSchema[bool] = Field(default=False, exclude=True)
    tipo_popup: SkipJsonSchema[str] = Field(default="", exclude=True)
    resolvida_manual: SkipJsonSchema[bool] = Field(default=False, exclude=True)

    @field_validator("preco_base", mode="before")
    @classmethod
    def _validar_preco_base(cls, v):
        return _parse_decimal_seguro(v, padrao=Decimal("0")) or Decimal("0")

    @field_validator("valor_unitario", mode="before")
    @classmethod
    def _fallback_valor_unitario(cls, v):
        if v is None or (isinstance(v, str) and not v.strip()):
            return VALOR_UNITARIO_PADRAO
        return str(v)


class ResultadoLeitura(BaseModel):
    itens: list[ItemCorte] = Field(description="Lista de peças lidas nas folhas.")


PROMPT_SISTEMA = """\
Você é um assistente especialista em extrair linhas de relatórios e fotos de folhas de \
corte a laser (padrão HMAGRO / CypCut). Extraia APENAS os dados brutos em JSON, sem fazer \
nenhum cálculo matemático.

REGRAS OBRIGATÓRIAS:
1. FILTRO DE TESTE: IGNORE e NÃO inclua na lista nenhuma peça ou linha cuja referência \
contenha a palavra "TESTE" (ex.: "3.28.053 TESTE_0225", ignorando maiúsculas/minúsculas).
2. CÓDIGO ("codigo"): fica na coluna "Referência" (ex.: "3.11.334", "L3.28.051", \
"3.12.084D"). Ignore letras no início antes do código numérico (prefixos), retornando \
sem essas letras iniciais (ex.: "L3.28.051" -> "3.28.051"), mas mantenha SEMPRE letras no \
meio ou no final (como "D" em "3.12.084D"), além de pontos, hífens e barras. NÃO use o \
número da coluna "#".
3. DOBRA ("dobra"): fica na coluna "Dobra". Se a célula estiver vazia, em branco ou "NÃO", \
retorne exatamente "NÃO". Só retorne "SIM" se estiver explicitamente "SIM".
4. PREÇO BASE ("preco_base") e PESO ("peso"):
   - Extraia os valores brutos das colunas "Preço" (sem "R$", ex.: "43.77", "16.754") e \
"Peso / KG" (ex.: "3.92"). NÃO faça cálculos — o programa calculará o acréscimo em Python.
   - Se não houver preço, use "0,00". Se não houver peso, use "0".
5. QUANTIDADE ("quantidade"): número inteiro da coluna "Qtd Peças" (ou da coluna \
"Pedido" se "Qtd Peças" estiver vazia). Se houver correção riscada/escrita à caneta, \
priorize o número escrito à caneta.
6. SEM AGRUPAMENTO: NÃO agrupe nem some itens com códigos iguais; retorne cada linha do \
documento como um objeto separado na lista, na ordem exata em que aparece.
7. Não invente peças. Ignore cabeçalhos, totais, resumos da chapa e linhas em branco.
"""

PROMPT_CLASSIFICAR_POPUP = (
    "Analise esta caixa de aviso de um sistema ERP. Responda APENAS com uma destas "
    "3 palavras exatas, sem pontuação nem explicação: NAO_CADASTRADO (se disser que "
    "o produto/código não existe ou não está cadastrado), FALTA_MATERIAL (se disser "
    "falta de material, saldo ou estoque insuficiente) ou OUTRO (para qualquer outro aviso)."
)


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


def _ler_texto_hwnd(hwnd: int) -> str:
    """Lê texto visível (título ou controle filho) de um HWND via GetWindowTextW."""
    if not hwnd:
        return ""
    try:
        user32 = ctypes.windll.user32
        tamanho = user32.GetWindowTextLengthW(int(hwnd))
        if tamanho <= 0:
            return ""
        buffer = ctypes.create_unicode_buffer(tamanho + 1)
        user32.GetWindowTextW(int(hwnd), buffer, tamanho + 1)
        return buffer.value.strip()
    except Exception:
        return ""


def obter_texto_janela_completo(hwnd_popup: int) -> str:
    """Coleta título e textos de todos os controles filhos do pop-up em minúsculas."""
    if not hwnd_popup:
        return ""
    partes: list[str] = []
    titulo = _ler_texto_hwnd(hwnd_popup)
    if titulo:
        partes.append(titulo)
    try:
        user32 = ctypes.windll.user32
        enum_proc = ctypes.WINFUNCTYPE(
            ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
        )

        def _visitar_filho(child_hwnd, _lparam):
            txt = _ler_texto_hwnd(int(child_hwnd))
            if txt:
                partes.append(txt)
            return True

        cb = enum_proc(_visitar_filho)
        user32.EnumChildWindows(int(hwnd_popup), cb, 0)
    except Exception:
        pass
    return " ".join(partes).lower()


def classificar_texto_popup(texto: str) -> str | None:
    """Classifica o texto extraído do pop-up pelas palavras-chave conhecidas."""
    txt = (texto or "").lower()
    if not txt:
        return None
    if any(
        t in txt
        for t in ("cadastr", "inexist", "não encontrad", "nao encontrad", "inválid", "invalid")
    ):
        return "NAO_CADASTRADO"
    if any(t in txt for t in ("material", "estoque", "saldo", "insuficiente")):
        return "FALTA_MATERIAL"
    return None


def _capturar_popup_em_ram(hwnd_popup: int) -> Image.Image:
    """Captura o recorte do pop-up (com margem de 40px) apenas em memória RAM."""
    try:
        rect = ctypes.wintypes.RECT()
        ok = bool(hwnd_popup) and bool(
            ctypes.windll.user32.GetWindowRect(int(hwnd_popup), ctypes.byref(rect))
        )
        if ok:
            largura = rect.right - rect.left
            altura = rect.bottom - rect.top
            if largura > 50 and altura > 50:
                x1 = max(0, rect.left - 40) if rect.left >= 0 else rect.left - 40
                y1 = max(0, rect.top - 40) if rect.top >= 0 else rect.top - 40
                x2 = rect.right + 40
                y2 = rect.bottom + 40
                bbox = (x1, y1, x2, y2)
                if x1 < 0 or y1 < 0:
                    return ImageGrab.grab(bbox=bbox, all_screens=True)
                return ImageGrab.grab(bbox=bbox)
    except Exception:
        pass
    return ImageGrab.grab()


def _classificar_popup_com_gemini(img: Image.Image) -> str:
    """Classifica o print do pop-up em RAM via Gemini; retorna 'OUTRO' em falhas."""
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key or api_key == "sua_chave_aqui":
        return "OUTRO"
    try:
        client = genai.Client(api_key=api_key)
        conteudo = [
            {
                "type": "image",
                "data": imagem_para_base64(img),
                "mime_type": "image/jpeg",
            },
            {
                "type": "text",
                "text": PROMPT_CLASSIFICAR_POPUP,
            },
        ]
        modelos = [MODELO_GEMINI] + [m for m in MODELOS_RESERVA if m != MODELO_GEMINI]
        for n, modelo in enumerate(modelos):
            try:
                interaction = client.interactions.create(
                    model=modelo,
                    input=conteudo,
                    generation_config={"thinking_level": "minimal"},
                    timeout=12,
                )
                bruto = (interaction.output_text or "").strip().upper()
                palavra = re.sub(r"[^A-Z_]", "", bruto)
                if palavra in ("NAO_CADASTRADO", "FALTA_MATERIAL", "OUTRO"):
                    return palavra
                return "OUTRO"
            except Exception as e:
                if _erro_de_sobrecarga(e) and n < len(modelos) - 1:
                    continue
                return "OUTRO"
    except Exception:
        return "OUTRO"
    return "OUTRO"


def classificar_popup(hwnd_popup: int) -> str:
    """Classifica pop-up em 2 camadas: texto ctypes instantâneo e print em RAM + Gemini."""
    texto = obter_texto_janela_completo(hwnd_popup)
    tipo = classificar_texto_popup(texto)
    if tipo in ("NAO_CADASTRADO", "FALTA_MATERIAL"):
        return tipo
    try:
        img = _capturar_popup_em_ram(hwnd_popup)
        return _classificar_popup_com_gemini(img)
    except Exception:
        return "OUTRO"


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

    if "fator_dobra" in dados:
        try:
            val_dobra = Decimal(str(dados["fator_dobra"]).strip().replace(",", "."))
            if val_dobra.is_finite():
                val_dobra = max(Decimal("0"), min(val_dobra, Decimal("100")))
                cfg["fator_dobra"] = val_dobra
        except (InvalidOperation, TypeError, ValueError):
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
                  fator_dobra: Decimal | float | str = FATOR_DOBRA_PADRAO,
                  caminho: str | None = None) -> None:
    """Salva calibração de posições e parâmetros do robô em config.json."""
    caminho_alvo = caminho or CAMINHO_CONFIG
    op = _normalizar_ponto(pos_operacao)
    prod = _normalizar_ponto(pos_produto)
    try:
        dobra_dec = Decimal(str(fator_dobra).strip().replace(",", "."))
        if not dobra_dec.is_finite():
            dobra_dec = FATOR_DOBRA_PADRAO
    except (InvalidOperation, TypeError, ValueError):
        dobra_dec = FATOR_DOBRA_PADRAO
    dados = {
        "pos_operacao": [op[0], op[1]] if op is not None else None,
        "pos_produto": [prod[0], prod[1]] if prod is not None else None,
        "setas_operacao": int(setas_operacao),
        "fator_dobra": float(round(dobra_dec, 4)),
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


def limpar_campo_produto_apos_popup() -> None:
    """Apaga imediatamente o código inválido no campo Produto sem clicar fora."""
    _enviar_tecla_win(VK_BACK)
    time.sleep(0.08)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.05)
    _enviar_tecla_win(VK_BACK)
    time.sleep(0.05)
    _enviar_tecla_win(VK_END, estendida=True)
    time.sleep(0.05)
    _enviar_shift_home_win()
    time.sleep(0.05)
    _enviar_tecla_win(VK_BACK)
    time.sleep(0.15)


def normalizar_dobra(valor) -> str:
    """Retorna 'SIM' apenas se explícito; vazio, nulo ou qualquer outro texto vira 'NÃO'."""
    return "SIM" if str(valor or "").strip().upper() == "SIM" else "NÃO"


def _parse_decimal_seguro(valor, padrao: Decimal | None = Decimal("0")) -> Decimal | None:
    """Converte strings de preço/peso ('R$ 16.754', '43,77', '1.234,56') em Decimal exato."""
    if isinstance(valor, Decimal):
        return valor if valor.is_finite() else padrao
    if isinstance(valor, (int, float)):
        try:
            d = Decimal(str(valor))
            return d if d.is_finite() else padrao
        except InvalidOperation:
            return padrao

    s = (str(valor) if valor is not None else "").strip()
    if not s:
        return padrao
    s = re.sub(r"(?i)r\$|kg", "", s).strip()
    if not s:
        return padrao

    # Trata formato legado '199.906,00' (ponto de 3 casas decimais + ',00')
    if s.endswith(",00") and s.count(".") == 1 and s.count(",") == 1:
        parte_ant = s[:-3]
        partes_ponto = parte_ant.split(".")
        if (
            len(partes_ponto) == 2
            and partes_ponto[0].isdigit()
            and len(partes_ponto[1]) == 3
            and partes_ponto[1].isdigit()
        ):
            s = parte_ant

    s = s.replace(" ", "")
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if s.count(",") == 1:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(".") > 1:
        partes = s.split(".")
        if all(len(p) == 3 and p.isdigit() for p in partes[1:]):
            s = "".join(partes)

    try:
        d = Decimal(s)
        return d if d.is_finite() else padrao
    except InvalidOperation:
        return padrao


def _formatar_peso(valor) -> str:
    """Formata o peso em kg com vírgula e 2 casas decimais (ex.: '3,92')."""
    d = _parse_decimal_seguro(valor, padrao=Decimal("0"))
    if d is None:
        return "0,00"
    return f"{d.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)}".replace(".", ",")


def calcular_valor_unitario(preco_base, peso, dobra,
                            fator_dobra=FATOR_DOBRA_PADRAO) -> str:
    """Calcula valor final = preco_base + (peso * fator_dobra) se Dobra == 'SIM'."""
    p_base = _parse_decimal_seguro(preco_base, padrao=None)
    if p_base is None:
        return VALOR_UNITARIO_PADRAO
    p_peso = _parse_decimal_seguro(peso, padrao=Decimal("0")) or Decimal("0")
    fator = _parse_decimal_seguro(fator_dobra, padrao=FATOR_DOBRA_PADRAO)
    if fator is None:
        fator = FATOR_DOBRA_PADRAO

    if normalizar_dobra(dobra) == "SIM":
        valor_final = p_base + (p_peso * fator)
    else:
        valor_final = p_base

    return f"{valor_final.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)}".replace(".", ",")


def _formatar_valor_unitario(valor) -> str:
    """Formata um valor numérico para 2 casas decimais com vírgula."""
    d = _parse_decimal_seguro(valor, padrao=None)
    if d is None:
        return VALOR_UNITARIO_PADRAO
    return f"{d.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)}".replace(".", ",")


def _formatar_preco_base(valor) -> str:
    """Formata preco_base com vírgula, mínimo 2 casas e 3ª casa preservada se != 0."""
    d = _parse_decimal_seguro(valor, padrao=Decimal("0")) or Decimal("0")
    d_3 = d.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
    s_3 = f"{d_3:.3f}"
    if s_3.endswith("0"):
        return f"{d_3.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):.2f}".replace(".", ",")
    return s_3.replace(".", ",")


def gerar_observacao_dobra(
    preco_base, peso, fator_dobra, valor_final, dobra
) -> str:
    """Gera texto 'Original + Dobra → Final' para Dobra == 'SIM'; vazio caso contrário."""
    if str(dobra or "").strip().upper() != "SIM":
        return ""
    p_base = _parse_decimal_seguro(preco_base, padrao=Decimal("0")) or Decimal("0")
    p_peso = _parse_decimal_seguro(peso, padrao=Decimal("0")) or Decimal("0")
    fator = _parse_decimal_seguro(fator_dobra, padrao=FATOR_DOBRA_PADRAO)
    if fator is None:
        fator = FATOR_DOBRA_PADRAO

    acrescimo_dobra = (p_peso * fator).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    acrescimo_fmt = f"{acrescimo_dobra:.2f}".replace(".", ",")
    preco_base_fmt = _formatar_preco_base(p_base)
    valor_final_fmt = _formatar_valor_unitario(valor_final)
    return (
        f"Original: R$ {preco_base_fmt} + Dobra: R$ {acrescimo_fmt} "
        f"→ Final: R$ {valor_final_fmt}"
    )


def normalizar_codigo(codigo: str) -> str:
    """Remove letras no início do código (e espaço/hífen logo após), preservando meio/fim."""
    codigo = (codigo or "").strip()
    if not re.search(r"\d", codigo):
        return codigo
    return re.sub(r"^[A-Za-z]+[\s\-]*", "", codigo)


def _normalizar_item(item: ItemCorte) -> ItemCorte:
    """Garante formatação consistente de dobra, valor_unitario e local_destino."""
    dobra_norm = normalizar_dobra(item.dobra)
    val = _formatar_valor_unitario(item.valor_unitario)
    dest = (item.local_destino or "").strip() or LOCAL_DESTINO_PADRAO
    updates: dict = {}
    if dobra_norm != item.dobra:
        updates["dobra"] = dobra_norm
    if val != item.valor_unitario:
        updates["valor_unitario"] = val
    if dest != item.local_destino:
        updates["local_destino"] = dest
    return item.model_copy(update=updates) if updates else item


def _parse_int_seguro(valor, padrao: int | None = None) -> int | None:
    m = re.search(r"\d+", str(valor or ""))
    if not m:
        return padrao
    try:
        return int(m.group())
    except ValueError:
        return padrao


def _limpar_texto_celula(valor) -> str:
    """Remove quebras de linha, tabs e espaços múltiplos dentro da célula do PDF."""
    return re.sub(r"\s+", " ", str(valor or "")).strip()


def _contem_palavra_teste(texto) -> bool:
    """Detecta 'teste' como token isolado (ex.: 'TESTE_0225'), evitando falsos positivos."""
    return bool(
        re.search(r"(?<![A-Za-zÀ-ÿ])teste(?![A-Za-zÀ-ÿ])", str(texto or ""), re.IGNORECASE)
    )


def _contar_linhas_peca_texto(texto_pagina: str) -> int:
    """Conta linhas no texto da página que aparentam ser peças com preço (R$)."""
    total = 0
    termos_resumo = ("total", "subtotal", "resumo", "custo", "valor total")
    for linha in (texto_pagina or "").splitlines():
        linha_limpa = linha.strip()
        if not linha_limpa:
            continue
        if linha_limpa.lower().startswith(termos_resumo):
            continue
        if re.search(r"r\$\s*[\d.,]+", linha_limpa, re.IGNORECASE) and (
            re.match(r"^\s*\d+\s+", linha_limpa)
            or re.search(r"[\d.,]+\s*x\s*[\d.,]+", linha_limpa, re.IGNORECASE)
        ):
            total += 1
    return total


def _montar_item_corte(
    codigo_bruto: str,
    dobra_bruta,
    preco_bruto,
    peso_bruto,
    quantidade: int,
    fator_dobra=FATOR_DOBRA_PADRAO,
    texto_linha_bruta: str = "",
) -> ItemCorte | None:
    """Constrói ItemCorte independente (sem agrupar) ou None se for item de TESTE."""
    ref_raw = _limpar_texto_celula(codigo_bruto)
    if not ref_raw:
        return None
    if _contem_palavra_teste(ref_raw) or _contem_palavra_teste(texto_linha_bruta):
        return None

    codigo_limpo = normalizar_codigo(ref_raw)
    if not codigo_limpo:
        return None

    dobra_norm = normalizar_dobra(dobra_bruta)
    p_base_dec = _parse_decimal_seguro(preco_bruto, padrao=Decimal("0")) or Decimal("0")
    peso_dec = _parse_decimal_seguro(peso_bruto, padrao=Decimal("0")) or Decimal("0")
    peso_str = _formatar_peso(peso_dec)
    valor_final = calcular_valor_unitario(p_base_dec, peso_dec, dobra_norm, fator_dobra)
    obs = gerar_observacao_dobra(p_base_dec, peso_dec, fator_dobra, valor_final, dobra_norm)

    return ItemCorte(
        codigo=codigo_limpo,
        dobra=dobra_norm,
        preco_base=p_base_dec,
        peso=peso_str,
        valor_unitario=valor_final,
        valor_manual=False,
        local_destino=LOCAL_DESTINO_PADRAO,
        quantidade_final=max(0, int(quantidade)),
        alterado_a_mao=False,
        observacao=obs,
    )


def _normalizar_cabecalho(celula) -> str:
    return re.sub(r"\s+", " ", str(celula or "").replace("\n", " ")).strip().lower()


def _extrair_itens_tabelas_pagina(
    pagina,
    fator_dobra,
    stats: dict | None = None,
    mapa_cols_herdado: dict[str, int] | None = None,
) -> tuple[list[ItemCorte], dict[str, int] | None]:
    """Extrai linhas de corte via `extract_tables()` de uma página pdfplumber."""
    itens: list[ItemCorte] = []
    tabelas = pagina.extract_tables() or []
    mapa_ultimo: dict[str, int] | None = (
        dict(mapa_cols_herdado) if mapa_cols_herdado else None
    )
    for tabela in tabelas:
        mapa_cols: dict[str, int] | None = None
        for linha in tabela:
            if not linha:
                continue
            celulas = [_limpar_texto_celula(c) for c in linha]
            if not any(celulas):
                continue

            normas = [_normalizar_cabecalho(c) for c in celulas]
            idx_ref_hdr = next(
                (i for i, h in enumerate(normas) if "referência" in h or "referencia" in h),
                None,
            )
            if idx_ref_hdr is not None and any(
                k in " ".join(normas) for k in ("dobra", "preço", "preco", "peso", "qtd")
            ):
                mapa_cols = {"ref": idx_ref_hdr}
                for i, h in enumerate(normas):
                    if "dobra" in h and "dobra" not in mapa_cols:
                        mapa_cols["dobra"] = i
                    elif ("preço" in h or "preco" in h) and "preco" not in mapa_cols:
                        mapa_cols["preco"] = i
                    elif ("qtd" in h or "peças" in h or "pecas" in h) and "qtd" not in mapa_cols:
                        mapa_cols["qtd"] = i
                    elif "pedido" in h and "pedido" not in mapa_cols:
                        mapa_cols["pedido"] = i
                    elif "peso" in h and "peso" not in mapa_cols:
                        mapa_cols["peso"] = i
                mapa_ultimo = dict(mapa_cols)
                continue

            mapa_ativo = mapa_cols
            eh_herdado = False
            if mapa_ativo is None and mapa_cols_herdado:
                mapa_ativo = mapa_cols_herdado
                eh_herdado = True

            if not mapa_ativo:
                continue

            idx_ref = mapa_ativo["ref"]
            if idx_ref >= len(celulas):
                continue
            if eh_herdado and len(celulas) <= max(mapa_ativo.values()):
                continue

            ref_bruta = celulas[idx_ref]
            if not ref_bruta:
                continue
            if _normalizar_cabecalho(ref_bruta) in ("referência", "referencia"):
                continue

            preco_bruto = (
                celulas[mapa_ativo["preco"]]
                if "preco" in mapa_ativo and mapa_ativo["preco"] < len(celulas)
                else ""
            )
            if eh_herdado and "preco" in mapa_ativo:
                preco_txt = preco_bruto.strip().lower()
                if "r$" not in preco_txt and not re.search(r"\d", preco_txt):
                    continue

            qtd_bruta = (
                celulas[mapa_ativo["qtd"]]
                if "qtd" in mapa_ativo and mapa_ativo["qtd"] < len(celulas)
                else ""
            )
            if not qtd_bruta.strip() and "pedido" in mapa_ativo and mapa_ativo["pedido"] < len(celulas):
                qtd_bruta = celulas[mapa_ativo["pedido"]]

            qtd = _parse_int_seguro(qtd_bruta)
            if qtd is None:
                continue

            texto_linha = " ".join(c for c in celulas if c)
            if _contem_palavra_teste(ref_bruta) or _contem_palavra_teste(texto_linha):
                if stats is not None:
                    stats["ignorados_teste"] = stats.get("ignorados_teste", 0) + 1
                continue

            dobra_bruta = (
                celulas[mapa_ativo["dobra"]]
                if "dobra" in mapa_ativo and mapa_ativo["dobra"] < len(celulas)
                else ""
            )
            peso_bruto = (
                celulas[mapa_ativo["peso"]]
                if "peso" in mapa_ativo and mapa_ativo["peso"] < len(celulas)
                else "0"
            )

            item = _montar_item_corte(
                ref_bruta,
                dobra_bruta,
                preco_bruto,
                peso_bruto,
                qtd,
                fator_dobra=fator_dobra,
                texto_linha_bruta=texto_linha,
            )
            if item is not None:
                itens.append(item)

    return itens, mapa_ultimo


def _extrair_texto_pdf_regex(texto_pagina: str, fator_dobra, stats: dict | None = None) -> list[ItemCorte]:
    """Parser secundário por linhas de texto quando extract_tables() não detectar grade."""
    itens: list[ItemCorte] = []
    padrao = re.compile(
        r"^\s*\d+\s+"                      # coluna #
        r"(?P<ref>.+?)\s+"                 # Referência (pode conter texto/TESTE)
        r"(?:(?P<dobra>SIM|N[ÃA]O)\s+)?"   # Dobra opcional
        r"R\$\s*(?P<preco>[\d.,]+)\s+"     # Preço
        r"[\d.,]+\s*x\s*[\d.,]+\s+"        # Dimensões
        r"(?P<qtd>\d+)\s+"                 # Qtd Peças
        r"(?:\d+\s+)?"                     # Pedido opcional
        r"\d{2}:\d{2}:\d{2}(?:\.\d+)?\s+"  # Tempo
        r"(?P<peso>[\d.,]+)\s*$",          # Peso / KG
        re.IGNORECASE,
    )
    for linha in (texto_pagina or "").splitlines():
        if "r$" not in linha.lower():
            continue
        m = padrao.match(linha)
        if not m:
            continue
        ref_bruta = _limpar_texto_celula(m.group("ref"))
        if _contem_palavra_teste(ref_bruta) or _contem_palavra_teste(linha):
            if stats is not None:
                stats["ignorados_teste"] = stats.get("ignorados_teste", 0) + 1
            continue
        qtd = _parse_int_seguro(m.group("qtd"))
        if qtd is None:
            continue
        item = _montar_item_corte(
            ref_bruta,
            m.group("dobra") or "",
            m.group("preco"),
            m.group("peso"),
            qtd,
            fator_dobra=fator_dobra,
            texto_linha_bruta=linha,
        )
        if item is not None:
            itens.append(item)
    return itens


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


def _extrair_bruto_com_gemini(conteudo: list[dict], aviso=None) -> list[ItemBrutoGemini]:
    """Chama a API Gemini com temperature=0.0 e retorna itens brutos sem cálculo."""
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key or api_key == "sua_chave_aqui":
        raise RuntimeError(
            "GEMINI_API_KEY não configurada.\n\n"
            f"Edite o arquivo .env em:\n{PASTA_BASE}"
        )

    client = genai.Client(api_key=api_key)
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
                generation_config={
                    "temperature": 0.0,
                    "thinking_level": NIVEL_RACIOCINIO,
                },
                response_format={
                    "type": "text",
                    "mime_type": "application/json",
                    "schema": ResultadoLeituraBruta.model_json_schema(),
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
        return ResultadoLeituraBruta.model_validate_json(texto).itens

    raise RuntimeError(f"Todos os modelos estão indisponíveis agora: {ultimo_erro}")


def _converter_brutos_em_itens(
    brutos: list[ItemBrutoGemini],
    fator_dobra=FATOR_DOBRA_PADRAO,
    stats: dict | None = None,
) -> list[ItemCorte]:
    """Filtra 'teste', limpa prefixos/dobra e calcula valor_unitario sem agrupar."""
    resultado: list[ItemCorte] = []
    for bruto in brutos:
        if _contem_palavra_teste(bruto.codigo):
            if stats is not None:
                stats["ignorados_teste"] = stats.get("ignorados_teste", 0) + 1
            continue
        item = _montar_item_corte(
            bruto.codigo,
            bruto.dobra,
            bruto.preco_base,
            bruto.peso,
            bruto.quantidade,
            fator_dobra=fator_dobra,
        )
        if item is not None:
            resultado.append(item)
    return resultado


_LOCK_PDFIUM = threading.Lock()


def _extrair_pdf_rapido_pypdfium2(
    caminho_pdf: str, fator_dobra=FATOR_DOBRA_PADRAO
) -> tuple[list[ItemCorte], int] | None:
    """Via rápida via PDFium (C++): extrai texto e valida todas as linhas contra o contador."""
    if pypdfium2 is None or not os.path.isfile(caminho_pdf):
        return None
    try:
        with _LOCK_PDFIUM:
            itens: list[ItemCorte] = []
            ign_total = 0
            doc = pypdfium2.PdfDocument(caminho_pdf)
            try:
                for i in range(len(doc)):
                    page = doc[i]
                    try:
                        tp = page.get_textpage()
                        try:
                            texto_pag = tp.get_text_bounded() or ""
                        finally:
                            tp.close()
                    finally:
                        page.close()

                    if not texto_pag.strip():
                        return None
                    esperadas = _contar_linhas_peca_texto(texto_pag)
                    st_pag = {"ignorados_teste": 0}
                    extraidos_pag = _extrair_texto_pdf_regex(
                        texto_pag, fator_dobra, stats=st_pag
                    )
                    lidas_pag = len(extraidos_pag) + st_pag["ignorados_teste"]
                    if esperadas > 0 and lidas_pag != esperadas:
                        return None
                    itens.extend(extraidos_pag)
                    ign_total += st_pag["ignorados_teste"]
            finally:
                doc.close()
            return (itens, ign_total) if itens else None
    except Exception:
        return None


def extrair_itens_pdf(
    caminho_pdf: str,
    fator_dobra=FATOR_DOBRA_PADRAO,
    aviso=None,
    stats: dict | None = None,
) -> list[ItemCorte]:
    """Extrai peças de PDF (via rápida PDFium -> pdfplumber -> fallback Gemini)."""
    rapido = _extrair_pdf_rapido_pypdfium2(caminho_pdf, fator_dobra)
    if rapido is not None:
        itens_rapidos, ign_rapidos = rapido
        if stats is not None:
            stats["ignorados_teste"] = stats.get("ignorados_teste", 0) + ign_rapidos
        return itens_rapidos

    itens: list[ItemCorte] = []
    nome_arq = os.path.basename(caminho_pdf)
    mapa_cols_atual: dict[str, int] | None = None
    paginas_escaneadas: list[int] = []

    with pdfplumber.open(caminho_pdf) as pdf:
        for num_pag, pagina in enumerate(pdf.pages, start=1):
            texto_pag = pagina.extract_text() or ""
            esperadas = _contar_linhas_peca_texto(texto_pag)

            stats_tab = {"ignorados_teste": 0}
            extraidos_pag, novo_mapa = _extrair_itens_tabelas_pagina(
                pagina,
                fator_dobra,
                stats=stats_tab,
                mapa_cols_herdado=mapa_cols_atual,
            )
            if novo_mapa is not None:
                mapa_cols_atual = novo_mapa
            ign_pag = stats_tab["ignorados_teste"]

            if not extraidos_pag or esperadas > len(extraidos_pag) + ign_pag:
                stats_rx = {"ignorados_teste": 0}
                extraidos_rx = _extrair_texto_pdf_regex(
                    texto_pag, fator_dobra, stats=stats_rx
                )
                if len(extraidos_rx) + stats_rx["ignorados_teste"] > len(extraidos_pag) + ign_pag:
                    extraidos_pag = extraidos_rx
                    ign_pag = stats_rx["ignorados_teste"]

            if stats is not None:
                stats["ignorados_teste"] = stats.get("ignorados_teste", 0) + ign_pag
            itens.extend(extraidos_pag)

            total_lidas = len(extraidos_pag) + ign_pag
            if esperadas > total_lidas:
                if stats is not None:
                    stats.setdefault("avisos_auditoria", []).append(
                        f"{nome_arq} (pág. {num_pag}): {esperadas} linha(s) com preço "
                        f"detectadas no texto, mas {total_lidas} processada(s) — confira esta página."
                    )
            elif not extraidos_pag and not ign_pag and not texto_pag.strip():
                paginas_escaneadas.append(num_pag)

    if itens:
        if paginas_escaneadas and stats is not None:
            pags_fmt = ", ".join(map(str, paginas_escaneadas))
            stats.setdefault("avisos_auditoria", []).append(
                f"{nome_arq}: página(s) {pags_fmt} sem texto selecionável "
                "(possível imagem escaneada) — confira ou envie como foto."
            )
        return itens

    # Fallback automático para PDF escaneado / sem tabela extraível via pdfplumber
    if stats is not None:
        stats.setdefault("avisos_auditoria", []).append(
            f"{nome_arq}: PDF sem texto selecionável lido via IA (Gemini) "
            "— recomenda-se conferência visual dos códigos e valores."
        )
    if aviso:
        aviso("⏳ PDF sem tabela selecionável. Lendo com Gemini...")
    with open(caminho_pdf, "rb") as arq:
        pdf_b64 = base64.b64encode(arq.read()).decode("utf-8")

    conteudo = [
        {
            "type": "document",
            "data": pdf_b64,
            "mime_type": "application/pdf",
        },
        {
            "type": "text",
            "text": (
                "Extraia todas as linhas válidas de peças deste PDF em JSON bruto, "
                "sem calcular dobra, sem incluir itens com 'TESTE' e sem agrupar códigos repetidos."
            ),
        },
    ]
    brutos = _extrair_bruto_com_gemini(conteudo, aviso=aviso)
    return _converter_brutos_em_itens(brutos, fator_dobra=fator_dobra, stats=stats)


def ler_folhas_com_gemini(
    imagens: list[Image.Image],
    fator_dobra=FATOR_DOBRA_PADRAO,
    aviso=None,
    stats: dict | None = None,
) -> list[ItemCorte]:
    """Envia imagens para a IA e retorna cada linha extraída sem agrupar duplicatas."""
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
            f"Leia as {len(imagens)} folha(s) de corte a laser enviadas e extraia cada "
            "linha válida em JSON bruto, ignorando itens com 'TESTE' e sem agrupar nem "
            "somar códigos repetidos."
        ),
    })
    brutos = _extrair_bruto_com_gemini(conteudo, aviso=aviso)
    return _converter_brutos_em_itens(brutos, fator_dobra=fator_dobra, stats=stats)


# =============================================================================
# INTERFACE GRÁFICA
# =============================================================================
class App:
    # Mapeamento das colunas editáveis direto na célula
    COLUNAS_EDITAVEIS = {
        "#1": "codigo",
        "#2": "dobra",
        "#3": "peso",
        "#4": "valor_unitario",
        "#5": "local_destino",
        "#6": "quantidade_final",
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Leitor de Corte a Laser → WK Radar (Entrada da Produção)")
        self.root.geometry("1260x720")
        self.root.minsize(1040, 560)

        self.itens: list[ItemCorte] = []
        self.ocupado = False
        self.calibrando = False
        self.digitando = False
        self.pausado = False
        self.parar_solicitado = False
        self._esc_anterior = False
        self._f12_anterior = False
        self.uid_pausado_popup: str | None = None
        self._ignorados_teste_lote = 0

        cfg = carregar_config()
        self.pos_operacao: tuple[int, int] | None = cfg.get("pos_operacao")
        self.pos_produto: tuple[int, int] | None = cfg.get("pos_produto")
        self.hwnd_wk: int = 0

        # Variáveis do Painel de Calibração
        setas_ini = cfg.get("setas_operacao", SETAS_OPERACAO_PADRAO)
        fator_dobra_ini = cfg.get("fator_dobra", FATOR_DOBRA_PADRAO)
        espera_prod_ini = cfg.get("espera_apos_produto", ESPERA_APOS_PRODUTO)
        enters_ini = cfg.get("enters_apos_qtd", ENTERS_APOS_QTD_PADRAO)
        espera_ini = cfg.get("espera_apos_ok", ESPERA_APOS_OK)
        self.var_setas_operacao = tk.StringVar(value=str(setas_ini))
        self.var_fator_dobra = tk.StringVar(value=str(fator_dobra_ini).replace(".", ","))
        self.var_espera_apos_produto = tk.StringVar(value=f"{float(espera_prod_ini):.1f}")
        self.var_enters_apos_qtd = tk.StringVar(value=str(enters_ini))
        self.var_espera_apos_ok = tk.StringVar(value=f"{float(espera_ini):.1f}")

        self._montar_interface()
        self._atualizar_indicador_calibracao()
        self.var_fator_dobra.trace_add("write", lambda *_a: self._ao_mudar_fator_dobra())

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

        self.btn_abrir = ttk.Button(topo, text="📂 Abrir PDF / Foto(s)",
                                    command=self.abrir_arquivos)
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
            text=("Duplo clique na linha: editar campos (Código, Dobra, Peso, Preço, Destino, Qtd)  •  "
                  "F12 = Pausar/Retomar  |  Esc = Parar  |  Mouse no canto sup. esq. = Abortar"),
            foreground="#005fb8", padding=(10, 2, 10, 0),
        ).pack(fill="x")

        # --- tabela central
        meio = ttk.Frame(self.root, padding=(10, 6, 10, 6))
        meio.pack(fill="both", expand=True)

        colunas = (
            "codigo", "dobra", "peso", "valor_unit", "destino", "qtd", "status", "observacao"
        )
        self.tree = ttk.Treeview(meio, columns=colunas, show="headings",
                                 selectmode="extended")
        self.tree.heading("codigo", text="Código")
        self.tree.heading("dobra", text="Dobra")
        self.tree.heading("peso", text="Peso (kg)")
        self.tree.heading("valor_unit", text="Valor Unitário")
        self.tree.heading("destino", text="Destino")
        self.tree.heading("qtd", text="Quantidade")
        self.tree.heading("status", text="Status")
        self.tree.heading("observacao", text="Observação")

        self.tree.column("codigo", width=190, anchor="w")
        self.tree.column("dobra", width=70, anchor="center")
        self.tree.column("peso", width=85, anchor="center")
        self.tree.column("valor_unit", width=105, anchor="center")
        self.tree.column("destino", width=70, anchor="center")
        self.tree.column("qtd", width=85, anchor="center")
        self.tree.column("status", width=165, anchor="center")
        self.tree.column("observacao", width=330, minwidth=260, anchor="w")

        self.tree.tag_configure("dobra_sim", background=COR_DOBRA_SIM)
        self.tree.tag_configure("lancado", background=COR_LANCADO)
        self.tree.tag_configure("popup", background=COR_POPUP)
        self.tree.tag_configure("nao_cadastrado", background=COR_NAO_CADASTRADO)

        scroll_y = ttk.Scrollbar(meio, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(meio, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        meio.rowconfigure(0, weight=1)
        meio.columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

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
        self.menu.add_command(label="📝 Editar item completo...",
                              command=self._editar_selecionada_dialogo)
        self.menu.add_command(label="✏️ Editar Quantidade",
                              command=self.editar_qtd_selecionada)
        self.menu.add_command(label="🔄 Alternar Dobra (SIM / NÃO)",
                              command=self._alternar_dobra_selecionada)
        self.menu.add_command(label="⚖️ Editar Peso (kg)",
                              command=lambda: self._editar_selecionada("#3"))
        self.menu.add_command(label="💲 Editar Valor Unitário",
                              command=lambda: self._editar_selecionada("#4"))
        self.menu.add_command(label="📦 Editar Destino",
                              command=lambda: self._editar_selecionada("#5"))
        self.menu.add_command(label="🔤 Editar Código",
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

        ttk.Label(linha_cal, text="Setas Operação:").pack(side="left")
        self.spn_setas_operacao = ttk.Spinbox(
            linha_cal, from_=0, to=20, increment=1, width=4,
            textvariable=self.var_setas_operacao, font=("Segoe UI", 10))
        self.spn_setas_operacao.pack(side="left", padx=(4, 10))

        ttk.Label(linha_cal, text="Fator Dobra:").pack(side="left")
        self.spn_fator_dobra = ttk.Spinbox(
            linha_cal, from_=0.0, to=100.0, increment=0.1, width=6,
            textvariable=self.var_fator_dobra, font=("Segoe UI", 10))
        self.spn_fator_dobra.pack(side="left", padx=(4, 10))

        ttk.Label(linha_cal, text="Espera Produto (s):").pack(side="left")
        self.spn_espera_apos_produto = ttk.Spinbox(
            linha_cal, from_=0.2, to=15.0, increment=0.1, width=5,
            textvariable=self.var_espera_apos_produto, font=("Segoe UI", 10))
        self.spn_espera_apos_produto.pack(side="left", padx=(4, 10))

        ttk.Label(linha_cal, text="Enters até OK:").pack(side="left")
        self.spn_enters_apos_qtd = ttk.Spinbox(
            linha_cal, from_=1, to=50, increment=1, width=4,
            textvariable=self.var_enters_apos_qtd, font=("Segoe UI", 10))
        self.spn_enters_apos_qtd.pack(side="left", padx=(4, 10))

        ttk.Label(linha_cal, text="Espera OK (s):").pack(side="left")
        self.spn_espera_apos_ok = ttk.Spinbox(
            linha_cal, from_=0.2, to=15.0, increment=0.1, width=5,
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
            fator_dobra=self._ler_fator_dobra(),
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

    def _idx_por_uid(self, uid: str | int | None) -> int | None:
        if uid is None:
            return None
        if isinstance(uid, int):
            return uid if 0 <= uid < len(self.itens) else None
        for i, it in enumerate(self.itens):
            if it.uid == uid:
                return i
        return None

    def _limpar_estado_popup(self):
        self.uid_pausado_popup = None
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
            and self._idx_por_uid(self.uid_pausado_popup) is not None
        )
        self._definir_botoes_retomada(pode_retomar_popup)

    def _ler_setas_operacao(self) -> int:
        try:
            valor = int(self.var_setas_operacao.get().strip())
            return max(0, min(valor, 50))
        except ValueError:
            self.var_setas_operacao.set(str(SETAS_OPERACAO_PADRAO))
            return SETAS_OPERACAO_PADRAO

    def _ler_fator_dobra(self, corrigir_campo: bool = True) -> Decimal:
        d = _parse_decimal_seguro(self.var_fator_dobra.get(), padrao=None)
        if d is None or d < Decimal("0") or d > Decimal("100"):
            if corrigir_campo:
                self.var_fator_dobra.set(str(FATOR_DOBRA_PADRAO).replace(".", ","))
            return FATOR_DOBRA_PADRAO
        return d

    def _ao_mudar_fator_dobra(self):
        if not hasattr(self, "tree"):
            return
        d = _parse_decimal_seguro(self.var_fator_dobra.get(), padrao=None)
        if d is not None and Decimal("0") <= d <= Decimal("100"):
            self.recalcular_valores_tabela(d)

    def recalcular_valores_tabela(self, fator_dobra: Decimal | None = None):
        """Recalcula Valor Unitário e Observação a partir do preco_base original."""
        fator = fator_dobra if fator_dobra is not None else self._ler_fator_dobra()
        filhos = set(self.tree.get_children())
        alterou = False
        for i, item in enumerate(self.itens):
            if item.valor_manual:
                novo_val = item.valor_unitario
            else:
                novo_val = calcular_valor_unitario(item.preco_base, item.peso, item.dobra, fator)
            nova_obs = gerar_observacao_dobra(
                item.preco_base, item.peso, fator, novo_val, item.dobra
            )
            if novo_val != item.valor_unitario or nova_obs != item.observacao:
                item = item.model_copy(update={
                    "valor_unitario": novo_val,
                    "observacao": nova_obs,
                })
                self.itens[i] = item
                alterou = True
                if item.uid in filhos:
                    self.tree.item(
                        item.uid, tags=self._tags_item(item),
                        values=self._valores_linha(item),
                    )
        if alterou:
            self._atualizar_resumo()

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
            if item.tipo_popup == "FALTA_MATERIAL":
                return "⚠️ Falta de material"
            return "⚠️ Pop-up detectado"
        if item.nao_cadastrado:
            return "⏭️ Não cadastrado"
        if item.lancado:
            return "✅ Resolvida manual" if item.resolvida_manual else "✅ Lançado"
        if item.quantidade_final <= 0:
            return "⚪ Ignorado (0)"
        return "⏳ Pendente"

    def _tags_item(self, item: ItemCorte) -> tuple[str, ...]:
        if item.erro_popup:
            return ("popup",)
        if item.nao_cadastrado:
            return ("nao_cadastrado",)
        if item.lancado:
            return ("lancado",)
        if normalizar_dobra(item.dobra) == "SIM":
            return ("dobra_sim",)
        return ()

    def _valores_linha(self, item: ItemCorte) -> tuple:
        return (
            item.codigo,
            normalizar_dobra(item.dobra),
            item.peso,
            item.valor_unitario,
            item.local_destino,
            item.quantidade_final,
            self._status_item_texto(item),
            item.observacao,
        )

    def atualizar_tabela(self):
        self._fechar_editor(salvar=False)
        self.tree.delete(*self.tree.get_children())
        for item in self.itens:
            self.tree.insert("", "end", iid=item.uid, tags=self._tags_item(item),
                             values=self._valores_linha(item))
        self._atualizar_resumo()

    def _atualizar_resumo(self):
        n_dobra = sum(1 for it in self.itens if normalizar_dobra(it.dobra) == "SIM")
        n_lanc = sum(1 for it in self.itens if it.lancado and not it.nao_cadastrado)
        n_pulados = sum(1 for it in self.itens if it.nao_cadastrado)
        soma_pecas = sum(max(0, it.quantidade_final) for it in self.itens)
        resumo = (
            f"{len(self.itens)} linhas ({soma_pecas} peças)  •  "
            f"{n_dobra} c/ dobra  •  {n_lanc} lançados (✅)"
        )
        if n_pulados:
            resumo += f"  •  {n_pulados} não cadastrados (⏭️)"
        if self._ignorados_teste_lote:
            resumo += f"  •  {self._ignorados_teste_lote} TESTE ignorado(s)"
        self.var_total.set(resumo)

    def _atualizar_linha_tabela(self, alvo):
        idx = self._idx_por_uid(alvo)
        if idx is None:
            return
        item = self.itens[idx]
        iid = item.uid
        if iid in self.tree.get_children():
            self.tree.item(iid, tags=self._tags_item(item),
                           values=self._valores_linha(item))
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
        self._atualizar_resumo()

    def _marcar_lancado(self, alvo, valor: bool = True, resolvida_manual: bool = False):
        idx = self._idx_por_uid(alvo)
        if idx is not None:
            self.itens[idx] = self.itens[idx].model_copy(update={
                "lancado": valor,
                "erro_popup": False,
                "nao_cadastrado": False,
                "tipo_popup": "",
                "resolvida_manual": resolvida_manual if valor else False,
            })
            self._atualizar_linha_tabela(idx)

    def _marcar_nao_cadastrado(self, alvo):
        idx = self._idx_por_uid(alvo)
        if idx is not None:
            self.itens[idx] = self.itens[idx].model_copy(update={
                "lancado": True,
                "nao_cadastrado": True,
                "erro_popup": False,
                "tipo_popup": "",
                "resolvida_manual": False,
            })
            self._atualizar_linha_tabela(idx)

    def _marcar_popup_item(self, alvo, tipo_popup: str = "OUTRO"):
        idx = self._idx_por_uid(alvo)
        if idx is not None:
            self.itens[idx] = self.itens[idx].model_copy(update={
                "lancado": False,
                "erro_popup": True,
                "nao_cadastrado": False,
                "tipo_popup": tipo_popup,
                "resolvida_manual": False,
            })
            self._atualizar_linha_tabela(idx)

    def alternar_digitado_selecionados(self):
        if self.ocupado:
            return
        uids = list(self.tree.selection())
        if not uids:
            iid = self._selecionada()
            if iid is not None:
                uids = [iid]
        indices = [i for u in uids if (i := self._idx_por_uid(u)) is not None]
        if not indices:
            return
        novo_estado = not all(self.itens[idx].lancado for idx in indices)
        for idx in indices:
            self.itens[idx] = self.itens[idx].model_copy(update={
                "lancado": novo_estado,
                "erro_popup": False,
                "nao_cadastrado": False,
                "tipo_popup": "",
                "resolvida_manual": False,
            })
        self.atualizar_tabela()
        for idx in indices:
            self.tree.selection_add(self.itens[idx].uid)

    def desmarcar_todos_digitados(self):
        if self.ocupado or not self.itens:
            return
        self._limpar_estado_popup()
        if not any(it.lancado or it.erro_popup or it.nao_cadastrado for it in self.itens):
            self._set_status("Nenhum item está marcado como lançado.")
            return
        self.itens = [
            it.model_copy(update={
                "lancado": False,
                "erro_popup": False,
                "nao_cadastrado": False,
                "tipo_popup": "",
                "resolvida_manual": False,
            })
            for it in self.itens
        ]
        self.atualizar_tabela()
        self._set_status("🔄 Marcações de lançado (✅) limpas.")

    # --------------------------------------------------------------- leitura
    def abrir_arquivos(self):
        if self.ocupado:
            return
        caminhos = filedialog.askopenfilenames(
            title="Selecione um ou mais PDFs ou fotos das folhas de corte",
            filetypes=[
                ("PDF e Imagens", "*.pdf *.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff"),
                ("Arquivos PDF", "*.pdf"),
                ("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not caminhos:
            return
        self._processar_caminhos(list(caminhos), acumular=False)

    abrir_fotos = abrir_arquivos

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

        if isinstance(conteudo, Image.Image):
            self._processar_entradas([("imagem", conteudo)], acumular=True)
            return
        if isinstance(conteudo, list) and conteudo:
            caminhos = [c for c in conteudo if isinstance(c, str) and os.path.isfile(c)]
            if caminhos:
                self._processar_caminhos(caminhos, acumular=True)
                return

        messagebox.showwarning(
            "Nada para colar",
            "Não há imagem nem arquivo PDF na área de transferência.\n\n"
            "No WhatsApp Web: clique com o botão direito na foto → 'Copiar imagem' "
            "ou copie um PDF no Explorer e use Ctrl+V aqui.",
        )

    def _processar_caminhos(self, caminhos: list[str], acumular: bool):
        entradas: list[tuple[str, object]] = []
        try:
            for c in caminhos:
                if str(c).lower().endswith(".pdf"):
                    entradas.append(("pdf", c))
                else:
                    img = Image.open(c)
                    img.load()
                    entradas.append(("imagem", img))
        except Exception as e:
            messagebox.showerror("Erro ao abrir arquivo", str(e))
            return
        if not entradas:
            return
        self._processar_entradas(entradas, acumular=acumular)

    def _processar_entradas(self, entradas: list[tuple[str, object]], acumular: bool):
        fator_dobra = self._ler_fator_dobra()
        self._set_ocupado(True, modo="ia")
        total_ent = len(entradas)
        self._set_status(f"⏳ Lendo {total_ent} arquivo(s)... aguarde.")

        def tarefa():
            try:
                stats: dict = {"ignorados_teste": 0, "avisos_auditoria": []}
                novos: list[ItemCorte] = []
                lote_imgs: list[Image.Image] = []
                lote_pdfs: list[str] = []
                erros_arquivos: list[str] = []
                aviso_fn = lambda m: self.root.after(0, self._set_status, m)

                def _flush_imagens():
                    if lote_imgs:
                        novos.extend(
                            ler_folhas_com_gemini(
                                lote_imgs, fator_dobra=fator_dobra,
                                aviso=aviso_fn, stats=stats,
                            )
                        )
                        lote_imgs.clear()

                def _flush_pdfs():
                    if not lote_pdfs:
                        return
                    qtd_pdfs = len(lote_pdfs)
                    if qtd_pdfs == 1:
                        caminho_pdf = lote_pdfs[0]
                        nome_pdf = os.path.basename(caminho_pdf)
                        aviso_fn(f"⏳ Lendo PDF: {nome_pdf}...")
                        try:
                            novos.extend(
                                extrair_itens_pdf(
                                    caminho_pdf, fator_dobra=fator_dobra,
                                    aviso=aviso_fn, stats=stats,
                                )
                            )
                        except Exception as exc:
                            if total_ent == 1:
                                raise
                            erros_arquivos.append(f"{nome_pdf}: {exc}")
                        finally:
                            lote_pdfs.clear()
                        return

                    aviso_fn(f"⏳ Lendo {qtd_pdfs} PDF(s) em paralelo...")
                    resultados: list[tuple[list[ItemCorte], dict, str | None]] = [
                        ([], {}, None) for _ in range(qtd_pdfs)
                    ]

                    def _ler_um(indice: int, caminho_pdf: str):
                        st_local: dict = {"ignorados_teste": 0, "avisos_auditoria": []}
                        try:
                            itens_pdf = extrair_itens_pdf(
                                caminho_pdf, fator_dobra=fator_dobra,
                                aviso=None, stats=st_local,
                            )
                            return indice, itens_pdf, st_local, None
                        except Exception as exc:
                            return indice, [], st_local, f"{os.path.basename(caminho_pdf)}: {exc}"

                    concluidos = 0
                    with ThreadPoolExecutor(max_workers=min(4, qtd_pdfs)) as pool:
                        futuros = [
                            pool.submit(_ler_um, i, c) for i, c in enumerate(lote_pdfs)
                        ]
                        for fut in as_completed(futuros):
                            idx_res, itens_res, st_res, err_res = fut.result()
                            resultados[idx_res] = (itens_res, st_res, err_res)
                            concluidos += 1
                            nome_pdf = os.path.basename(lote_pdfs[idx_res])
                            aviso_fn(f"⏳ Lendo PDFs ({concluidos}/{qtd_pdfs}): {nome_pdf}...")

                    for itens_res, st_res, err_res in resultados:
                        if err_res:
                            erros_arquivos.append(err_res)
                        else:
                            novos.extend(itens_res)
                            stats["ignorados_teste"] += st_res.get("ignorados_teste", 0)
                            stats["avisos_auditoria"].extend(
                                st_res.get("avisos_auditoria", [])
                            )
                    lote_pdfs.clear()

                for tipo, dado in entradas:
                    if tipo == "pdf":
                        _flush_imagens()
                        lote_pdfs.append(str(dado))
                    else:
                        _flush_pdfs()
                        lote_imgs.append(dado)  # type: ignore[arg-type]
                _flush_pdfs()
                _flush_imagens()

                if not novos and erros_arquivos:
                    raise RuntimeError(
                        "Não foi possível ler os arquivos selecionados:\n\n"
                        + "\n".join(f"• {e}" for e in erros_arquivos)
                    )

                avisos = list(stats.get("avisos_auditoria", [])) + [
                    f"Erro ao ler {e}" for e in erros_arquivos
                ]
                self.root.after(
                    0, self._leitura_ok, novos, acumular,
                    stats["ignorados_teste"], avisos,
                )
            except Exception as e:
                self.root.after(0, self._leitura_erro, e)

        threading.Thread(target=tarefa, daemon=True).start()

    def _leitura_ok(
        self,
        novos: list[ItemCorte],
        acumular: bool,
        ignorados_teste: int = 0,
        avisos: list[str] | None = None,
    ):
        if acumular:
            self.itens.extend(novos)
            self._ignorados_teste_lote += ignorados_teste
        else:
            self._limpar_estado_popup()
            self.itens = list(novos)
            self._ignorados_teste_lote = ignorados_teste
        self.atualizar_tabela()
        self._set_ocupado(False)
        modo = "adicionados ao final" if acumular else "carregados"
        extra = f" ({ignorados_teste} TESTE ignorado(s))" if ignorados_teste else ""
        if avisos:
            self._set_status(
                f"⚠️ {len(novos)} itens {modo}{extra} — {len(avisos)} aviso(s) de conferência."
            )
            messagebox.showwarning(
                "Conferência recomendada na leitura do PDF",
                "\n".join(f"• {a}" for a in avisos),
            )
        else:
            self._set_status(f"✅ {len(novos)} itens {modo}{extra}.")

    def _leitura_erro(self, erro: Exception):
        self._set_ocupado(False)
        self._set_status("❌ Erro na leitura.")
        msg = str(erro)
        if "APIConnectionError" in type(erro).__name__ or "getaddrinfo" in msg:
            msg = ("Sem conexão com a internet (ou ela caiu durante a leitura).\n\n"
                   "Verifique a internet e tente de novo.\n\n"
                   f"Detalhe técnico: {erro}")
        messagebox.showerror("Erro na leitura", msg)

    # ---------------------------------------------------------------- edição
    def limpar_lista(self):
        if self.ocupado or not self.itens:
            return
        if messagebox.askyesno("Limpar Lista", "Apagar todos os itens da lista?"):
            self.itens.clear()
            self._ignorados_teste_lote = 0
            self._limpar_estado_popup()
            self.atualizar_tabela()
            self._set_status("Lista limpa. Pronto para um novo lote.")

    def remover_selecionados(self, _event=None):
        if self.ocupado:
            return
        uids_sel = set(self.tree.selection())
        if not uids_sel:
            return
        if self.uid_pausado_popup in uids_sel:
            self._limpar_estado_popup()
        removidos = sum(1 for it in self.itens if it.uid in uids_sel)
        self.itens = [it for it in self.itens if it.uid not in uids_sel]
        self.atualizar_tabela()
        self._set_ocupado(False)
        self._set_status(f"{removidos} linha(s) removida(s).")

    # ------------------------------------------------- edição direta e diálogo
    def _selecionada(self):
        """uid da linha em foco/selecionada (ou None)."""
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
        idx = self._idx_por_uid(iid)
        if idx is None:
            return
        if coluna == "#7":                  # Status: alterna ✅ Lançado
            self._marcar_lancado(iid, not self.itens[idx].lancado)
        elif coluna == "#8":                # Observação: coluna informativa
            return "break"
        else:
            # Abre a janela de edição completa focando o campo clicado
            self._abrir_dialogo(iid, coluna_foco=coluna)
        return "break"

    def editar_qtd_selecionada(self):
        self._editar_selecionada("#6")

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

    def _alternar_dobra_selecionada(self):
        if self.ocupado:
            return
        iid = self._selecionada()
        idx = self._idx_por_uid(iid)
        if idx is None:
            return
        antigo = self.itens[idx]
        nova_dobra = "NÃO" if normalizar_dobra(antigo.dobra) == "SIM" else "SIM"
        fator = self._ler_fator_dobra()
        novo_val = calcular_valor_unitario(
            antigo.preco_base, antigo.peso, nova_dobra, fator
        )
        nova_obs = gerar_observacao_dobra(
            antigo.preco_base, antigo.peso, fator, novo_val, nova_dobra
        )
        novo = antigo.model_copy(update={
            "dobra": nova_dobra,
            "valor_unitario": novo_val,
            "valor_manual": False,
            "observacao": nova_obs,
        })
        self._registrar_correcao(idx, antigo, novo)

    def _tecla_na_tabela(self, event):
        """Digitar um número com a linha selecionada já começa a editar a Qtd (#6)."""
        if self.ocupado:
            return None
        if event.char and event.char.isdigit():
            iid = self._selecionada()
            if iid:
                self._editar_celula(iid, "#6", texto_inicial=event.char)
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

    def _registrar_correcao(self, alvo, antigo: ItemCorte, novo: ItemCorte):
        """Atualiza o item sem aplicar normalizar_codigo (controle manual preservado)."""
        idx = self._idx_por_uid(alvo)
        if idx is None:
            return
        novo = _normalizar_item(novo)
        notas = []
        if novo.quantidade_final != antigo.quantidade_final:
            notas.append(f"qtd: {antigo.quantidade_final} → {novo.quantidade_final}")
        if novo.codigo != antigo.codigo:
            notas.append(f"código: {antigo.codigo} → {novo.codigo}")
        if novo.dobra != antigo.dobra:
            notas.append(f"dobra: {antigo.dobra} → {novo.dobra}")
        if novo.peso != antigo.peso:
            notas.append(f"peso: {antigo.peso} → {novo.peso}")
        if novo.valor_unitario != antigo.valor_unitario:
            notas.append(f"valor: {antigo.valor_unitario} → {novo.valor_unitario}")
        if novo.local_destino != antigo.local_destino:
            notas.append(f"destino: {antigo.local_destino} → {novo.local_destino}")
        self.itens[idx] = novo
        self.atualizar_tabela()
        iid = novo.uid
        if iid in self.tree.get_children():
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
        idx = self._idx_por_uid(iid)
        if idx is None:
            return
        antigo = self.itens[idx]
        nova = max(0, antigo.quantidade_final + delta)
        if nova != antigo.quantidade_final:
            self._registrar_correcao(
                idx, antigo, antigo.model_copy(update={"quantidade_final": nova}))

    def _editar_celula(self, iid: str, coluna: str, texto_inicial: str | None = None):
        """Abre uma caixinha de edição em cima da célula (#1 a #6)."""
        self._fechar_editor(salvar=True)
        campo = self.COLUNAS_EDITAVEIS.get(coluna)
        idx = self._idx_por_uid(iid)
        if not campo or idx is None or iid not in self.tree.get_children():
            return
        self.tree.see(iid)
        self.tree.update_idletasks()
        bbox = self.tree.bbox(iid, coluna)
        if not bbox:
            return
        x, y, w, h = bbox
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
        self._editor = (ent, iid, campo)
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
        ent, iid, campo = self._editor
        self._editor = None
        try:
            texto = ent.get().strip()
            ent.destroy()
        except tk.TclError:
            return
        idx = self._idx_por_uid(iid)
        if not salvar or idx is None:
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
        elif campo == "dobra":
            nova_dobra = normalizar_dobra(texto)
            fator = self._ler_fator_dobra()
            novo_val = calcular_valor_unitario(
                antigo.preco_base, antigo.peso, nova_dobra, fator
            )
            nova_obs = gerar_observacao_dobra(
                antigo.preco_base, antigo.peso, fator, novo_val, nova_dobra
            )
            novo = antigo.model_copy(update={
                "dobra": nova_dobra,
                "valor_unitario": novo_val,
                "valor_manual": False,
                "observacao": nova_obs,
            })
        elif campo == "peso":
            novo_peso = _formatar_peso(texto)
            fator = self._ler_fator_dobra()
            novo_val = calcular_valor_unitario(
                antigo.preco_base, novo_peso, antigo.dobra, fator
            )
            nova_obs = gerar_observacao_dobra(
                antigo.preco_base, novo_peso, fator, novo_val, antigo.dobra
            )
            novo = antigo.model_copy(update={
                "peso": novo_peso,
                "valor_unitario": novo_val,
                "valor_manual": False,
                "observacao": nova_obs,
            })
        elif campo == "valor_unitario":
            novo_val = _formatar_valor_unitario(texto) or VALOR_UNITARIO_PADRAO
            fator = self._ler_fator_dobra()
            nova_obs = gerar_observacao_dobra(
                antigo.preco_base, antigo.peso, fator, novo_val, antigo.dobra
            )
            novo = antigo.model_copy(update={
                "valor_unitario": novo_val,
                "valor_manual": True,
                "observacao": nova_obs,
            })
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
    def _abrir_dialogo(self, iid: str, coluna_foco: str = "#6"):
        self._fechar_editor(salvar=True)
        idx = self._idx_por_uid(iid)
        if idx is None:
            return
        item = self.itens[idx]

        janela = tk.Toplevel(self.root)
        janela.title("Editar item — Entrada da Produção")
        janela.transient(self.root)
        janela.resizable(False, False)

        frm = ttk.Frame(janela, padding=15)
        frm.pack(fill="both", expand=True)

        pbase_inicial = _formatar_preco_base(item.preco_base)
        var_cod = tk.StringVar(value=item.codigo)
        var_dob = tk.StringVar(value=normalizar_dobra(item.dobra))
        var_pbase = tk.StringVar(value=pbase_inicial)
        var_peso = tk.StringVar(value=item.peso)
        var_val = tk.StringVar(value=item.valor_unitario)
        var_dst = tk.StringVar(value=item.local_destino)
        var_qtd = tk.StringVar(value=str(item.quantidade_final))
        var_obs = tk.StringVar(value=item.observacao)
        var_dig = tk.BooleanVar(value=item.lancado)
        val_editado_manual = [False]
        atualizando_auto = [False]

        ttk.Label(frm, text="Código (Produto):").grid(row=0, column=0, sticky="w", pady=4)
        ent_cod = ttk.Entry(frm, textvariable=var_cod, width=42)
        ent_cod.grid(row=0, column=1, pady=4, sticky="w")

        ttk.Label(frm, text="Dobra:").grid(row=1, column=0, sticky="w", pady=4)
        cmb_dob = ttk.Combobox(frm, textvariable=var_dob, values=("NÃO", "SIM"),
                               state="readonly", width=13)
        cmb_dob.grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Preço Base:").grid(row=2, column=0, sticky="w", pady=4)
        ent_pbase = ttk.Entry(frm, textvariable=var_pbase, width=16)
        ent_pbase.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Peso (kg):").grid(row=3, column=0, sticky="w", pady=4)
        ent_peso = ttk.Entry(frm, textvariable=var_peso, width=16)
        ent_peso.grid(row=3, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Valor Unitário (Final):").grid(row=4, column=0, sticky="w", pady=4)
        ent_val = ttk.Entry(frm, textvariable=var_val, width=16)
        ent_val.grid(row=4, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Local de Destino:").grid(row=5, column=0, sticky="w", pady=4)
        ent_dst = ttk.Entry(frm, textvariable=var_dst, width=16)
        ent_dst.grid(row=5, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Quantidade:").grid(row=6, column=0, sticky="w", pady=4)
        ent_qtd = ttk.Entry(frm, textvariable=var_qtd, width=16)
        ent_qtd.grid(row=6, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Observação:").grid(row=7, column=0, sticky="w", pady=4)
        ent_obs = ttk.Entry(frm, textvariable=var_obs, width=52, state="readonly")
        ent_obs.grid(row=7, column=1, pady=4, sticky="w")

        ttk.Checkbutton(frm, text="Já lançado no WK Radar (✅)",
                        variable=var_dig).grid(row=8, column=1, sticky="w", pady=4)

        def _preco_base_corrente() -> Decimal:
            txt_pbase = var_pbase.get().strip()
            if txt_pbase == pbase_inicial:
                return item.preco_base
            return _parse_decimal_seguro(txt_pbase, padrao=Decimal("0")) or Decimal("0")

        def _recalcular_no_dialogo(*_args):
            atualizando_auto[0] = True
            try:
                val_editado_manual[0] = False
                fator = self._ler_fator_dobra()
                p_base = _preco_base_corrente()
                novo_val = calcular_valor_unitario(p_base, var_peso.get(), var_dob.get(), fator)
                var_val.set(novo_val)
                var_obs.set(
                    gerar_observacao_dobra(p_base, var_peso.get(), fator, novo_val, var_dob.get())
                )
            finally:
                atualizando_auto[0] = False

        def _ao_mudar_val_no_dialogo(*_args):
            if atualizando_auto[0]:
                return
            val_editado_manual[0] = True
            fator = self._ler_fator_dobra()
            p_base = _preco_base_corrente()
            val_fmt = _formatar_valor_unitario(var_val.get())
            var_obs.set(
                gerar_observacao_dobra(p_base, var_peso.get(), fator, val_fmt, var_dob.get())
            )

        var_dob.trace_add("write", _recalcular_no_dialogo)
        var_pbase.trace_add("write", _recalcular_no_dialogo)
        var_peso.trace_add("write", _recalcular_no_dialogo)
        var_val.trace_add("write", _ao_mudar_val_no_dialogo)

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
            dobra = normalizar_dobra(var_dob.get())
            fator = self._ler_fator_dobra()
            p_base_dec = _preco_base_corrente()
            peso = _formatar_peso(var_peso.get())
            valor_unit = _formatar_valor_unitario(var_val.get())
            obs_final = gerar_observacao_dobra(p_base_dec, peso, fator, valor_unit, dobra)
            destino = var_dst.get().strip() or LOCAL_DESTINO_PADRAO
            marcado_lancado = var_dig.get()
            novo = item.model_copy(update={
                "codigo": codigo,
                "dobra": dobra,
                "preco_base": p_base_dec,
                "peso": peso,
                "valor_unitario": valor_unit,
                "valor_manual": val_editado_manual[0] or (item.valor_manual and valor_unit == item.valor_unitario),
                "local_destino": destino,
                "quantidade_final": qtd,
                "observacao": obs_final,
                "lancado": marcado_lancado,
                "erro_popup": False,
                "nao_cadastrado": False,
                "tipo_popup": "",
                "resolvida_manual": item.resolvida_manual if marcado_lancado else False,
            })
            janela.destroy()
            self._registrar_correcao(idx, item, novo)

        botoes = ttk.Frame(frm)
        botoes.grid(row=9, column=0, columnspan=2, pady=(12, 0), sticky="e")
        ttk.Button(botoes, text="Salvar", command=salvar).pack(side="right", padx=4)
        ttk.Button(botoes, text="Cancelar", command=janela.destroy).pack(side="right")

        janela.bind("<Return>", salvar)
        janela.bind("<Escape>", lambda _e: janela.destroy())

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
            "#2": cmb_dob,
            "#3": ent_peso,
            "#4": ent_val,
            "#5": ent_dst,
            "#6": ent_qtd,
        }
        alvo = mapa_foco.get(coluna_foco, ent_qtd)
        alvo.focus_set()
        if isinstance(alvo, ttk.Entry):
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
        com_qtd = [(it.uid, it) for it in self.itens if it.quantidade_final > 0]
        if not com_qtd:
            messagebox.showwarning("Nada para digitar",
                                   "Não há itens com quantidade maior que zero.")
            return

        pendentes = [(uid, it) for uid, it in com_qtd if not it.lancado]
        if not pendentes:
            if not messagebox.askyesno(
                "Tudo já lançado",
                "Todos os itens já estão marcados como lançados (✅).\n\n"
                "Deseja desmarcar todos e digitar tudo novamente?",
            ):
                return
            self.desmarcar_todos_digitados()
            pendentes = [(it.uid, it) for it in self.itens if it.quantidade_final > 0]

        self._iniciar_fila_digitacao(pendentes)

    def iniciar_digitacao_da_selecionada(self):
        """Digita a partir da linha selecionada até o final da lista."""
        if self.ocupado:
            return
        self._fechar_editor(salvar=True)
        if not self._garantir_calibrado():
            return
        iid = self._selecionada()
        idx_inicio = self._idx_por_uid(iid)
        if idx_inicio is None:
            messagebox.showinfo(
                "Digitar da Linha Selecionada",
                "Selecione na tabela a linha a partir da qual deseja começar a digitar.",
            )
            return
        fila = [(it.uid, it) for it in self.itens[idx_inicio:] if it.quantidade_final > 0]
        if not fila:
            messagebox.showwarning(
                "Nada para digitar",
                "Não há itens com quantidade maior que zero a partir da linha selecionada.",
            )
            return
        self._iniciar_fila_digitacao(fila)

    def _iniciar_fila_digitacao(self, fila: list[tuple[str | int, ItemCorte]]):
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
        copia_fila = [(alvo, it.model_copy()) for alvo, it in fila]
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

    def _digitar(self, fila: list[tuple[str | int, ItemCorte]],
                 pos_operacao, pos_produto, hwnd_wk: int,
                 setas_operacao: int, enters_apos_qtd: int, espera_apos_ok: float,
                 espera_apos_produto: float = ESPERA_APOS_PRODUTO):
        total = len(fila)
        digitados_agora = 0
        pulados_agora = 0
        operacao_ja_selecionada = False

        def checar_popup(idx_atual, codigo_atual: str) -> str:
            if not verificar_popup(hwnd_wk):
                return "OK"

            hwnd_popup = obter_hwnd_ativo()
            titulo_popup = obter_titulo_janela(hwnd_popup)
            tipo = classificar_popup(hwnd_popup)

            if tipo == "NAO_CADASTRADO":
                pyautogui.press("esc")
                if self._dormir_interrompivel(0.35):
                    return "PARAR"
                if verificar_popup(hwnd_wk):
                    pyautogui.press("enter")
                    if self._dormir_interrompivel(0.35):
                        return "PARAR"
                if not verificar_popup(hwnd_wk):
                    if self._dormir_interrompivel(0.2):
                        return "PARAR"
                    limpar_campo_produto_apos_popup()
                    if self._aguardar_pausa_ou_parada():
                        return "PARAR"
                    self.root.after(0, self._marcar_nao_cadastrado, idx_atual)
                    return "PULAR"
                # Trava de segurança: não fechou com Esc nem Enter -> pausa normal
                tipo = "OUTRO"

            try:
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            except Exception:
                pass
            self.root.after(
                0, self._pausar_por_popup, idx_atual, codigo_atual, titulo_popup, tipo
            )
            return "PAUSAR"

        def _lancar_uma_peca(idx, item: ItemCorte, pular_operacao: bool = False) -> str:
            # 1. Selecionar "Entrada da Produção" com reset pela tecla Home
            if not pular_operacao:
                res = checar_popup(idx, item.codigo)
                if res != "OK":
                    return res
                pyautogui.click(pos_operacao)
                if self._dormir_interrompivel(0.25):
                    return "PARAR"

                res = checar_popup(idx, item.codigo)
                if res != "OK":
                    return res
                pyautogui.press("home")
                if self._dormir_interrompivel(0.15):
                    return "PARAR"

                for s in range(setas_operacao):
                    res = checar_popup(idx, item.codigo)
                    if res != "OK":
                        return res
                    pyautogui.press("down")
                    if s < setas_operacao - 1 and self._dormir_interrompivel(0.1):
                        return "PARAR"
                if setas_operacao > 0 and self._dormir_interrompivel(0.15):
                    return "PARAR"

                res = checar_popup(idx, item.codigo)
                if res != "OK":
                    return res
                pyautogui.press("enter")
                if self._dormir_interrompivel(0.35):
                    return "PARAR"

            # 2. Pular direto para o campo Produto (ignorando campos dinâmicos no meio)
            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            pyautogui.click(pos_produto)
            if self._dormir_interrompivel(0.25 if pular_operacao else 0.35):
                return "PARAR"

            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            digitar_campo_seguro(item.codigo, limpar=True)
            if self._aguardar_pausa_ou_parada():
                return "PARAR"

            # 3. Fluxo por Enter a partir de Produto
            # Enter 1: sai de Produto / passa por Tipo de Estoque (consulta no banco)
            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            pyautogui.press("enter")
            if self._dormir_interrompivel(espera_apos_produto):
                return "PARAR"
            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res

            # Enter 2: vai para Valor Unitário
            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            pyautogui.press("enter")
            if self._dormir_interrompivel(0.25):
                return "PARAR"

            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            valor_unit = (item.valor_unitario or "").strip() or VALOR_UNITARIO_PADRAO
            digitar_campo_seguro(valor_unit, limpar=True)
            if self._aguardar_pausa_ou_parada():
                return "PARAR"

            # Enter 3: vai para Local de Destino (aguarda 1.5s, limpa e digita destino)
            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            pyautogui.press("enter")
            if self._dormir_interrompivel(1.5):
                return "PARAR"

            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            destino_val = (item.local_destino or "").strip() or LOCAL_DESTINO_PADRAO
            digitar_campo_seguro(destino_val, limpar=True)
            if self._aguardar_pausa_ou_parada():
                return "PARAR"

            # Enter 4: sai de Local de Destino / campo intermediário
            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            pyautogui.press("enter")
            if self._dormir_interrompivel(0.25):
                return "PARAR"

            # Enter 5: vai para Quantidade
            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            pyautogui.press("enter")
            if self._dormir_interrompivel(0.25):
                return "PARAR"

            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            digitar_campo_seguro(str(item.quantidade_final), limpar=True)
            if self._aguardar_pausa_ou_parada():
                return "PARAR"

            # Enters finais até o OK
            for _ in range(enters_apos_qtd):
                res = checar_popup(idx, item.codigo)
                if res != "OK":
                    return res
                pyautogui.press("enter")
                if self._dormir_interrompivel(0.12):
                    return "PARAR"

            # 15. Monitoramento ativo pós-OK (checa pop-up a cada 0.1s durante espera_apos_ok)
            restante_ok = espera_apos_ok
            while restante_ok > 0:
                if self._aguardar_pausa_ou_parada():
                    return "PARAR"
                res = checar_popup(idx, item.codigo)
                if res != "OK":
                    return res
                passo_ok = min(0.1, restante_ok)
                time.sleep(passo_ok)
                restante_ok -= passo_ok
            if self.parar_solicitado:
                return "PARAR"
            res = checar_popup(idx, item.codigo)
            if res != "OK":
                return res
            return "LANCADO"

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

                idx_exib = self._idx_por_uid(idx)
                linha_lbl = f"linha {idx_exib + 1}" if idx_exib is not None else "item"
                self.root.after(
                    0, self._set_status,
                    f"⌨️ Digitando {pos}/{total} ({linha_lbl}): {item.codigo} "
                    f"(qtd {item.quantidade_final})  [F12 pausa • Esc para]"
                )

                pular_op = operacao_ja_selecionada
                operacao_ja_selecionada = False
                resultado = _lancar_uma_peca(idx, item, pular_operacao=pular_op)
                if resultado == "PAUSAR":
                    return
                if resultado == "PARAR":
                    break
                if resultado == "PULAR":
                    pulados_agora += 1
                    operacao_ja_selecionada = True
                    continue

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
        msg_final = f"✅ Concluído! {digitados_agora} itens lançados no WK Radar."
        if pulados_agora > 0:
            msg_final = (
                f"✅ Concluído! {digitados_agora} lançado(s) e "
                f"{pulados_agora} pulado(s) (não cadastrado)."
            )
        self.root.after(0, self._fim_digitacao, msg_final)

    def _pausar_por_popup(self, alvo, codigo: str, titulo_popup: str,
                          tipo_popup: str = "OUTRO"):
        self.digitando = False
        self.pausado = False
        self.parar_solicitado = False
        idx = self._idx_por_uid(alvo)
        self.uid_pausado_popup = self.itens[idx].uid if idx is not None else None
        if self.uid_pausado_popup is not None:
            self._marcar_popup_item(self.uid_pausado_popup, tipo_popup=tipo_popup)
        self._set_ocupado(False)
        motivo = "FALTA DE MATERIAL" if tipo_popup == "FALTA_MATERIAL" else "POP-UP"
        self._set_status(
            f"⚠️ PAUSADO POR {motivo} na peça '{codigo}' — Janela: {titulo_popup}"
        )
        self._restaurar_janela_app()

    def continuar_proxima_peca(self):
        """Marca a peça pausada como resolvida manualmente e retoma pela próxima."""
        if self.ocupado or self.uid_pausado_popup is None:
            return
        if not self._garantir_calibrado():
            return
        idx_pausado = self._idx_por_uid(self.uid_pausado_popup)
        if idx_pausado is None:
            self._limpar_estado_popup()
            return
        uid_pausado = self.itens[idx_pausado].uid
        self._marcar_lancado(uid_pausado, True, resolvida_manual=True)
        self._limpar_estado_popup()
        fila = [
            (it.uid, it)
            for it in self.itens[idx_pausado + 1:]
            if it.quantidade_final > 0 and not it.lancado
        ]
        if not fila:
            self._set_status("✅ Peça marcada como resolvida manual. Não há mais peças pendentes.")
            return
        self._iniciar_fila_digitacao(fila)

    def repetir_peca_pausada(self):
        """Retoma a digitação começando novamente pela peça onde ocorreu o pop-up."""
        if self.ocupado or self.uid_pausado_popup is None:
            return
        if not self._garantir_calibrado():
            return
        idx_pausado = self._idx_por_uid(self.uid_pausado_popup)
        if idx_pausado is None:
            self._limpar_estado_popup()
            return
        uid_pausado = self.itens[idx_pausado].uid
        self._marcar_lancado(uid_pausado, False, resolvida_manual=False)
        self._limpar_estado_popup()
        fila = [
            (it.uid, it)
            for it in self.itens[idx_pausado:]
            if it.quantidade_final > 0 and not it.lancado
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
