# 🔦 Leitor de Corte a Laser → WK Radar (Entrada da Produção)

Aplicação desktop em Python (`tkinter`) para Windows que lê **relatórios PDF (padrão HMAGRO / CypCut)** localmente em alta velocidade — além de **PDFs escaneados e fotos** (com anotações/rasuras à caneta) via **Google Gemini** — e lança automaticamente os itens na tela de **Entrada da Produção (Inclusão)** do ERP **WK Radar**.

---

## ✨ Principais Funcionalidades

### 📄 Leitura Rápida, Auditada e Sem Perdas (PDF + Imagens)
- **Extração Nativa em Alta Velocidade (`pypdfium2` + `pdfplumber`):**
  - **Via expressa em C++ (`pypdfium2`):** lê relatórios PDF digitais instantaneamente (10 PDFs / 430 linhas em menos de 1 segundo) e valida linha por linha contra o total de itens com preço da página.
  - **Fallback estrutural (`pdfplumber`):** se alguma página divergir na contagem ou variar o formato, extrai automaticamente pelas tabelas com **herança de cabeçalho entre páginas**, limpeza de quebras de linha dentro das células e **auditoria automática** (avisa página e arquivo caso falte alguma linha ou haja página escaneada sem texto).
  - **Lotes em paralelo (`ThreadPoolExecutor`):** ao abrir vários PDFs de uma vez, processa em paralelo mantendo a ordem original na tabela; se um arquivo do lote estiver corrompido, os demais continuam e o erro aparece resumido no aviso final.
  - **Fallback com IA (Google Gemini, `temperature=0.0`):** usado automaticamente para PDFs totalmente escaneados, fotos enviadas por arquivo ou imagens coladas da área de transferência.
- **Regras de Extração e Limpeza:**
  - **Filtro de `TESTE`:** descarta automaticamente linhas cuja referência/linha contenha a palavra isolada `TESTE` (ex.: `3.28.053 TESTE_0225`), exibindo no resumo quantos itens foram ignorados.
  - **Normalização de Código:** remove letras de prefixo no início antes dos números (ex.: `L3.28.051` → `3.28.051`), preservando pontos, barras, hífens e letras no meio/final (`3.12.084D`), apenas na leitura inicial (edições manuais na tabela são preservadas).
  - **Sem Agrupamento de Repetidos:** cada linha do relatório permanece independente (identificada por `uid` único), mantendo exatamente a ordem e as ocorrências originais sem somar quantidades de códigos iguais.

### 🧮 Cálculo Exato de Dobra e Coluna `Observação`
- **Cálculo em `Decimal` (`ROUND_HALF_UP`):**
  - Coluna `Dobra` vazia ou em branco é normalizada automaticamente para **`NÃO`**; somente **`SIM`** explícito aplica o acréscimo de dobra.
  - Fórmula: `Valor Unitário = Preço Base + (Peso × Fator Dobra)` quando `Dobra == "SIM"`, ou apenas `Preço Base` quando `Dobra == "NÃO"` (fallback `"0,00"`).
  - O `preco_base` original fica protegido em `Decimal` dentro de cada item, permitindo alterar **`Fator Dobra`** (padrão `1,2`), **`Dobra`** ou **`Peso`** a qualquer momento sem nunca somar o acréscimo duas vezes.
- **Coluna Informativa `Observação`:**
  - Para peças com `Dobra == "SIM"`, exibe o detalhamento: ``Original: R$ 43,77 + Dobra: R$ 4,70 → Final: R$ 48,47``.
  - Para peças com `Dobra == "NÃO"`, permanece em branco.
  - Coluna exclusivamente informativa (não editável por duplo clique na célula e não enviada ao WK Radar). Se o `Valor Unitário` for alterado manualmente, atualiza apenas o trecho `→ Final: R$ ...`.

### 🖥️ Conferência e Edição na Tabela (8 Colunas)
Colunas: **`Código` | `Dobra` | `Peso (kg)` | `Valor Unitário` | `Destino` | `Quantidade` | `Status` | `Observação`**

- **Legenda de cores:**
  - Azul claro (`#e7f3ff`): item com `Dobra == "SIM"`.
  - Verde claro (`#d4edda`): item já lançado (`✅ Lançado` / `✅ Resolvida manual`).
  - Laranja claro (`#ffe5b4`): pulado automaticamente por não estar cadastrado (`⏭️ Não cadastrado`).
  - Vermelho claro (`#f8d7da`): pausado por pop-up (`⚠️ Falta de material` ou `⚠️ Pop-up detectado`).
- **Atalhos e edição:**
  - **📂 Abrir PDF / Foto(s):** aceita múltiplos arquivos `.pdf`, `.png`, `.jpg`, `.jpeg`, `.bmp`, `.webp`, `.tif`, `.tiff`.
  - **📋 Colar Imagem (`Ctrl+V`):** cola imagens direto do WhatsApp Web ou arquivos copiados no Explorer (acumulando ao final da lista).
  - **Duplo clique** nas colunas `#1` a `#6` (ou tecla `F2`) abre o diálogo completo (Código, Dobra, Preço Base, Peso, Valor Unitário, Destino, Quantidade, prévia da Observação e Status); duplo clique em **Status** (`#7`) alterna `✅ Lançado`.
  - **Teclas rápidas na tabela:** digitar número edita `Quantidade`; `+` / `-` somam ou subtraem `1`; `Delete` remove linhas; **botão direito** abre o menu de contexto (incluindo alternar dobra, editar células e **▶️ Digitar a partir desta linha**).

### 🤖 Robô Digitador no WK Radar (Calibração, Digitação Segura e Pop-ups)
- **Calibração Persistente em 2 Pontos (`config.json`):**
  - Clique em **`🎯 Calibrar Posições`** apenas uma vez: contagem em 2 etapas de **5s + 5s** (com aviso sonoro) para gravar as coordenadas dos campos **Operação** e **Produto** e o identificador da janela do WK Radar, salvando tudo em `config.json`.
  - Indicador visual **`✅ Calibrado`** / **`⚠️ Não calibrado`**, com início imediato ao clicar em **`▶ Digitar no WK Radar`** ou **`▶️ Da Linha Selecionada`**.
- **Parâmetros Configuráveis no Painel:**
  - `Setas Operação` (padrão `1`)
  - `Fator Dobra` (padrão `1,2`, recalcula a tabela em tempo real)
  - `Espera Produto (s)` (padrão `1.0`, aguarda consulta do banco no WK Radar)
  - `Enters até OK` (padrão `18`)
  - `Espera OK (s)` (padrão `1.5`, monitoramento ativo pós-OK)
- **Digitação Direta Segura (`digitar_campo_seguro`):**
  - Limpa cada campo via Win32 (`keybd_event` com scan codes reais e `KEYEVENTF_EXTENDEDKEY`, imune ao **NumLock** ligado) e digita direto via `pyautogui.write` para códigos/números simples (usando `pyperclip` + `Ctrl+V` apenas como fallback para caracteres especiais). `Local de Destino` usa padrão `"23"`.
- **Classificação Híbrida de Pop-ups (`classificar_popup`):**
  - **Camada 1 (instantânea):** lê título e controles filhos da janela via `ctypes` (`GetWindowTextW` / `EnumChildWindows`).
  - **Camada 2 (fallback em RAM):** captura recorte da caixa de aviso apenas em memória RAM (sem gravar arquivos em disco) e consulta o Gemini.
  - **`NAO_CADASTRADO`:** fecha o aviso (`Esc` / `Enter`), apaga o código inválido no próprio campo **Produto** sem disparar loop de `OnExit` ao sair do campo, marca a linha em laranja (`⏭️ Não cadastrado`), mantém a Operação já selecionada e pula automaticamente para a próxima peça.
  - **`FALTA_MATERIAL` ou `OUTRO`:** pausa com alerta sonoro, destaca em vermelho, restaura a janela do app e habilita os botões **`▶ Continuar da Próxima Peça`** e **`🔁 Repetir Peça Pausada`**.
- **Atalhos Globais de Segurança:** `F12` (**Pausar / Retomar**), `Esc` (**Parar**) e **Fail-Safe** do PyAutoGUI (mover o mouse para o canto superior esquerdo aborta na hora).

---

## 🚀 Como Instalar e Executar

### 1. Pré-requisitos
- Windows 10 ou 11
- Python 3.11+ (testado com Python 3.12)
- Chave de API do Google Gemini ([Google AI Studio](https://aistudio.google.com/apikey)) — necessária apenas para leitura de fotos, PDFs escaneados ou fallback visual de pop-ups (PDFs digitais rodam 100% localmente)

### 2. Criar o ambiente virtual e instalar dependências

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configurar a chave da API (`.env`)

Copie o arquivo `.env.example` para `.env` e informe sua chave:

```env
GEMINI_API_KEY=sua_chave_aqui
```

### 4. Abrir a aplicação

```powershell
.\.venv\Scripts\python.exe main.py
```

---

## 📦 Como Gerar o Executável Portátil (`.exe`)

Execute o script [gerar_exe.bat](gerar_exe.bat) com duplo clique ou pelo terminal:

```powershell
.\gerar_exe.bat
```

O script encerra automaticamente qualquer `LeitorCorteLaser.exe` que esteja aberto, executa um build limpo (`pyinstaller --clean --noconfirm --onefile --windowed`) e gera o executável em `dist\LeitorCorteLaser.exe` (copiando o `.env` para `dist\.env`).

> **Importante:** ao mover o `LeitorCorteLaser.exe` para outra pasta ou computador, mantenha o arquivo `.env` na mesma pasta do `.exe`. As posições calibradas e parâmetros são salvos automaticamente em `config.json` ao lado do executável.

---

## ⌨️ Como Usar com o WK Radar

1. Abra os relatórios **PDF** ou cole imagens (`Ctrl+V`) em **`📂 Abrir PDF / Foto(s)`**.
2. Confira na tabela a coluna **`Dobra`**, o **`Valor Unitário`** calculado e o detalhamento na coluna **`Observação`**.
3. Na primeira vez (ou se mudar a resolução/posição da janela), deixe a tela de **Entrada da Produção** do WK Radar aberta, clique em **`🎯 Calibrar Posições`** e posicione o mouse sobre **Operação** (5s) e depois sobre **Produto** (5s). A calibração fica salva em `config.json`.
4. Com o WK Radar aberto atrás, clique em **`▶ Digitar no WK Radar`** (ou selecione uma linha e clique em **`▶️ Da Linha Selecionada`**) para iniciar os lançamentos automaticamente.
