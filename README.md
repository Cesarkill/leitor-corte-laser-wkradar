# 🔦 Leitor de Folhas de Corte a Laser → WK Radar

Aplicação desktop em Python (`tkinter`) para Windows que lê fotos de folhas de controle de corte a laser — inclusive com **rasuras e anotações feitas à caneta** — usando a API do **Google Gemini** (`google-genai` com *Structured Outputs* via Pydantic) e preenche automaticamente a tela de **Entrada da Produção** no ERP **WK Radar**.

---

## ✨ Principais Funcionalidades

- **Leitura Inteligente com IA (Google Gemini):**
  - Mantém o código exato da peça (hífens, barras, pontos e letras).
  - Prioriza automaticamente quantidades corrigidas à caneta quando o valor impresso estiver riscado/rasurado, destacando a linha em **amarelo (`#fff3cd`)** para conferência humana.
  - Preenche automaticamente `Valor Unitário` (padrão `"0,00"`) e `Local de Destino` (padrão `"23"`).
- **Suporte a Múltiplas Folhas e Lotes:**
  - **📂 Abrir Foto(s) da Folha:** seleciona 1 ou mais imagens de uma vez e envia todas juntas na mesma requisição, **consolidando e somando** códigos repetidos entre as páginas.
  - **📋 Colar Imagem (`Ctrl+V`):** cola imagens direto do WhatsApp Web (ou arquivos copiados no Explorer) **acumulando** os novos itens ao final da tabela.
  - **🗑️ Limpar Lista:** zera a tabela para começar um novo lote.
- **Conferência e Edição Rápida na Tabela:**
  - **Duplo clique** em qualquer linha abre a janela para editar **Código (Produto)**, **Valor Unitário**, **Local de Destino**, **Quantidade Final**, marcação de rasura, observação e status.
  - **Atalhos na tabela:** selecionar uma linha e digitar um número altera a quantidade na hora; `+` e `-` somam ou subtraem 1; `Delete` remove linhas selecionadas.
- **Robô Digitador para o WK Radar (Entrada da Produção):**
  - Contagem regressiva de **10 segundos** para posicionar o mouse sobre o campo **OPERAÇÃO** no WK Radar.
  - Captura a posição do campo **OPERAÇÃO** e preenche item por item usando `pyperclip` + `Ctrl+V` (sem erros de acentuação/cedilha).
  - Marca em tempo real cada linha concluída com **`✅ Digitado`** (fundo verde `#d4edda`).
  - Permite **▶️ Digitar da Linha Selecionada** para retomar um lote do ponto exato onde parou.
  - **Painel de calibração na janela:** ajuste rápido de `Espera após salvar (s)` (padrão `1.5s`) e `Tecla para gravar` (`ENTER` ou `F2`).
  - **Atalhos globais de segurança** (funcionam mesmo com a janela do WK Radar em foco):
    - `F12`: **Pausar / Retomar** a digitação.
    - `Esc`: **Parar** a digitação imediatamente.
    - **Fail-Safe do PyAutoGUI:** jogar o mouse para o canto superior esquerdo da tela aborta a execução.

---

## 🚀 Como Instalar e Executar

### 1. Pré-requisitos
- Windows 10 ou 11
- Python 3.11+ (testado com Python 3.12)
- Uma chave de API do Google Gemini ([Google AI Studio](https://aistudio.google.com/apikey))

### 2. Criar o ambiente virtual e instalar dependências

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configurar a chave da API (`.env`)

Copie o arquivo `.env.example` para `.env` e coloque sua chave:

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

O executável será gerado em `dist\LeitorCorteLaser.exe` (o script já copia o arquivo `.env` automaticamente para a pasta `dist\`).

> **Importante:** se você levar o `LeitorCorteLaser.exe` para outro computador, mantenha o arquivo `.env` na mesma pasta do `.exe`.

---

## ⌨️ Como Usar com o WK Radar

1. Abra ou cole (`Ctrl+V`) as fotos das folhas de controle de corte a laser.
2. Confira os itens na tabela (especialmente as **linhas amarelas** com rasura à caneta).
3. Ajuste, se necessário, o tempo de **`Espera após salvar (s)`** (`1.5s`) e a **`Tecla para gravar`** (`ENTER` ou `F2`).
4. Clique em **`⌨️ Digitar no WK Radar (10s)`** (ou **`▶️ Digitar da Linha Selecionada (10s)`**).
5. Durante os 10 segundos de contagem regressiva, **clique no campo `OPERAÇÃO` no WK Radar e deixe o mouse parado sobre ele**.
6. Ao final da contagem, o robô captura a posição do campo `OPERAÇÃO` e lança automaticamente todas as peças pendentes.
