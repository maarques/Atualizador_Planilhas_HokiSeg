# Atualizador de Planilhas HokiSeg  
![Python](https://img.shields.io/badge/python-3.8%2B-blue) ![Pandas](https://img.shields.io/badge/pandas-1.0%2B-blueviolet) ![OpenPyXL](https://img.shields.io/badge/openpyxl-3.0%2B-green) ![Tkinter](https://img.shields.io/badge/tkinter-GUI-orange)  

Uma ferramenta com **interface gráfica** para ler analíticos de pagamento de comissões em PDF (inicialmente para Porto Seguro e Amil) e inseri-los de forma consolidada numa planilha Excel de controle financeiro da HokiSeg.

---

## ✨ Funcionalidades Principais  
- Interface gráfica simples: qualquer usuário seleciona o PDF de origem e a planilha de destino.  
- Leitura inteligente de PDF: extração de dados tabulares complexos de extratos de comissão.  
- Processamento e consolidação: agrupa automaticamente lançamentos duplicados (por exemplo, o mesmo cliente aparece várias vezes no PDF) e soma seus valores de prêmio e comissão.  
- Regras de negócio embutidas: aplica filtros específicos (por exemplo: ignora linhas “PIC – Bonus Mensal”) e lógicas de arredondamento.  
- Atualização segura: adiciona os novos dados ao final da planilha Excel, **sem apagar ou sobrescrever** dados já existentes.  
- Portabilidade: o projeto pode ser empacotado num único arquivo `.exe` para execução em Windows sem necessidade de instalar Python.

---

## 🧭 Como Usar (para usuários)  
1. Execute o arquivo `.exe` (ex: `AutomacaoHokiSeg.exe`).  
2. Na tela:  
   - Clique em **„1. Selecionar PDF”** e escolha o extrato de comissão em PDF.  
   - Clique em **„2. Selecionar Planilha”** e escolha o arquivo da planilha (ex: `Planilha Financeira out-2025.xlsx`).
   - Clique em **„3. Escolha a Seguradora”**.
   - Clique em **„4. Processar e Atualizar Planilha”**.  
3. Aguarde a barra de log exibir a mensagem de sucesso.  
4. Pronto! Um novo arquivo será salvo (ex: `Planilha_financeira_out-2025_ATUALIZADA.xlsx`) na mesma pasta da planilha original.

---

## 🛠️ Como Executar (para desenvolvedores)
Se você quiser rodar o código-fonte e fazer melhorias:

```Bash

# Clonar o repositório
git clone https://github.com/maarques/Atualizador_Planilhas_HokiSeg.git
cd Atualizador_Planilhas_HokiSeg

# Criar e ativar ambiente virtual
python -m venv venv

# Windows
venv\Scripts\activate
# Unix/macOS
source venv/bin/activate

# Instalar dependências
pip install -r requirements.txt

# Executar a aplicação
python main.py
```
### Para gerar o executável (.exe no Windows)
Use o PyInstaller. O comando abaixo gera um único arquivo executável sem o terminal de console.

```Bash

# O --windowed (ou --noconsole) é importante para aplicações de GUI
pyinstaller --onefile --windowed --name="AtualizadorHokiSeg_v12" main.py
# O arquivo AtualizadorHokiSeg_v12.exe será criado na pasta dist/.
```

## ⚙️ Lógica de Processamento (Multi-Parser)
O sistema agora é capaz de processar múltiplos layouts de PDF, um para cada seguradora.

Destino dos dados: Planilha Excel (Planilha financeira … .xlsx), aba “Comissão”.

Fontes de Dados (Parsers)
O sistema seleciona o parser correto com base na escolha do usuário na interface.

1. Porto Seguro
Fonte: Extrato analítico de pagamentos de comissões (PDF).

Filtro de exclusão: Linhas contendo o texto “PIC – Bonus Mensal” são ignoradas.

Regras Específicas: A coluna "Parcela a Receber" é fixada com o valor "12" durante a extração.

2. Amil
Fonte: Extrato de comissão (PDF).

Lógica: Utiliza RegEx (Expressões Regulares) para identificar e extrair dados de múltiplos "blocos de contrato" dentro da mesma página do PDF.

Dados Extraídos: Cliente, Apólice, Parcela, Dt. Pagamento, Prêmio, Taxa (%), e Comissão.

Agrupamento (Pós-processamento)
Após a extração de todas as fontes, os dados são unificados e agrupados por Cliente + Apólice + Parcela. Durante o agrupamento:

Valor (Prêmio): é somado (sum).

Valor Comissão: é somado (sum) e arredondado para cima ao centavo.

Porcentagem (Taxa): prevalece a maior taxa (max) encontrada no grupo.

## 🧩 Estrutura do Projeto
A estrutura foi atualizada para suportar múltiplos parsers de forma modular.
```
Atualizador_Planilhas_HokiSeg/
├── .gitignore
├── main.py              # Ponto de entrada: inicia a aplicação
├── ui.py                # Interface gráfica (Tkinter)
├── processing.py        # Lógica central: orquestra o UI, os parsers e o processamento
├── data_processing.py   # Lógica de negócio: agrupamento com Pandas, escrita no Excel
├── parsers/             # Módulo contendo todos os parsers de seguradoras
│   ├── __init__.py
│   ├── amil.py   # Parser específico da Amil
│   └── porto_seguro.py  # Parser específico da Porto Seguro
└── requirements.txt     # Dependências do projeto
```
