<div align="center">

# **🤖 Atualizador de Planilhas HokiSeg 🤖**
</div>

<div align="center"> <img src="https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python&logoColor=white" alt="Python"> <img src="https://img.shields.io/badge/Pandas-2.0-purple?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"> <img src="https://img.shields.io/badge/OpenPyXL-3.1-green?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="OpenPyXL"> <img src="https://img.shields.io/badge/Tkinter-GUI-orange?style=for-the-badge&logo=python&logoColor=white" alt="Tkinter"> <img src="https://img.shields.io/badge/PyInstaller-5.13-gray?style=for-the-badge&logo=windowsterminal&logoColor=white" alt="PyInstaller"> </div>

Ferramenta com interface gráfica para ler analíticos de pagamento de comissões em PDF (ex: Porto Seguro (A ideia é sempre atualizar o projeto para poder ler mais analíticos de seguradoras diferentes)) e inseri-los de forma consolidada na Planilha Financeira HokiSeg.

<div align="center"> <h2>✨ Funcionalidades Principais ✨</h2> </div>

Interface Gráfica Simples: Permite que qualquer usuário selecione o PDF de origem e a planilha de destino.

Leitura Inteligente de PDF: Extrai dados tabulares complexos dos extratos de comissão.

Processamento e Consolidação: Agrupa automaticamente lançamentos duplicados (ex: "Markus"), somando seus valores de prêmio e comissão.

Regras de Negócio Embutidas: Aplica filtros específicos (ex: ignora "PIC - Bonus Mensal") e lógicas de arredondamento.

Atualização Segura: Adiciona os novos dados ao final da planilha Excel, sem apagar ou sobrescrever dados existentes.

Portabilidade: O projeto é empacotado em um único arquivo .exe que roda em qualquer computador Windows sem precisar instalar Python.

<div align="center"> <h2>🚀 Como Usar (Para Usuários) 🚀</h2> </div>

A aplicação foi desenhada para ser o mais simples possível.

Execute o arquivo AutomacaoHokiSeg.exe.

Na tela principal, clique em "1. Selecionar PDF" e escolha o extrato de comissão baixado.

Clique em "2. Selecionar Planilha" e escolha o arquivo Planilha financeira out-2025.xlsx (ou a versão mais atual).

Clique no botão verde "3. Processar e Atualizar Planilha".

Aguarde a barra de log mostrar a mensagem de sucesso.

Pronto! Um novo arquivo (ex: Planilha financeira out-2025_ATUALIZADA.xlsx) será salvo na mesma pasta da planilha original, contendo os novos dados.

<div align="center"> <h2>🔧 Como Executar (Para Desenvolvedores) 🔧</h2> </div>

Se você quiser rodar o projeto a partir do código-fonte para fazer melhorias:

Clone o repositório:

```Bash

git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio
```
Crie e ative um ambiente virtual:

```Bash

python -m venv venv
.\venv\Scripts\activate
```
Instale as dependências:

```Bash

pip install -r requirements.txt
```
Execute a aplicação:

```Bash

python main.py
<div align="center"> <h3>📦 Para gerar um novo .exe 📦</h3> </div>
```
Use o PyInstaller após instalar as dependências:

```Bash

# Comando para gerar o .exe único e sem console
pyinstaller --onefile --noconsole --name="AutomacaoHokiSeg" main.py
# O executável final estará na pasta dist/.
```
<div align="center"> <h2>⚙️ Regras de Negócio Implementadas ⚙️</h2> </div>

Este script contém lógicas de negócio específicas para o processo da HokiSeg:

Fonte de Dados: Extrato Analítico de Pagamentos de Comissões da Porto Seguro (PDF).

Destino dos Dados: Planilha Excel Planilha financeira ... .xlsx, aba "Comissão".

Filtro de Exclusão: Linhas contendo o texto "PIC - Bonus Mensal" no PDF são completamente ignoradas.

Mapeamento de Colunas: Os dados são inseridos na planilha seguindo este mapeamento (PDF -> Excel):

Apl/Prop. -> Coluna B (Apólice)

Prêmio -> Coluna C (Valor)

Histórico -> Coluna E (Cliente)

Marca -> Coluna F (Seguradora, com "Porto" -> "Porto Seguro")

Fixo 12 -> Coluna G (Parcela a receber)

Parc. -> Coluna H (Parcela)

Data -> Coluna I (Dt. Pagamento)

Comissão -> Coluna J (Valor Comissão)

Taxa -> Coluna K (Porcentagem)

Fixos ("Pago", "Vida Presente", "Calina") -> Colunas L, M, N.

Agrupamento: Os dados são agrupados por Cliente, Apólice e Parcela.

Lógica de Agregação:

Valor (Prêmio): É somado (sum).

Valor Comissão: É somado (sum) e arredondado para cima (math.ceil) ao centavo mais próximo.

Porcentagem (Taxa): A maior taxa (max) do grupo é a que prevalece.

<div align="center"> <h2>📂 Estrutura do Projeto 📂</h2> </div>

O código é separado por responsabilidades para facilitar a manutenção:
```
AutomacaoHokiSeg/
├── .gitignore         # Ignora arquivos desnecessários (venv, build, dist)
├── main.py            # Ponto de entrada: Apenas inicia a aplicação
├── ui.py              # Contém toda a lógica da interface gráfica (Tkinter)
├── processing.py      # Contém toda a lógica de negócio (Pandas, PdfPlumber, OpenPyXL)
└── requirements.txt   # Lista de dependências do projeto
```
<div align="center">

<div align="center">

</div>
