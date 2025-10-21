# Atualizador de Planilhas HokiSeg  
![Python](https://img.shields.io/badge/python-3.8%2B-blue) ![Pandas](https://img.shields.io/badge/pandas-1.0%2B-blueviolet) ![OpenPyXL](https://img.shields.io/badge/openpyxl-3.0%2B-green) ![Tkinter](https://img.shields.io/badge/tkinter-GUI-orange)  

Uma ferramenta com **interface gráfica** para ler analíticos de pagamento de comissões em PDF (inicialmente para Porto Seguro) e inseri-los de forma consolidada numa planilha Excel de controle financeiro da HokiSeg.

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
   - Clique em **„3. Processar e Atualizar Planilha”**.  
3. Aguarde a barra de log exibir a mensagem de sucesso.  
4. Pronto! Um novo arquivo será salvo (ex: `Planilha_financeira_out-2025_ATUALIZADA.xlsx`) na mesma pasta da planilha original.

---

## 🛠️ Como Executar (para desenvolvedores)  
Se você quiser rodar o código-fonte e fazer melhorias:  
```bash
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
Para gerar o executável (.exe no Windows)
bash
Copiar código
pyinstaller --onefile --noconsole --name="AutomacaoHokiSeg" main.py
O arquivo AutomacaoHokiSeg.exe será criado na pasta dist/.

⚙️ Regras de Negócio Implementadas
Fonte de dados: extrato analítico de pagamentos de comissões da Porto Seguro (formato PDF).

Destino dos dados: planilha Excel (Planilha financeira … .xlsx), aba “Comissão”.

Filtro de exclusão: linhas contendo o texto “PIC – Bonus Mensal” são ignoradas.

Mapeamento de colunas (PDF → Excel):

Apl/Prop. → Coluna B (Apólice)

Prêmio → Coluna C (Valor)

Histórico → Coluna E (Cliente)

Marca → Coluna F (Seguradora)

Parcela a Receber (fixo = 12) → Coluna G

Parc. → Coluna H

Data → Coluna I (Dt. Pagamento)

Comissão → Coluna J (Valor Comissão)

Taxa → Coluna K (Porcentagem)

Fixos (“Pago”, “Vida Presente”, “Calina”) → Colunas L, M, N

Agrupamento: os dados são agrupados por Cliente + Apólice + Parcela.

Valor (Prêmio): somado (sum).

Valor Comissão: somado (sum) e arredondado para cima ao centavo.

Porcentagem (Taxa): prevalece a maior taxa do grupo (max).

🧩 Estrutura do Projeto
```bash
Copiar código
Atualizador_Planilhas_HokiSeg/
├── .gitignore           # Ignora arquivos da build, venv, dist etc.
├── main.py              # Ponto de entrada: inicia a aplicação
├── ui.py                # Interface gráfica (Tkinter)
├── processing.py        # Lógica de negócio: pandas, pdfplumber, openpyxl
└── requirements.txt     # Dependências do projeto
✅ Contribuições & Melhorias Futuras
Contribuições são bem-vindas! Algumas ideias para evolução:

Suporte a outros formatos de analítico de seguradoras diferentes da Porto Seguro e Amil.

Reconhecimento automático de colunas em PDF com layout variável.

Tradução/localização para outros idiomas.

Interface web para upload de arquivos e processamento online.

Versão multiplataforma (Windows + macOS + Linux) empacotada.

