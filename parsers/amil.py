import pdfplumber
import re
from datetime import datetime

def parse(pdf_file):
    """
    Analisa um PDF de extrato da Amil.
    Usa RegEx para extrair dados, pois o PDF não é tabular.
    Retorna os dados no formato padronizado (lista de listas).
    """
    print("Iniciando parser da Amil...")
    dados_extraidos = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            texto_pagina = page.extract_text()
            
            if not texto_pagina:
                continue
            
            # Encontrar a Data de Pagamento
            padrao_data = r"Data pagamento previsto: (\d{2})/(\d{2})/(\d{4})"
            match_data = re.search(padrao_data, texto_pagina)
            
            if not match_data:
                print("⚠️ Aviso: Não foi possível encontrar a 'Data pagamento previsto' na página. Pulando.")
                continue

            dia, mes, ano = match_data.groups()
            data_pag_str = f"{ano}-{mes}-{dia}"
            
            # Divide o texto da página em blocos, um para cada "Contrato:"
            blocos_contrato = texto_pagina.split("Contrato:")[1:]
            
            if not blocos_contrato:
                print("⚠️ Aviso: Nenhuma seção 'Contrato:' encontrada na página.")
                continue

            print(f"Encontrados {len(blocos_contrato)} blocos de contrato na página...")

            # Iterar e extrair dados de cada Bloco ---
            for bloco in blocos_contrato:
                try:
                    padrao_contrato = r"^\s*(\d{10})\s*-\s*(?:[\d\s]+)?([A-Za-zÀ-ÿ].*?)(?=Proposta:|$)"
                    match_contrato = re.search(padrao_contrato, bloco) 
                    
                    apolice = match_contrato.group(1).strip()
                    cliente = match_contrato.group(2).strip()

                    # Parcela e Valor (Prêmio)
                    padrao_parcela_premio = r"\d+\s+(\d+)\s+\d{2}/\d{4}\s+([\d.,]+)"
                    match_pp = re.search(padrao_parcela_premio, bloco)
                    parcela = match_pp.group(1).strip()
                    premio = match_pp.group(2).strip()

                    # Taxa (Porcentagem) e Valor Comissão
                    padrao_taxa_comissao = r"([\d,]+)\s*%\s*[\d.,]+\s+([\d.,]+)"
                    match_tc = re.search(padrao_taxa_comissao, bloco)
                    taxa = match_tc.group(1).strip()
                    comissao = match_tc.group(2).strip()

                    # Montar a lista padronizada                 
                    dados_extraidos.append([
                        cliente,
                        "Amil",
                        apolice,
                        parcela,
                        data_pag_str,
                        premio,
                        taxa,
                        comissao
                    ])

                except AttributeError as e:
                    print(f"⚠️ Aviso: Falha ao analisar um bloco de contrato. Pode estar mal formatado. Erro: {e}")
                    continue

    print(f"Parser da Amil finalizado. {len(dados_extraidos)} linhas extraídas.")
    return dados_extraidos
