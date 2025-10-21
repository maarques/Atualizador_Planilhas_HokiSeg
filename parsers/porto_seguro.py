import pdfplumber

def parse(pdf_file):
    """
    Analisa um PDF de extrato da Porto Seguro.
    Extrai os dados brutos e os retorna como uma lista de listas.
    """
    print("Iniciando parser da Porto Seguro...")
    dados_extraidos = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tabelas = page.extract_tables()
            if not tabelas: continue
            
            for tabela in tabelas:
                for linha in tabela:
                    if not linha or not linha[0] or linha[0] in [None, "Histórico", "Total Comissão Bruta - Susep Produção: 6C7NTJ"]:
                        continue
                    
                    try:
                        historico = (linha[0] or "").strip()
                        # Filtro específico da Porto
                        if "PIC - Bonus Mensal" in historico:
                            continue 
                            
                        marca = (linha[1] or "").strip()
                        apolice = (linha[4] or "").strip()
                        parcela = (linha[6] or "").strip()
                        data_pag = (linha[8] or "").strip()
                        premio = (linha[10] or "").strip()
                        taxa = (linha[11] or "").strip()
                        comissao = (linha[12] or "").strip()
                        
                        # Validação
                        if historico and apolice and premio and comissao and taxa and parcela and data_pag:
                            dados_extraidos.append([
                                historico, marca, apolice, parcela, data_pag, 
                                premio, taxa, comissao
                            ])
                    except (IndexError, TypeError):
                        continue
                        
    if not dados_extraidos:
        print("⚠️ Nenhum dado foi extraído pelo parser da Porto Seguro.")
    
    print(f"Parser da Porto Seguro finalizado. {len(dados_extraidos)} linhas extraídas.")
    return dados_extraidos
