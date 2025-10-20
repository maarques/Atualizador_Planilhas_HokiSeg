import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from datetime import datetime
import os
import math
import calendar

"""
Função principal de processamento.
Lê o PDF, processa os dados e atualiza a planilha Excel.
"""

def processar_planilha(pdf_file, planilha_file):
    
    sheet_name_to_update = "Comissão" 

    dir_name = os.path.dirname(planilha_file)
    mes_atual = calendar.month_name[datetime.now().month].lower()
    ano_atual = datetime.now().year
    novo_arquivo = os.path.join(dir_name, f"Planilha_Financeira_{mes_atual}_{ano_atual}.xlsx")

    print("ETAPA 1: Lendo PDF...")
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
                        if "PIC - Bonus Mensal" in historico: continue 
                        marca = (linha[1] or "").strip()
                        apolice = (linha[4] or "").strip()
                        parcela = (linha[6] or "").strip()
                        data_pag = (linha[8] or "").strip()
                        premio = (linha[10] or "").strip()
                        taxa = (linha[11] or "").strip()
                        comissao = (linha[12] or "").strip()
                        
                        if historico and apolice and premio and comissao and taxa and parcela and data_pag:
                            dados_extraidos.append([
                                historico, marca, apolice, parcela, data_pag, 
                                premio, taxa, comissao
                            ])
                    except (IndexError, TypeError):
                        continue

    if not dados_extraidos:
        print("⚠️ Nenhum dado foi extraído do PDF.")
        return

    print(f"Dados brutos extraídos: {len(dados_extraidos)} linhas.")

    print("ETAPA 2: Limpando e convertendo dados...")
    df = pd.DataFrame(dados_extraidos, columns=[
        'Historico', 'Marca', 'Apolice', 'Parcela', 'Data_Pag', 
        'Premio', 'Taxa', 'Comissao'
    ])

    try:
        df['Premio_float'] = pd.to_numeric(df['Premio'].str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
        df['Comissao_float'] = pd.to_numeric(df['Comissao'].str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
        df['Taxa_float'] = pd.to_numeric(df['Taxa'].str.replace(",", ".", regex=False)) / 100.0
        df['Parcela_int'] = pd.to_numeric(df['Parcela'])
        df['Data_obj'] = pd.to_datetime(df['Data_Pag'])
    except ValueError as e:
        print(f"❌ ERRO: Falha ao converter dados. Verifique o formato do PDF. Erro: {e}")
        raise
    
    df = df.dropna(subset=['Premio_float', 'Comissao_float', 'Taxa_float', 'Parcela_int', 'Data_obj'])

    print("ETAPA 3: Agrupando e somando dados...")
    group_keys = ['Historico', 'Apolice', 'Parcela_int']

    df_agrupado = df.groupby(group_keys, as_index=False).agg(
        Valor_Total=('Premio_float', 'sum'),
        Comissao_Total=('Comissao_float', 'sum'),
        Marca=('Marca', 'first'),
        Data_obj=('Data_obj', 'first'),
        Taxa_float=('Taxa_float', 'max')
    )

    df_agrupado['Comissao_Total'] = df_agrupado['Comissao_Total'].apply(lambda x: math.ceil(x * 100) / 100)
    
    print(f"Dados lidos e agrupados. {len(df_agrupado)} linhas únicas serão inseridas.")

    print("ETAPA 4: Preparando colunas finais...")
    df_agrupado = df_agrupado.rename(columns={
        'Historico': 'Cliente', 'Apolice': 'Apólice', 'Parcela_int': 'Parcela',
        'Data_obj': 'Dt. Pagamento', 'Taxa_float': 'Porcentagem',
        'Valor_Total': 'Valor', 'Comissao_Total': 'Valor Comissão'
    })

    df_agrupado['Seguradora'] = df_agrupado['Marca'].apply(lambda x: "Porto Seguro" if "Porto" in x else x)
    df_agrupado['Parcela a receber'] = 12
    df_agrupado['Situação'] = 'Pago'
    df_agrupado['Tipo'] = 'Vida Presente'
    df_agrupado['Colaborador'] = 'Calina'

    print(f"ETAPA 5: Abrindo a planilha '{os.path.basename(planilha_file)}'...")
    try:
        wb = load_workbook(planilha_file)
    except FileNotFoundError:
        print(f"❌ ERRO: Arquivo '{planilha_file}' não encontrado. Verifique o nome.")
        raise

    try:
        ws = wb[sheet_name_to_update]
    except KeyError:
        print(f"❌ ERRO: A aba '{sheet_name_to_update}' não foi encontrada no arquivo Excel.")
        print(f"Abas disponíveis: {wb.sheetnames}")
        raise

    first_empty_row = 5 
    while ws.cell(row=first_empty_row, column=2).value is not None:
        first_empty_row += 1

    print(f"Inserindo {len(df_agrupado)} novas linhas a partir da linha {first_empty_row}...")

    colunas_para_escrever = {
        2: 'Apólice', 3: 'Valor', 5: 'Cliente', 6: 'Seguradora',
        7: 'Parcela a receber', 8: 'Parcela', 9: 'Dt. Pagamento',
        10: 'Valor Comissão', 11: 'Porcentagem', 12: 'Situação',
        13: 'Tipo', 14: 'Colaborador'
    }

    for _, row in df_agrupado.iterrows():
        for col_idx, col_name in colunas_para_escrever.items():
            ws.cell(row=first_empty_row, column=col_idx, value=row[col_name])
        
        ws.cell(row=first_empty_row, column=3).number_format = '"R$" #,##0.00'
        ws.cell(row=first_empty_row, column=9).number_format = 'DD/MM/YYYY'
        ws.cell(row=first_empty_row, column=10).number_format = '"R$" #,##0.00'
        ws.cell(row=first_empty_row, column=11).number_format = '0.00%'
        first_empty_row += 1 

    print(f"Salvando como '{os.path.basename(novo_arquivo)}'...")
    wb.save(novo_arquivo)

    print(f"\n✅ Planilha atualizada com sucesso: {novo_arquivo}")
    print(f"➕ Foram adicionadas {len(df_agrupado)} linhas (agrupadas e somadas):\n")

    print(f"{'Apólice':<15} | {'Cliente':<30} | {'Parc.':>5} | {'Valor':>10} | {'V. Comissão':>10} | {'Taxa':>6}")
    print("-" * 85)
    for _, row in df_agrupado.iterrows():
        print(f"{row['Apólice']:<15} | {row['Cliente']:<30} | {row['Parcela']:>5} | {row['Valor']:>10.2f} | {row['Valor Comissão']:>10.2f} | {row['Porcentagem']:>6.2%}")