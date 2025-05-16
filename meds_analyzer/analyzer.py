import pandas as pd
import os
from datetime import datetime, timedelta
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl import load_workbook
import re

def formatar_cabecalhos(path_excel):
    wb = load_workbook(path_excel)

    fill_azul = PatternFill(start_color='FF010AFD', end_color='FF010AFD', fill_type='solid')
    fonte_branca_negrito = Font(color='FFFFFFFF', bold=True)
    alinhamento_central = Alignment(horizontal='center', vertical='center', wrap_text=True)
    fonte_vermelha_negrito = Font(color='FF0000', bold=True)

    for aba in wb.sheetnames:
        ws = wb[aba]

        for cell in ws[1]:
            cell.fill = fill_azul
            cell.font = fonte_branca_negrito
            cell.alignment = alinhamento_central

        headers = [cell.value for cell in ws[1]]
        try:
            col_med_amount = headers.index('% Valor MEDs x Pix-In (Dia)') + 1
            col_acumulado = headers.index('% Valor MEDs x Pix-In (Semana)') + 1
        except ValueError:
            continue  # pula a aba se não encontrar as colunas necessárias

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            try:
                valor_raw_1 = row[col_med_amount - 1].value
                if isinstance(valor_raw_1, str):
                    valor_1 = float(re.sub(r'[^\d.,]', '', valor_raw_1).replace(',', '.')) / 100 if '%' in valor_raw_1 else float(valor_raw_1.replace(',', '.'))
                else:
                    valor_1 = float(valor_raw_1)
                if valor_1 > 0.0035:
                    row[col_med_amount - 1].font = fonte_vermelha_negrito
            except:
             pass

        try:
            valor_raw_2 = row[col_acumulado - 1].value
            if isinstance(valor_raw_2, str):
                valor_2 = float(re.sub(r'[^\d.,]', '', valor_raw_2).replace(',', '.')) / 100 if '%' in valor_raw_2 else float(valor_raw_2.replace(',', '.'))
            else:
                valor_2 = float(valor_raw_2)
            if valor_2 > 0.0035:
                row[col_acumulado - 1].font = fonte_vermelha_negrito
        except:

            try:
                valor_raw_2 = row[col_acumulado - 1].value
                if isinstance(valor_raw_2, str):
                    valor_2 = float(re.sub(r'[^\d.,]', '', valor_raw_2).replace(',', '.')) / 100 if '%' in valor_raw_2 else float(valor_raw_2.replace(',', '.'))
                else:
                    valor_2 = float(valor_raw_2)
                if valor_2 > 0.0035:
                    row[col_acumulado - 1].font = fonte_vermelha_negrito
            except:
                pass

    wb.save(path_excel)
    print(f" Estilização aplicada: {path_excel}")

# Cnverter valores monetários no padrão brasileiro
def converter_valor_brasileiro(valor):
    if pd.isna(valor):
        return 0.0
    valor_str = str(valor).strip()
    valor_str = re.sub(r'[^\d,.-]', '', valor_str)  # remove caracteres estranhos
    if ',' in valor_str and '.' in valor_str:
        # Ex: 1.234,56 → 1234.56
        valor_str = valor_str.replace('.', '').replace(',', '.')
    elif ',' in valor_str:
        # Ex: 1234,56 → 1234.56
        valor_str = valor_str.replace(',', '.')
    return float(valor_str)

def formatar_reais(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

def inserir_percentual_acumulado_mensal(df_analise):
    # Garantir datetime na coluna 'Data'
    df_analise["Data"] = pd.to_datetime(df_analise["Data"])
    df_analise["Ano"] = df_analise["Data"].dt.year
    df_analise["Mes"] = df_analise["Data"].dt.month

    # Converter colunas para número
    df_analise["Transacoes_Convertido"] = pd.to_numeric(df_analise["Qtd. Transações"], errors='coerce').fillna(0)
    df_analise["MEDs_Convertido"] = pd.to_numeric(df_analise["Qtd. MEDs"], errors='coerce').fillna(0)

    # Acumulados mês a mês
    df_analise["Acum_PIX"] = df_analise.groupby(["Ano", "Mes"])["Transacoes_Convertido"].cumsum()
    df_analise["Acum_MEDs"] = df_analise.groupby(["Ano", "Mes"])["MEDs_Convertido"].cumsum()

    # Percentual acumulado do mês
    df_analise["Percentual Abertura MEDs x Transações - Mês"] = df_analise.apply(
        lambda row: f"{(row['Acum_MEDs'] / row['Acum_PIX'] * 100):.2f}%" if row["Acum_PIX"] != 0 else "0,00%",
        axis=1
    )

    # Limpeza
    df_analise.drop(columns=[
        "Ano", "Mes", "Transacoes_Convertido", "MEDs_Convertido",
        "Acum_PIX", "Acum_MEDs"
    ], inplace=True)

    return df_analise

def gerar_analise_geral_e_diaria( caminho_meds="SEU_CAMINHO/MEDS_base.xlsx", caminho_transacional="SEU_CAMINHO/base_transacional.xlsx", caminho_clientes="SEU_CAMINHO/clientes.xlsx", saida_geral="MEDS - Analise Geral.xlsx",
    saida_diaria="MEDS - Report Diário.xlsx", intervalo_dias=3 ):
    abas_meds = pd.ExcelFile(caminho_meds).sheet_names
    abas_pix = pd.ExcelFile(caminho_transacional).sheet_names

    df_clientes = pd.read_excel(caminho_clientes, dtype=str)
    df_clientes["Documento"] = df_clientes["Documento"].str.replace(r'\D', '', regex=True)
    mapa_cnpj_para_nome = dict(zip(df_clientes["Documento"], df_clientes["Pessoa Nome"]))

    todas_analises = []
    for aba in abas_meds:
        cnpj_atual = aba.strip()
        nome_empresa = mapa_cnpj_para_nome.get(cnpj_atual, f"Empresa_{cnpj_atual}")[:31]
        try:
            df_meds = pd.read_excel(caminho_meds, sheet_name=aba, dtype=str)
            df_meds["ValorTransacao"] = pd.to_numeric(df_meds["ValorTransacao"].str.replace(',', '.'), errors='coerce').fillna(0)
            df_meds["Created At: Day"] = pd.to_datetime(df_meds["DtHrCriacaoNotificacaoInfracao"]).dt.date
        except Exception as e:
            print(f"Erro ao ler MEDs da aba {aba}: {e}")
            continue

        aba_transacional = cnpj_atual if cnpj_atual in abas_pix else None

        if aba_transacional:
            try:
                df_pix = pd.read_excel(caminho_transacional, sheet_name=aba_transacional, dtype=str)
                df_pix["Transactions"] = pd.to_numeric(df_pix["Transactions"].str.replace('.', '').str.replace(',', '.'), errors='coerce').fillna(0).astype(int)
                # df_pix["Sum of Amount"] = pd.to_numeric(df_pix["Sum of Amount"].str.replace(',', '.', regex=False).str.replace('.', '', regex=False), errors='coerce').fillna(0).astype(float).round(2)
                df_pix["Sum of Amount"] = df_pix["Sum of Amount"].apply(converter_valor_brasileiro).fillna(0).round(2)
                df_pix["Data"] = pd.to_datetime(df_pix["Created At: Day"]).dt.date
            except Exception as e:
                print(f"Erro ao ler transações da aba {aba_transacional}: {e}")
                df_pix = pd.DataFrame(columns=["Data", "Transactions", "Sum of Amount"])
        else:
            print(f"⚠️ Nenhuma aba de transações para {aba}. A análise será feita apenas com os MEDs.")
            df_pix = pd.DataFrame(columns=["Data", "Transactions", "Sum of Amount"])

        analise_diaria = []
        datas_meds = set(df_meds["Created At: Day"].unique())
        datas_pix = set(df_pix["Data"].unique()) if not df_pix.empty else set()
        todas_datas = sorted(datas_meds.union(datas_pix))
        for data in todas_datas:
            data_date = pd.to_datetime(data).date()
            meds_dia = df_meds[df_meds["Created At: Day"] == data_date]
            pix_dia = df_pix[df_pix["Data"] == data_date] if not df_pix.empty else pd.DataFrame(columns=["Transactions", "Sum of Amount"])

            total_meds = meds_dia["ValorTransacao"].sum()
            qtd_meds = len(meds_dia)
            meds_menor_500 = meds_dia[meds_dia["ValorTransacao"] < 500]
            meds_maior_500 = meds_dia[meds_dia["ValorTransacao"] >= 500]

            total_pix = pix_dia["Sum of Amount"].sum()
            qtd_pix = pix_dia["Transactions"].sum()

            meds_acumulado = df_meds[(df_meds["Created At: Day"] >= data_date - timedelta(days=6)) & (df_meds["Created At: Day"] <= data_date)]["ValorTransacao"].sum()
            pix_acumulado = df_pix[(df_pix["Data"] >= data_date - timedelta(days=6)) & (df_pix["Data"] <= data_date)]["Sum of Amount"].sum()

            acumulado_percentual = f"{(meds_acumulado / pix_acumulado * 100):.2f}%" if pix_acumulado else "0,00%"

            # ✅ NOVO BLOCO PARA ACUMULADO MENSAL
            primeiro_dia_mes = data_date.replace(day=1)
            df_meds_mes = df_meds[(df_meds["Created At: Day"] >= primeiro_dia_mes) & (df_meds["Created At: Day"] <= data_date)]
            df_pix_mes = df_pix[(df_pix["Data"] >= primeiro_dia_mes) & (df_pix["Data"] <= data_date)]
            valor_meds_mes = df_meds_mes["ValorTransacao"].sum()
            valor_pix_mes = df_pix_mes["Sum of Amount"].sum()
            percentual_acumulado_mes_valor = f"{(valor_meds_mes / valor_pix_mes * 100):.2f}%" if valor_pix_mes else "0,00%"

            linha = {
                    "Data": data_date,
                    "Qtd. Transações": int(qtd_pix),
                    "Valor Pix-In": formatar_reais(total_pix),
                    "Qtd. MEDs": qtd_meds,
                    "Valor MEDs": formatar_reais(total_meds),
                    "MEDs < R$ 500": len(meds_menor_500),
                    "% MEDs < R$ 500": f"{(meds_menor_500['ValorTransacao'].sum() / total_meds * 100):.2f}%" if total_meds else "0,00%",
                    "MEDs >= R$ 500": len(meds_maior_500),
                    "% MEDs >= R$ 500": f"{(len(meds_maior_500) / qtd_meds * 100):.2f}%" if qtd_meds else "0,00%",
                    "% MEDs x Pix-In (Qtd)": f"{(qtd_meds / qtd_pix * 100):.2f}%" if qtd_pix else "0,00%",
                    "% Valor MEDs x Pix-In (Dia)": f"{(total_meds / total_pix * 100):.2f}%" if total_pix else "0,00%",
                    "% Valor MEDs x Pix-In (Semana)": acumulado_percentual,
                    "% Valor MEDs x Pix-In (Mês)": percentual_acumulado_mes_valor
                }
            analise_diaria.append(linha)

        df_analise = pd.DataFrame(analise_diaria)
        df_analise = inserir_percentual_acumulado_mensal(df_analise)

        # Reorganiza colunas
        ordem_colunas = [
                        "Data",
                        "Qtd. Transações",
                        "Valor Pix-In",
                        "Qtd. MEDs",
                        "Valor MEDs",
                        "MEDs < R$ 500",
                        "% MEDs < R$ 500",
                        "MEDs >= R$ 500",
                        "% MEDs >= R$ 500",
                        "% MEDs x Pix-In (Qtd)",
                        "% Valor MEDs x Pix-In (Dia)",
                        "% Valor MEDs x Pix-In (Semana)",
                        "% Valor MEDs x Pix-In (Mês)"
            ]
        ordem_existente = [col for col in ordem_colunas if col in df_analise.columns]
        df_analise = df_analise[ordem_existente]
        todas_analises.append((nome_empresa, df_analise))

    # Garante que o relatório geral tenha linhas zeradas mesmo que sem MEDs ou PIX
    datas_completas = set()
    for _, df_empresa in todas_analises:
        datas_completas.update(df_empresa["Data"].unique())

    if not datas_completas:
        datas_completas = {datetime.today().date()}

    for i, (nome_empresa, df_empresa) in enumerate(todas_analises):
        datas_existentes = set(df_empresa["Data"].unique())
        datas_faltando = datas_completas - datas_existentes

        linhas_zeradas = []
        for data_faltante in sorted(datas_faltando):
            linhas_zeradas.append({
              "Data": data_faltante,
              "Qtd. Transações": 0,
              "Valor Pix-In": 0,
              "Qtd. MEDs": 0,
              "Valor MEDs": 0,
              "MEDs < R$ 500": 0,
              "% MEDs < R$ 500": "0,00%",
              "MEDs >= R$ 500": 0,
              "% MEDs >= R$ 500": "0,00%",
              "% MEDs x Pix-In (Qtd)": "0,00%",
              "% Valor MEDs x Pix-In (Dia)": "0,00%",
              "% Valor MEDs x Pix-In (Semana)": "0,00%",
              "% Valor MEDs x Pix-In (Mês)": "0,00%"
            })
        if linhas_zeradas:
            todas_analises[i] = (nome_empresa, pd.concat([df_empresa, pd.DataFrame(linhas_zeradas)], ignore_index=True))

    # Grava o relatório geral considerando se ja foi feito
    dados_existentes = {}
    if os.path.exists(saida_geral):
        print("📂 Carregando dados existentes do arquivo geral...")
        with pd.ExcelFile(saida_geral) as reader:
            for aba in reader.sheet_names:
                df_antigo = pd.read_excel(reader, sheet_name=aba, dtype=str)
                df_antigo["Data"] = pd.to_datetime(df_antigo["Data"]).dt.date
                dados_existentes[aba] = df_antigo

    with pd.ExcelWriter(saida_geral, engine='openpyxl', mode='w') as writer:
        for nome_empresa, df_novo in todas_analises:
            if nome_empresa in dados_existentes:
                df_antigo = dados_existentes[nome_empresa]
                df_novo["Data"] = pd.to_datetime(df_novo["Data"]).dt.date
                datas_antigas = set(df_antigo["Data"])
                df_novos_dias = df_novo[~df_novo["Data"].isin(datas_antigas)]
                df_combinado = pd.concat([df_antigo, df_novos_dias], ignore_index=True)
                df_combinado = df_combinado.sort_values("Data")
            else:
                df_combinado = df_novo.sort_values("Data")
            
            df_combinado["Data"] = pd.to_datetime(df_combinado["Data"]).dt.strftime("%d/%m/%Y")

            df_combinado.to_excel(writer, sheet_name=nome_empresa, index=False)

    # Pergunta e gera relatório diário
    print("Deseja gerar análise dos últimos 3 dias ou apenas de ontem?")
    print("Digite 1 para últimos 3 dias ou 2 para apenas ontem:")
    escolha = input().strip()

    if escolha == "1":
        dias = [datetime.today().date() - timedelta(days=i) for i in range(1, 4)]
    else:
        dias = [datetime.today().date() - timedelta(days=1)]

    analises_diarias = [
    (nome_empresa, df[df["Data"].isin(dias)])
    for nome_empresa, df in todas_analises
    if not df[df["Data"].isin(dias)].empty
]

    if analises_diarias:
        with pd.ExcelWriter(saida_diaria, engine='openpyxl', mode='w') as writer_diaria:
            for nome_empresa, df_filtrado in analises_diarias:
                nome_aba_seguro = nome_empresa[:31]
                
                # Tratativa de formatação de data
                df_filtrado["Data"] = pd.to_datetime(df_filtrado["Data"]).dt.strftime("%d/%m/%Y")

                df_filtrado.to_excel(writer_diaria, sheet_name=nome_aba_seguro, index=False)

        formatar_cabecalhos(saida_diaria)
        print(f"\n📄 Relatório diário gerado com sucesso: {saida_diaria}")
    else:
        print("⚠️ Nenhuma empresa com dados novos para os dias selecionados.Sem necessidade de gerar novo relatório.")

    formatar_cabecalhos(saida_geral)
    formatar_cabecalhos(saida_diaria)
    print("\n✅ Arquivos 'MEDS - Analise Geral.xlsx' e 'MEDS - Report Diário.xlsx' gerados com sucesso.")
