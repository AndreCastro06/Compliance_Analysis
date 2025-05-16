import pandas as pd
import os
from openpyxl import load_workbook

def gerar_base_transacional(metabase_path: str, clientes_path: str, output_path: str):
    print("📊 Gerando base transacional consolidada...")

    df_metabase = pd.read_csv(metabase_path, sep=',', decimal=',', dtype=str)
    df_metabase.columns = df_metabase.columns.str.strip()

    print("Colunas disponíveis no metabase_pix.csv:")
    print(list(df_metabase.columns))

    df_metabase["Transactions"] = df_metabase["Transactions"].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    df_metabase["Transactions"] = pd.to_numeric(df_metabase["Transactions"], errors='coerce').fillna(0)

    df_metabase["Sum of Amount"] = df_metabase["Sum of Amount"].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    df_metabase["Sum of Amount"] = pd.to_numeric(df_metabase["Sum of Amount"], errors='coerce').fillna(0)

    df_metabase["Created At: Day"] = pd.to_datetime(df_metabase["Created At: Day"], errors='coerce').dt.date

    df_clientes = pd.read_excel(clientes_path, dtype=str)
    df_clientes["Conta Numero"] = df_clientes["Conta Numero"].str.strip()
    df_clientes["Documento"] = df_clientes["Documento"].str.strip()

    analises = []

    dados_existentes = {}
    if os.path.exists(output_path):
        print("📂 Arquivo existente encontrado. Carregando dados anteriores...")
        with pd.ExcelFile(output_path) as reader:
            for aba in reader.sheet_names:
                df_antigo = pd.read_excel(reader, sheet_name=aba)
                df_antigo["Created At: Day"] = pd.to_datetime(df_antigo["Created At: Day"]).dt.date
                dados_existentes[aba] = df_antigo

    for documento, grupo in df_clientes.groupby("Documento"):
        contas = grupo["Conta Numero"].dropna().unique()
        df_cliente = df_metabase[df_metabase["Account Key"].isin(contas)]

        if df_cliente.empty:
            print(f"⚠️ Nenhuma aba de transações para {documento}. A análise será feita apenas com os MEDs.")
            continue

        df_cliente_agrupado = df_cliente.groupby("Created At: Day", as_index=False).agg({
            "Transactions": "sum",
            "Sum of Amount": "sum"
        })

        nome_aba = str(documento)[:31]

        if nome_aba in dados_existentes:
            df_antigo = dados_existentes[nome_aba]
            # Filtra apenas datas novas
            datas_existentes = set(df_antigo["Created At: Day"])
            df_novos = df_cliente_agrupado[~df_cliente_agrupado["Created At: Day"].isin(datas_existentes)]
            df_cliente_agrupado = pd.concat([df_antigo, df_novos], ignore_index=True)
            df_cliente_agrupado = df_cliente_agrupado.sort_values("Created At: Day")

        analises.append((nome_aba, df_cliente_agrupado))

    # Salva no Excel, substituindo só as abas necessárias
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
        for nome_aba, df in analises:
            df.to_excel(writer, sheet_name=nome_aba, index=False)

    print("✅ Base transacional consolidada com abas por cliente foi atualizada com sucesso.")