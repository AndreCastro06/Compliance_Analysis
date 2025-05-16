import pandas as pd
import os

def atualizar_meds_sem_cnpj(df_sem_cnpj, coluna_id_med, output_path):
    if df_sem_cnpj.empty:
        return

    try:
        if os.path.exists(output_path):
            df_existente = pd.read_excel(output_path, sheet_name="Sem_CNPJ", dtype=str)
            df_combinado = pd.concat([df_existente, df_sem_cnpj], ignore_index=True)
            df_combinado.drop_duplicates(subset=[coluna_id_med], inplace=True)
        else:
            df_combinado = df_sem_cnpj

        df_combinado.to_excel(output_path, sheet_name="Sem_CNPJ", index=False)
        print("✅ MEDs sem CNPJ atualizados com sucesso.")
    except Exception as e:
        print(f"⚠️ Erro ao salvar MEDs sem CNPJ: {e}")