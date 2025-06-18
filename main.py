from meds_analyzer.processor import processar_meds
from meds_analyzer.gerar_base_transacional import gerar_base_transacional
from meds_analyzer.analyzer import gerar_analise_geral_e_diaria

if __name__ == "__main__":
    print("🔄 Iniciando processamento dos MEDs...")
    processar_meds(
        caminho_csv="data/meds.csv",
        caminho_excel_saida="output/MEDS_Separados.xlsx"
    )

    print("📊 Gerando base transacional consolidada...")
    gerar_base_transacional(
        metabase_path="data/metabase_pix.csv",
        clientes_path="data/Base Clientes - Microcash (MED).xlsx",
        output_path="output/Base_transacional.xlsx"
    )

    print("📈 Gerando análise geral e diária dos MEDs...")
    gerar_analise_geral_e_diaria(
        caminho_meds="output/MEDS_Separados.xlsx",
        caminho_transacional="output/Base_transacional.xlsx",
        caminho_clientes="data/Base Clientes - Microcash (MED).xlsx",
        saida_geral="output/MEDS-Analise Geral.xlsx",
        saida_diaria="output/Analise_diaria.xlsx",
        intervalo_dias=3
    )

    print("✅ Tudo pronto! MEDs e base transacional atualizados com sucesso.")