import os
from utilities import carregar_dados_excel, extrair_matriculas_e_data_pdf, agrupar_por_setor, criar_pdfs_por_setor, criar_pdfs_individuais

# --- Função principal ---
def main():
    # Configurações
    arquivo_excel = "files//dados.xlsx"  # Caminho para o arquivo Excel
    arquivo_pdf = "files//arquivo.pdf"  # Caminho para o arquivo PDF
    output_dir = "files//mes//"  # Diretório de saída dos PDFs

    # Cria o diretório de saída, se não existir
    os.makedirs(output_dir, exist_ok=True)

    # Executar as etapas
    dados_matriculas = carregar_dados_excel(arquivo_excel)
    paginas_por_matricula, meses_anos_por_pagina = extrair_matriculas_e_data_pdf(arquivo_pdf, dados_matriculas.keys())
    setores_paginas, paginas_individuais = agrupar_por_setor(paginas_por_matricula, dados_matriculas)

    # Criar PDFs agrupados por setor
    criar_pdfs_por_setor(arquivo_pdf, setores_paginas, meses_anos_por_pagina, output_dir)

    # Criar PDFs individuais para setor OUT
    criar_pdfs_individuais(arquivo_pdf, paginas_individuais, meses_anos_por_pagina, output_dir)

if __name__ == "__main__":
    main()