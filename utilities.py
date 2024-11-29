import pandas as pd
import PyPDF2

# Lista de meses para identificação no PDF
MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

# --- Passo 1: Ler a base de dados em Excel ---
def carregar_dados_excel(arquivo_excel):
    df = pd.read_excel(arquivo_excel)  # Lê o arquivo Excel
    # Cria um dicionário {matrícula: {"setor": setor, "nome": nome}}
    dados = {row["Matrícula"]: {"setor": row["Setor"], "nome": row["Nome"]} for _, row in df.iterrows()}
    return dados

# --- Passo 2: Ler as páginas do PDF e identificar matrículas, mês e ano ---
def extrair_matriculas_e_data_pdf(arquivo_pdf, matriculas):
    pdf_reader = PyPDF2.PdfReader(arquivo_pdf)
    paginas_por_matricula = {}
    meses_anos_por_pagina = {}

    for num_pagina, pagina in enumerate(pdf_reader.pages):
        texto = pagina.extract_text()  # Extrai o texto da página

        # Identificar matrículas
        for matricula in matriculas:
            if str(matricula) in texto:  # Verifica se a matrícula está no texto
                if matricula not in paginas_por_matricula:
                    paginas_por_matricula[matricula] = []
                paginas_por_matricula[matricula].append(num_pagina)

        # Identificar mês/ano
        for mes in MESES:
            if mes in texto:
                # Encontrar o formato [Mês]/[Ano] no texto
                palavras = texto.split()
                for palavra in palavras:
                    if mes in palavra and "/" in palavra:  # Exemplo: "Outubro/2024"
                        meses_anos_por_pagina[num_pagina] = palavra
                        break

    return paginas_por_matricula, meses_anos_por_pagina

# --- Passo 3: Agrupar páginas por setor ---
def agrupar_por_setor(paginas_por_matricula, dados_matriculas):
    setores_paginas = {}
    paginas_individuais = {}

    for matricula, paginas in paginas_por_matricula.items():
        setor = dados_matriculas.get(matricula, {}).get("setor")
        nome = dados_matriculas.get(matricula, {}).get("nome")

        if setor == "OUTROS":
            # Para setor OUT, salva individualmente pelo nome
            for pagina in paginas:
                paginas_individuais[(matricula, nome)] = pagina
        else:
            # Para outros setores, agrupa normalmente
            if setor not in setores_paginas:
                setores_paginas[setor] = []
            setores_paginas[setor].extend(paginas)

    return setores_paginas, paginas_individuais

# --- Passo 4: Criar PDFs agrupados por setor com mês e ano ---
def criar_pdfs_por_setor(arquivo_pdf, setores_paginas, meses_anos_por_pagina, output_dir):
    pdf_reader = PyPDF2.PdfReader(arquivo_pdf)

    for setor, paginas in setores_paginas.items():
        pdf_writer = PyPDF2.PdfWriter()
        mes_ano = "Desconhecido"

        # Adiciona páginas e busca o primeiro mês/ano associado ao setor
        for num_pagina in sorted(paginas):
            pdf_writer.add_page(pdf_reader.pages[num_pagina])
            if num_pagina in meses_anos_por_pagina and mes_ano == "Desconhecido":
                mes_ano = meses_anos_por_pagina[num_pagina]  # Usa o primeiro mês/ano encontrado

        # Corrige nome do arquivo para evitar caracteres inválidos
        mes_ano = mes_ano.replace("/", "_")  # Ex.: "Outubro/2024" -> "Outubro_2024"
        arquivo_saida = f"{output_dir}/Contracheques_{mes_ano}_{setor}.pdf"

        # Salvar o PDF para o setor
        try:
            with open(arquivo_saida, "wb") as pdf_saida:
                pdf_writer.write(pdf_saida)
            print(f"PDF criado para o setor {setor} (mês/ano: {mes_ano}): {arquivo_saida}")
        except Exception as e:
            print(f"Erro ao criar o arquivo {arquivo_saida}: {e}")
            raise

# --- Passo 5: Criar PDFs individuais para setor OUT ---
def criar_pdfs_individuais(arquivo_pdf, paginas_individuais, meses_anos_por_pagina, output_dir):
    pdf_reader = PyPDF2.PdfReader(arquivo_pdf)

    for (matricula, nome), pagina in paginas_individuais.items():
        pdf_writer = PyPDF2.PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[pagina])

        # Identificar o mês/ano da página
        mes_ano = meses_anos_por_pagina.get(pagina, "Desconhecido")

        mes_ano = mes_ano.replace("/", "_")  # Ex.: "Outubro/2024" -> "Outubro_2024"

        # Salvar o PDF com o nome correspondente
        arquivo_saida = f"{output_dir}/Contracheque_{mes_ano}_{nome}_{matricula}.pdf"
        try:
            with open(arquivo_saida, "wb") as pdf_saida:
                pdf_writer.write(pdf_saida)
            print(f"PDF criado para matrícula {matricula} (nome: {nome}, data: {mes_ano}): {arquivo_saida}")
        except Exception as e:
            print(f"Erro ao criar o arquivo {arquivo_saida}: {e}")
            raise