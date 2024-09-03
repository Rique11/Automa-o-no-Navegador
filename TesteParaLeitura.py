import os
import re
import PyPDF2

# Função para encontrar a quantidade com base nos termos chave
def encontrar_quantidade(linha, termos_chave):
    # Expressão regular melhorada que busca o termo chave seguido de um número no formato decimal
    padrao_geral = r'({})\s*([\d]+[.,]\d+)'.format('|'.join(re.escape(termo) for termo in termos_chave))
    matches = re.findall(padrao_geral, linha)
    
    quantidades = []
    for match in matches:
        termo_chave, quantidade = match
        # Verifica se o termo chave não é parte de outra palavra (evita 'SALA' ser confundido com 'LA')
        if re.search(r'\b{}\b'.format(re.escape(termo_chave)), linha):
            quantidades.append(quantidade.replace(',', '.'))  # Troca vírgula por ponto se necessário
            print(linha)
            input("")
    return quantidades if quantidades else None

def processar_danfe(pdf_path): 
    print(f"Lendo DANFE: {pdf_path}")
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()

    linhas = text.splitlines()

    chave_acesso = None
    numeroNF = None
    numeroSerie = None
    quantidade = []

    # Lista de termos chave que podem indicar a quantidade
    termos_chave = ['PACOTE', 'LA', 'UN', 'Un', 'un', 'peca', 'PC', 'CX', 'Pc', 'pc', 'MT', 'JG', 'KIT', 'UND', 'POTE', 'CD', 'UNID', 'R', 'UNI', 'kit', 'DI','KT', 'Kit', 'pe', 'und','gramas','1']

    # Primeiro, busca pela chave de acesso e número de série/número de NF
    for i, linha in enumerate(linhas):
        if "0,00CHAVE DE ACESSO DA NF-e P/ CONSULTA DE AUTENTICIDADE NO SITE WWW.NFE.FAZENDA.GOV.BR" in linha:
            if i + 1 < len(linhas):
                chave_acesso = linhas[i + 1]
                print(f"Chave de Acesso: {chave_acesso}")
        if 'SÉRIE:' in linha:
            if i + 4 < len(linhas):
                linhaNumSerieEnumNF = linhas[i + 4]
                numeroSerie = linhaNumSerieEnumNF.split("-")[1].strip()
                numeroNF = linhaNumSerieEnumNF.split("-")[0].strip()
                print(f"Número Série: {numeroSerie}, Número NF: {numeroNF}")

    # Depois, busca pela quantidade
    for linha in linhas:
        quantidades_encontradas = encontrar_quantidade(linha, termos_chave)
        if quantidades_encontradas:
            quantidade.extend(quantidades_encontradas)
            print(f"Quantidade(s): {quantidade}")

    return chave_acesso, numeroNF, numeroSerie, quantidade
def ler_pdfs_na_pasta(pasta):
    i = 0
    y = 0 

    chave_acesso = numeroNF = numeroSerie = quantidade = None  # Inicializando as variáveis
    NFmultiplas = None
    numeroControle =  None
    dataVenc = None
    TotalRecolher = None  
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.pdf'):
            caminho_completo = os.path.join(pasta, arquivo)
            if arquivo.startswith('DANFE'):
                i += 1
                chave_acesso, numeroNF, numeroSerie, quantidade = processar_danfe(caminho_completo)
                print(chave_acesso, numeroControle, numeroSerie, quantidade)
                input('')
                # os.remove(caminho_completo)
                print(f"Arquivo DANFE excluído: {caminho_completo}")
    print(f'DANFE: {i} GUIA: {y}')
    return chave_acesso, numeroNF, numeroSerie, quantidade, NFmultiplas, numeroControle, dataVenc, TotalRecolher

diretorio_base = os.path.dirname(os.path.abspath(__file__))
caminhoData = os.path.join(diretorio_base, 'data')

#ler_pdfs_na_pasta(caminhoData)
