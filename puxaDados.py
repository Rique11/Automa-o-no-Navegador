import os
import re 
import PyPDF2
import pandas as pd

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
    return quantidades if quantidades else None

def processar_danfe(pdf_path):
    print(f"\nLendo DANFE: {pdf_path}\n")
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()

    linhas = text.splitlines()
    
    chave_acesso = None
    numeroNF = None
    numeroSerie = None
    valorNota = None
    quantidade = []

    # Lista de termos chave que podem indicar a quantidade
    termos_chave = ['PACOTE', 'LA', 'UN', 'Un', 'un', 'peca', 'PC', 'CX', 'Pc', 'pc', 'MT', 'JG', 'KIT', 'UND', 'POTE', 'CD', 'UNID', 'R', 'UNI', 'kit', 'DI','KT', 'Kit', 'pe', 'und','gramas','DZ','DS','unid','pç','KG','PAR','PR','PCT.','pct','RL','DP','UN1','Par','kg','Pares','Kg','g','Unidad','p?.','PT','REFIL','PÇ','PA','PE','PCT']

    # Primeiro, busca pela chave de acesso e número de série/número de NF
    for i, linha in enumerate(linhas):
        if "0,00CHAVE DE ACESSO DA NF-e P/ CONSULTA DE AUTENTICIDADE NO SITE WWW.NFE.FAZENDA.GOV.BR" in linha:
            if i + 1 < len(linhas):
                chave_acesso = linhas[i + 1]
                print(f"\nChave de Acesso: {chave_acesso}")
        if 'SÉRIE:' in linha:
            if i + 4 < len(linhas):
                linhaNumSerieEnumNF = linhas[i + 4]
                numeroSerie = linhaNumSerieEnumNF.split("-")[1].strip()
                numeroNF = linhaNumSerieEnumNF.split("-")[0].strip()
                print(f"\nNúmero Série: {numeroSerie}, Número NF: {numeroNF}")
                linhaValorNota = linhas[i + 5]
                valorNota = linhaValorNota.split(" ")[3].strip()
                print(f"valorNota: {valorNota}")
    # Depois, busca pela quantidade
    for linha in linhas:
        quantidades_encontradas = encontrar_quantidade(linha, termos_chave)
        if quantidades_encontradas:
            quantidade.extend(quantidades_encontradas)
            print(f"\nQuantidade(s): {quantidade}")

    return chave_acesso, numeroNF, numeroSerie,valorNota, quantidade



def verifica_e_processa_numero_controle(numeroControle, caminho_do_arquivo, novo_nome):
    # Carrega o arquivo Excel, se existir
    if os.path.exists(caminho_do_arquivo):
        df = pd.read_excel(caminho_do_arquivo, dtype={'N° de Controle': str})

        # Remove espaços em branco e garante que ambos sejam strings
        numeroControle_formatado = numeroControle.strip()

        # Verifica se o número de controle já existe na planilha
        if numeroControle_formatado in df['N° de Controle'].str.strip().values:
            print(f"Número de controle {numeroControle_formatado} já existe na planilha. Excluindo arquivo.")
            os.remove(novo_nome)
            return None

    return numeroControle

def processar_saida(pdf_path, destino, numeroNF):
    print(f"Lendo Saida: {pdf_path}")
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    #print(f"Texto extraído de Saida: {text}")

    linhas = text.splitlines()

    NFmultiplas = None
    numeroControle =  None
    dataVenc = None
    TotalRecolher = None  

    for i, linha in enumerate(linhas):
        if "DOCs:" in linha:
            if i + 1 < len(linhas):
                linha_seguinte = linhas[i + 1]
                NFmultiplas = linha_seguinte
        if 'ICMS DIFAL NÃO CONTRIBUINTE-ENTRADA' in linha:
            if i + 1 < len(linhas):
                linha_seguinte = linhas[i + 1]
                numeroControle = linha_seguinte
        if 'FUNCEP - FATURA/ENTRADA' in linha: 
            if i+1 < len(linhas):
                linha_seguinte = linhas[i+1]
                numeroControle = linha_seguinte
        if '20 - Nome da Firma ou Razão Social' in linha:
            if i - 1 >= 0:
                dataVenc = linhas[i - 1].split(" ")[0]
        if '27 - Autenticação Mecânica' in linha:
            if i - 1 >= 0:
                TotalRecolher = linhas[i - 1].split(" ")[0]

    if NFmultiplas:
        print(f"NFmultiplas: {NFmultiplas}")
    else:
        print("NFmultiplas não encontrada.")

    if numeroControle:
        print(f'numeroControle: {numeroControle}')
    else:
        print("numeroControle não encontrado.")

    if dataVenc:
        print(f"dataVenc: {dataVenc}")
    else:
        print("dataVenc não encontrada.")

    if TotalRecolher:
        print(f"TotalRecolher: {TotalRecolher}")
    else:
        print("TotalRecolher não encontrado.")

    if numeroNF is not None:
        numeroNF = numeroNF.zfill(9)

    # Renomeia o arquivo com um número sequencial e move para a pasta de destino
    novo_nome = os.path.join(destino, f"{numeroControle.strip()}{numeroNF}.pdf")

    # Verifica se o numeroControle já existe na planilha
    numeroControle = verifica_e_processa_numero_controle(numeroControle, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resultado_formatado.xlsx'), pdf_path)
    #print(f"NUMERO DE CONTROLE: {numeroControle}")
    # Se o número de controle for None, significa que ele já existe na planilha
        # Verifica se o arquivo de destino já existe
    if os.path.exists(novo_nome):
        return NFmultiplas, numeroControle, dataVenc, TotalRecolher  # Apaga o arquivo de destino se já existir
    if numeroControle is None:
        return None, None, None, None

    os.rename(pdf_path, novo_nome)
    # Se chegou aqui, o número de controle não é duplicado, então podemos renomear e mover o arquivo
    print(f"Arquivo Saida renomeado e movido para: {novo_nome}")

    return NFmultiplas, numeroControle, dataVenc, TotalRecolher
    

def ler_pdfs_na_pasta(pasta, destino, notaNF):
    i = 0
    y = 0 

    #chave_acesso, numeroNF, numeroSerie, quantidade = None
    contador_saida = 0  # Inicializa o contador para Saida
    chave_acesso = numeroNF = numeroSerie = valorNota = quantidade = None  # Inicializando as variáveis
    NFmultiplas = None
    numeroControle =  None
    dataVenc = None
    TotalRecolher = None  
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.pdf'):
            caminho_completo = os.path.join(pasta, arquivo)
            if arquivo.startswith('DANFE'):
                i += 1
                chave_acesso, numeroNF, numeroSerie,valorNota, quantidade = processar_danfe(caminho_completo)
                os.remove(caminho_completo)
                print(f"Arquivo DANFE excluído: {caminho_completo}")
            elif arquivo.startswith('Saida'):
                y += 1
                NFmultiplas, numeroControle, dataVenc,TotalRecolher = processar_saida(caminho_completo, destino, notaNF)
                contador_saida += 1  # Incrementa o contador após processar cada Saida

    return chave_acesso, numeroNF, numeroSerie,valorNota, quantidade, NFmultiplas, numeroControle, dataVenc, TotalRecolher
#ler_pdfs_na_pasta(r"C:\Users\João Vitor\BotGuias\BotNotas\data",r"C:\Users\João Vitor\BotGuias\BotNotas\Guias","DANFE_1725041995860")



