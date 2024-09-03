
import pandas as pd
import os
from decouple import config
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import zipfile
import re


def enviaEmailComTabela():
    # Carrega as credenciais do arquivo .env
    seu_email = config('EMAIL')
    sua_senha = config('SENHA')
    

    # Configuração do servidor SMTP
    servidor_email = smtplib.SMTP('smtp.gmail.com', 587)
    servidor_email.starttls()  # Inicia a conexão TLS

    # Login no servidor SMTP
    servidor_email.login(seu_email, sua_senha)

    # Criação do conteúdo do e-mail
    destinatario = 'suporte.difal@oprlogistica.com.br'
    assunto = 'Assunto do Email'
    corpo_email = 'Teste de envio de email por codigo em python'

    mensagem = MIMEMultipart()
    mensagem['From'] = seu_email
    mensagem['To'] = destinatario
    mensagem['Subject'] = assunto 
    mensagem.attach(MIMEText(corpo_email, 'plain'))

    diretorio_base = os.path.dirname(os.path.abspath(__file__))

    # Constrói o caminho relativo para o arquivo Excel
    caminho_do_arquivo = os.path.join(diretorio_base, 'resultado_formatado.xlsx')
    nome_do_arquivo = os.path.basename(caminho_do_arquivo)
    #Constroi o caminho para o progesso.txt
    caminhoArquivoProgresso = os.path.join(diretorio_base, 'progresso.txt')
    # Verifique se o arquivo existe
    if not os.path.isfile(caminho_do_arquivo):
        print(f"Erro: O arquivo {caminho_do_arquivo} não existe.")
    # Adiciona o anexo
    else:
        try:
            # Abrir o arquivo em modo leitura binária
            with open(caminho_do_arquivo, 'rb') as arquivo:
                parte = MIMEBase('application', 'octet-stream')
                parte.set_payload(arquivo.read())

            # Codifica o payload em Base64
            encoders.encode_base64(parte)
            parte.add_header('Content-Disposition', f'attachment; filename={nome_do_arquivo}')
            mensagem.attach(parte)

            # Anexar todos os PDFs da pasta 'guia'
            pasta_guia = os.path.join(diretorio_base, 'Guias')
            pdfs = [f for f in os.listdir(pasta_guia) if f.endswith('.pdf')]
            tamanho_total = 0

            for pdf in pdfs:
                caminho_pdf = os.path.join(pasta_guia, pdf)
                tamanho_total += os.path.getsize(caminho_pdf)

            #compactá-los em um .zip
            
            caminho_zip = os.path.join(diretorio_base, 'arquivos.zip')
            with zipfile.ZipFile(caminho_zip, 'w') as arquivo_zip:
                for pdf in pdfs:
                    caminho_pdf = os.path.join(pasta_guia, pdf)
                    arquivo_zip.write(caminho_pdf, os.path.basename(caminho_pdf))

            # Anexar o arquivo .zip
            with open(caminho_zip, 'rb') as arquivo:
                parte_zip = MIMEBase('application', 'octet-stream')
                parte_zip.set_payload(arquivo.read())
            encoders.encode_base64(parte_zip)
            parte_zip.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_zip)}')
            mensagem.attach(parte_zip)

            # Deletar o arquivo .zip após o envio
            os.remove(caminho_zip)

            # Envio do e-mail
            servidor_email.sendmail(seu_email, destinatario, mensagem.as_string())
            print("E-mail enviado com sucesso!")
            
            with open(caminhoArquivoProgresso, 'w') as arqProgresso:
                arqProgresso.write('0')
            os.remove(caminho_do_arquivo)
            # Deletar os PDFs após o envio
            for pdf in pdfs:
                os.remove(os.path.join(pasta_guia, pdf))

        except Exception as e:
            print(f"Erro ao enviar o e-mail: {e}")
        finally:
            servidor_email.quit()  # Encerra a conexão com o servidor SMTP


def criaTabela(numeroControle, NFPrincipal, chave_acesso,valorNota, numCTE, TotalRecolhe, NFMultiplas, numMDFe, dataVenc, quantidade):
    # Obtém o diretório base onde o script está localizado
    diretorio_base = os.path.dirname(os.path.abspath(__file__))

    # Constrói o caminho relativo para o arquivo Excel
    arquivo_excel = os.path.join(diretorio_base, 'resultado_formatado.xlsx')
    
    # Formata o número da NFPrincipal com 9 dígitos
    NFPrincipalFormat = str(NFPrincipal).zfill(9)
    
    # Verifique se o arquivo já existe
    if os.path.exists(arquivo_excel):
        # Se o arquivo existe, carregue-o
        df_existente = pd.read_excel(arquivo_excel, dtype={'NF Principal': str})
    else:
        # Se o arquivo não existe, crie um DataFrame vazio com as colunas necessárias
        df_existente = pd.DataFrame(columns=[
            "N° de Controle", 
            "NF Principal", 
            "Chave de Acesso NF Principal",
            "Chave de Acesso NF Principal (S/CAR)",
            "Valor Mercadoria NF",
            "ETIQUETA",
            "Tarifa",
            "CTE OPR", 
            "Valor a Pagar", 
            "Notas Adicionais", 
            "N° do MDFe", 
            "Renomeação", 
            "Data Vcto", 
            "QT VOLUME"
        ])
    
    renomeacao = f"{numeroControle}{NFPrincipalFormat}"
    chave_acesso_formatada = re.sub(r'\D', '', chave_acesso)
    quantidadeTotal = sum(float(q) for q in quantidade)
    valorNotaFormatada = f"R${valorNota}"
    # Criar um dicionário com os dados da nova linha
    nova_linha = {
        "N° de Controle": numeroControle,
        "NF Principal": NFPrincipalFormat,
        "Chave de Acesso NF Principal": chave_acesso,
        "Chave de Acesso NF Principal (S/CAR)":chave_acesso_formatada,
        "Valor Mercadoria NF": valorNotaFormatada,
        "CTE OPR": numCTE,
        "Valor a Pagar": TotalRecolhe,
        "Notas Adicionais": NFMultiplas,
        "N° do MDFe": numMDFe,
        "Renomeação": renomeacao,
        "Data Vcto": dataVenc,
        "QT VOLUME": quantidadeTotal
    }

    # Adicionar a nova linha ao DataFrame existente
    df_existente = df_existente._append(nova_linha, ignore_index=True)
    
    # Salvar o DataFrame atualizado no arquivo Excel
    with pd.ExcelWriter(arquivo_excel, engine='xlsxwriter') as writer:
        df_existente.to_excel(writer, index=False, sheet_name='Dados')
        
        # Acessar o workbook e worksheet
        workbook = writer.book
        worksheet = writer.sheets['Dados']

        # Aplicar formatação no cabeçalho
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        # Ajustar a largura das colunas e aplicar formatação no cabeçalho
        for col_num, value in enumerate(df_existente.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, len(value) + 5)  # Ajuste para o tamanho da célula

        # Forçar a coluna "NF Principal" a ser texto para preservar zeros à esquerda
        worksheet.set_column(df_existente.columns.get_loc("NF Principal"), df_existente.columns.get_loc("NF Principal"), 12, workbook.add_format({'num_format': '@'}))
        
        # Ajustar a largura das colunas com base no conteúdo
        for i, col in enumerate(df_existente.columns):
            max_len = df_existente[col].astype(str).map(len).max()  # Encontrar o tamanho máximo do conteúdo
            worksheet.set_column(i, i, max_len + 5)  # Ajustar a largura da célula com base no conteúdo

    print("Linha adicionada no arquivo Excel formatado com sucesso!")
#enviaEmailComTabela()
