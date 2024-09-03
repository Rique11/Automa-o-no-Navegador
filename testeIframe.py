class DadosTabela:
    def __init__(self, numNF, numLote, numCte, dataEmi, valoraPagar, numDAR, situacao):
        self.numNF = numNF
        self.numLote = numLote
        self.numCte = numCte
        self.dataEmi = dataEmi
        self.valoraPagar = valoraPagar
        self.numDAR = numDAR
        self.situacao = situacao

    def __str__(self):
        return f"{self.numNF},{self.numLote},{self.numCte},{self.dataEmi},{self.valoraPagar},{self.numDAR},{self.situacao}"
    
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException, NoSuchWindowException,WebDriverException
import urllib3
from urllib3.exceptions import MaxRetryError
import time
import traceback
import os
import shutil
from bs4 import BeautifulSoup
from puxaDados import ler_pdfs_na_pasta
from CriaTabela import criaTabela, enviaEmailComTabela
import PyPDF2
from selenium.common.exceptions import NoAlertPresentException


def inicializar_ambiente(diretorio_base, pasta_data, pasta_guias):
    # Caminhos das pastas e arquivo
    arquivo_tabela_formatada = os.path.join(diretorio_base, 'resultado_formatado.xlsx')

    # Função para garantir que uma pasta esteja vazia
    def limpar_pasta(pasta):
        if os.path.exists(pasta):
            for filename in os.listdir(pasta):
                file_path = os.path.join(pasta, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)  # Remove arquivos
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)  # Remove subdiretórios
                except Exception as e:
                    print(f'Erro ao excluir {file_path}: {e}')
        else:
            os.makedirs(pasta)  # Cria a pasta se não existir

    # Limpar as pastas 'data' e 'Guias'
    limpar_pasta(pasta_data)
    limpar_pasta(pasta_guias)

    # Verifica e remove o arquivo 'tabela_formatada.xlsx'
    if os.path.exists(arquivo_tabela_formatada):
        try:
            os.remove(arquivo_tabela_formatada)
            #print(f"Arquivo {arquivo_tabela_formatada} removido.")
        except Exception as e:
            print(f"Erro ao remover {arquivo_tabela_formatada}: {e}")
    else:
        print(f"Arquivo {arquivo_tabela_formatada} não encontrado.")

    print("Ambiente inicializado com sucesso.")

def close_alerts(navegador):
    while True:
        try:
            # Verifica se há algum alerta presente
            alert = navegador.switch_to.alert
            alert.accept()  # Fecha o alerta
            print("Alerta fechado.")
            time.sleep(1)  # Pequeno atraso para garantir que o alerta foi fechado
        except NoAlertPresentException:
            break  # Sai do loop se não houver mais alertas

def ler_progresso(caminho_arquivo_progresso):
    if os.path.exists(caminho_arquivo_progresso):
        with open(caminho_arquivo_progresso, 'r') as f:
            return int(f.read().strip())
    return 0

def salvar_progresso(caminho_arquivo_progresso, indice):
    with open(caminho_arquivo_progresso, 'w') as f:
        f.write(str(indice))


def run_backend_process(chave):
    # Obtém o diretório base onde o script está localizado
    diretorio_base = os.path.dirname(os.path.abspath(__file__))
    
    caminho_arquivo_progresso = os.path.join(diretorio_base, 'progresso.txt')
    indice_inicio = ler_progresso(caminho_arquivo_progresso)
    if indice_inicio == -1:
        indice_inicio = 0
    # Constrói os caminhos relativos a partir do diretório base
    caminhoData = os.path.join(diretorio_base, 'data')
    caminhoGuias = os.path.join(diretorio_base, 'Guias')
    if indice_inicio == 0:
        inicializar_ambiente(diretorio_base, caminhoData, caminhoGuias)
    else : 
        print("Reinciando com progresso salvo")
        
    # Limpar cache do webdriver_manager
    cache_path = os.path.expanduser("~/.wdm")
    if os.path.exists(cache_path):
        shutil.rmtree(cache_path)

    # Configurações de opções do Chrome
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--ignore-certificate-errors")  # Ignorar erros de certificado
    chrome_options.add_argument("--ignore-ssl-errors")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # chrome_options.add_argument("--headless")  # Descomente para rodar sem interface gráfica
    chrome_options.add_experimental_option(
            'prefs', {
            # Muda o diretório padrão de download
            "download.default_directory":  caminhoData,

            # Não perguntar ao fazer o download
            "download.prompt_for_download": False,

            # Abrir PDFs externamente, sem usar o visualizador embutido
            "plugins.always_open_pdf_externally": True
        }
    )
    try:
        def instanciaChromeDriver():
            driver_path = ChromeDriverManager().install()
            service = Service(driver_path)
            return service

        def abreEnderecoNoNavegador(navegador):
            navegador.get("https://www4.sefaz.pb.gov.br/atf/seg/SEGf_AcessarFuncao.jsp?cdFuncao=FIS_1302&amp;idSERVirtual=S&amp;h=https://www.sefaz.pb.gov.br/ser/servirtual/credenciamento/info")
            print("Navegador iniciado com sucesso!")

        def desceFinalPaginaCarregada():
            navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)


        # Instanciar o ChromeDriverManager
        service = instanciaChromeDriver()
        # Inicializar o navegador com as opções configuradas
        navegador = webdriver.Chrome(service=service, options=chrome_options)
        abreEnderecoNoNavegador(navegador)
        
        # Listar todos os iframes presentes na página
        iframes = navegador.find_elements(By.TAG_NAME, 'iframe')
        #print(f"Total de iframes encontrados: {len(iframes)}")

        # Acessar o iframe pelo índice
        if iframes:
            navegador.switch_to.frame(iframes[1])
            #print("Mudou para o iframe")

            try:
                # Localizar a barra de pesquisa e inserir uma informação
                barra_pesquisa = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.NAME, "edtChaveAcesso"))
                )
                barra_pesquisa.send_keys(chave)  # Substitua pela informação desejada
                barra_pesquisa.send_keys(Keys.RETURN)  # Pressionar Enter
                print("Informação inserida na barra de pesquisa")
                time.sleep(2)
                linhaMDFe = navegador.find_elements(By.CLASS_NAME, "tdPadrao")
                localNumeroMDFe = linhaMDFe[1].text.strip(" ")
                numeroMDFe = localNumeroMDFe.split(" ")[0]
                print(f"NUMERO MDF-e: {numeroMDFe}")

                # Aperta botaozinho que abre com a chave
                botaoAbreTabela = WebDriverWait(navegador, 5).until(
                    EC.presence_of_element_located((By.NAME, "rdbChavePrimaria"))
                )
                botaoAbreTabela.click()
                
                # Aperta no detalhes para abrir a tabela
                botaoDetalhar = WebDriverWait(navegador, 2).until(
                    EC.presence_of_element_located((By.NAME, "btnDetalhar"))
                )
                botaoDetalhar.click()

                tabela_desejada = None

                # Continuar rolando até encontrar a tabela desejada
                while tabela_desejada is None:
                    # Rola até o final da página
                    navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)

                    # Espera carregar as tabelas e extrai a desejada
                    tabelas = WebDriverWait(navegador, 10).until(
                        EC.presence_of_all_elements_located((By.CLASS_NAME, "fontePadrao"))
                    )
                   

                    # Verifica se a tabela desejada está entre as carregadas
                    for tabela in tabelas:
                        if "Lista NF-e (Destinada a Não Contribuinte)" in tabela.get_attribute("outerHTML"):
                            tabela_desejada = tabela
                            break

                    # Se a tabela não foi encontrada, continua rolando
                    if tabela_desejada is None:
                        print("Carregando...")
                    else:
                        print("Tabela desejada encontrada.") 

                infosTabela = []  
                linhasParaUsarAlternada = navegador.find_elements(By.CLASS_NAME, "tdAlternada")
                linhasParaUsarPadrao = navegador.find_elements(By.CLASS_NAME, "tdPadrao")
                linhasParaUsar = linhasParaUsarAlternada + linhasParaUsarPadrao
                print(indice_inicio)
                for i in range(indice_inicio, len(linhasParaUsar)):
                    linha = linhasParaUsar[i]
                    if 'EM ABERTO' in linha.text:
                        try:
                            navegador.execute_script("arguments[0].scrollIntoView(true);", linha)
                            time.sleep(1)  # Adicione um pequeno delay entre os cliques
                            checkBoxGuia = linha.find_element(By.NAME, 'chkNrChaveNFe')
                            checkBoxGuia.click()
                            close_alerts(navegador)

                                        # Extrai dados das colunas da linha
                            colunas = linha.find_elements(By.TAG_NAME, 'td')
                            dados = [coluna.text.strip() for coluna in colunas]
                            
                            if len(dados) > 12: 
                                numNF = dados[1]
                                numLote = dados[2]
                                numCte = dados[4]
                                dataEmi = dados[8]
                                valoraPagar = dados[12]
                                numDAR = dados[13]
                                situacao = dados[15]
                                
                                # Exemplo de armazenamento de dados (ou manipulação conforme necessário)
                                infosTabela.append(DadosTabela(numNF, numLote, numCte, dataEmi, valoraPagar, numDAR, situacao))
                                
                                # Imprime os dados ou manipula conforme necessário
                                #print(dados)
                            #Abre Danfe
                            NumNFe =  linha.find_element(By.CSS_SELECTOR, "a[href^='javascript:gerarDanfe']")
                            NumNFe.click()

                            desceFinalPaginaCarregada()

                            botaoEmiteDar = navegador.find_element(By.NAME, 'btnEmitirDAR_NC')
                            botaoEmiteDar.click()

                            # Mudar para a nova guia aberta
                            navegador.switch_to.window(navegador.window_handles[-1])
                            temGuia = True
                            close_alerts(navegador)
                            #Fechar a guia atual se o link for o desejado
                            current_url = navegador.current_url
                            # Mudar para a nova guia aberta
                            if len(navegador.window_handles) > 1:
                                navegador.switch_to.window(navegador.window_handles[-1])
                                try:
                                    close_alerts(navegador)
                                    current_url = navegador.current_url
                                    if "https://www4.sefaz.pb.gov.br/atf/seg/SEGf_EmitirMensagemTelaCheia.jsp?codigo=3269" in current_url:
                                        time.sleep(2)
                                        navegador.close()
                                except NoSuchWindowException:
                                    print("A guia foi fechada antes que pudéssemos interagir com ela.")
                            else:
                                print("Nenhuma guia nova foi aberta.")
                            
                                # Voltar para a guia original se a URL não corresponder
                            navegador.switch_to.window(navegador.window_handles[0])

                                # Mudar para o iframe novamente
                            navegador.switch_to.frame(iframes[1])
                        
                            navegador.execute_script("arguments[0].scrollIntoView(true);", linha)
                            checkBoxGuia.click()
                            time.sleep(3)
                            notaFiscal = numNF.zfill(9)
                            numMDFe = numeroMDFe 
                            chave_acesso, numeroNF, numeroSerie,valorNota, quantidade, NFMultiplas, numeroControle, dataVenc,totalRecolhe = ler_pdfs_na_pasta(caminhoData,caminhoGuias, notaFiscal)
                            print(f"DANFE: {chave_acesso}, {numeroNF}, {numeroSerie},{valorNota}, {quantidade}")
                            print(f"GUIA: {NFMultiplas}, {numeroControle}, {dataVenc},{totalRecolhe}")
                            
                            if numeroControle == None:
                                temGuia = False
                            else: 
                                temGuia = True
                
                            if temGuia == True:
                                print(numMDFe)
                                criaTabela(numeroControle, notaFiscal, chave_acesso,valorNota , numCte, totalRecolhe, NFMultiplas, numMDFe, dataVenc, quantidade)
                            salvar_progresso(caminho_arquivo_progresso, i)
                        except ElementClickInterceptedException:
                            navegador.execute_script("arguments[0].click();", checkBoxGuia)
                        except (MaxRetryError, WebDriverException, urllib3.exceptions.NewConnectionError, ConnectionRefusedError) as e:
                            print(f"Erro de conexão ou problema no WebDriver: {e}")
                            # Tente reiniciar o navegador, recarregar a página, ou reestabelecer a conexão
                            navegador.quit()
                enviaEmailComTabela()
            except TimeoutException:
                print("Timeout ao tentar localizar a barra de pesquisa")
                navegador.quit()
            except NoSuchElementException:
                print("Barra de pesquisa não encontrada")
                navegador.quit()
            finally:
                navegador.switch_to.default_content()
        else:
            print("Nenhum iframe encontrado")
        
    finally:
        # Mantenha o navegador aberto por um tempo para ver o resultado
        time.sleep(10)
        navegador.quit()

#run_backend_process("") #descomentar para teste sem front
