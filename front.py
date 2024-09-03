import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
from threading import Thread
from PyPDF2 import PdfReader
from testeIframe import run_backend_process
import time 
class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Robo Alda maria")
        self.root.geometry("500x220")
        self.root.configure(bg="black")

        self.button_style = {"bg": "orange", "fg": "black", "font": ("Helvetica", 12)}

        # Variável para armazenar as chaves válidas
        self.chaves_validas = []

        # Caminho do arquivo de progresso
        self.caminhoArquivoProgresso = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'progresso.txt')

        # Campo de entrada para digitar a chave
        self.pdf_entry = tk.Entry(root, width=50)
        self.pdf_entry.pack(pady=20)

        # Botão para inserir/processar a chave
        self.upload_pdf_button = tk.Button(root, text="Inserir Chave", command=self.process_key, **self.button_style)
        self.upload_pdf_button.pack(pady=20)

        self.button_frame = tk.Frame(root, bg="black")
        self.button_frame.pack(pady=20)

        self.start_button = tk.Button(self.button_frame, text="Start", command=self.start_operation, **self.button_style)
        self.start_button.grid(row=0, column=1, padx=10)

        self.delete_button = tk.Button(self.button_frame, text="Deletar Progresso", command=lambda: self.apagarProgressoSalvo(self.caminhoArquivoProgresso), **self.button_style)
        self.delete_button.grid(row=0, column=0, padx=10)

        self.operation_label = tk.Label(root, text="Operation in progress...", fg="orange", bg="black")

    def process_key(self):
        # Pega a chave digitada no campo de entrada
        key = self.pdf_entry.get()

        if key:  # Verifica se a entrada não está vazia
            # Adiciona a chave à lista
            self.chaves_validas.append(key)
            # Limpa o campo de entrada
            self.pdf_entry.delete(0, tk.END)
            # Exibe a lista atualizada no console (ou use outro método de exibição)
            print(f"Chave inserida: {key}")
            print(f"Lista atualizada: {self.chaves_validas}")

    
    def apagarProgressoSalvo(self, caminhoArquivoProgresso):
        if os.path.exists(caminhoArquivoProgresso):
            with open(caminhoArquivoProgresso, "w") as arqProgresso: 
                arqProgresso.write("0")
            print(f"Arquivo de progresso {caminhoArquivoProgresso} zerado com sucesso.")
        else:
            print(f"O arquivo {caminhoArquivoProgresso} não foi encontrado.")

    def start_operation(self):
        def confereProgressoSalvo():
            diretorio_base = os.path.dirname(os.path.abspath(__file__))
            caminhoArquivoProgresso = os.path.join(diretorio_base, 'progresso.txt')
            with open(caminhoArquivoProgresso, "r") as arqProgresso:
                linha = arqProgresso.readline().strip()  # Lê e remove espaços em branco
                if linha:
                    try:
                        progresso = int(linha)
                    except ValueError:
                        progresso = -1  # Define um valor padrão se a conversão falhar
                else:
                    progresso = -1  # Define um valor padrão se a linha estiver vazia
            
            return progresso
         
        def operation():
            diretorio_base = os.path.dirname(os.path.abspath(__file__))
            caminhoArquivoProgresso = os.path.join(diretorio_base, 'progresso.txt')

            with open(caminhoArquivoProgresso, "+r") as arqProgresso: 
                linha = arqProgresso.readline().strip()
                if linha:
                    try:
                        progresso = int(linha)
                    except ValueError:
                        progresso = -1
                else:
                    progresso = -1
                if progresso == 0: 
                    arqProgresso.seek(0)  # Volta para o início do arquivo
                    arqProgresso.truncate()  # Apaga todo o conteúdo do arquivo 
                    arqProgresso.write('-1')
                    arqProgresso.flush()  # Certifica que o conteúdo é gravado no disco
                    
                for chave in self.chaves_validas:
                    numeroProgresso = confereProgressoSalvo()
                    while numeroProgresso != 0:  # Verifica continuamente se o progresso.txt ainda existe
                        
                        try:
                            if numeroProgresso == 0:
                                break  # Sai do loop se o arquivo de progresso não existir
                            run_backend_process(chave)  # Função que realiza o processo de backend
                            numeroProgresso = confereProgressoSalvo()
                                                # Adicionar uma verificação após a execução do processo
                        except Exception as e:
                            print(f"Erro durante a operação: {e}")
                            time.sleep(2)  # Adiciona um pequeno delay antes de tentar novamente (opcional)
                        finally:
                            if not confereProgressoSalvo():
                                break  # Sai do loop se o arquivo de progresso não existir

        self.operation_label.pack()
        Thread(target=operation).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()