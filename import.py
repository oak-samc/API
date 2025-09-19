# -*- coding: utf-8 -*-
#
# Este script automatiza o envio de e-mails via Outlook,
# anexando arquivos PDF com base no nome da cidade.
#
# === CONFIGURAÇÕES ===
import os
import time
import re
import unicodedata
import win32com.client as win32
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, filedialog
import threading
 
# E-mails para cópia (CC)
copia_emails = ["alinemilacki@gmail.com", "yasmin.oliveira@cebraspe.org.br", "enade2025@cebraspe.org.br"]
 
# === TEXTO DO E-MAIL (HTML) ===
corpo_email_html = """
<p>Prezado (a) Coordenador (a) Estadual e Coordenador (a) de Local, bom dia!</p>
<p>Em virtude da aplicação das provas objetivas e de redação do Exame Nacional de Desempenho dos Estudantes (ENADE 2025), que ocorrerá no dia 23 de novembro de 2025, no período vespertino, encaminhamos anexo a distribuição de salas do seu município, referente à etapa de ensalamento e confirmação dos dados referentes ao espaço físico.</p>
<p><strong>Procedimento de Validação:</strong><br>
O Coordenador deverá visualizar sua distribuição e verificar se as informações na distribuição de salas estão corretas.
<ol>
<li><strong>Validar:</strong> Confirmar se as informações estão corretas, tais como</li>
<li><strong>Recusar:</strong> Caso as informações não estejam corretas, recusar e informar o motivo e os ajustes necessários.</li>
</ol></p>
<table style="width: 100%; border-collapse: collapse;">
<thead>
<tr style="background-color: #007bff; color: white;">
<th style="border: 1px solid #ddd; padding: 8px;">Dados para confirmação</th>
<th style="border: 1px solid #ddd; padding: 8px;">Certo</th>
<th style="border: 1px solid #ddd; padding: 8px;">Errado</th>
<th style="border: 1px solid #ddd; padding: 8px;">Ajuste</th>
</tr>
</thead>
<tbody>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Nome completo da instituição (nome exposto na fachada)?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Endereço completo da instituição (inclusive a cidade)?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Número de salas utilizadas e os respectivos andares?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Capacidade de candidatos distribuídos em cada sala?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Os Blocos foram agrupados de maneira correta?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">O quantitativo de sala por bloco está de acordo com a informação repassada por você?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">A escola com Atendimento Especializado tem a acessibilidade necessária?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
</tbody>
</table>
<p>Ressaltamos que os participantes foram ensalados conforme o cadastro das instituições do seu município no SinCef. Solicitamos que você proceda à conferência dos dados.</p>
<p>Para garantir a qualidade e a excelência do nosso trabalho e cumprir os prazos estabelecidos, solicitamos a resposta a esse e-mail até o dia 20 de setembro de 2025, às 09:00h (horário de Brasília).</p>
<p>Em caso de dúvidas, entre em contato com o Cebraspe pelo e-mail <a href="mailto:enade2025@cebraspe.org.br">enade2025@cebraspe.org.br</a> ou telefone (61) 2109-5810.</p>
"""
 
# === FUNÇÕES ===
 
def normalizar_texto(texto):
    """Remove acentos e deixa tudo maiúsculo."""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto.strip().upper()
 
def enviar_email_com_anexo(destinatario, assunto, corpo_html, anexo_path, cc_list=None, log_text=None):
    """Envia um e-mail via Outlook com anexo."""
    try:
        if not destinatario or "@" not in destinatario:
            messagebox.showerror("Erro", f"E-mail inválido: {destinatario}")
            if log_text: log_text.insert(tk.END, f"❌ Erro: E-mail do destinatário inválido.\n")
            return
       
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
 
        mail.To = destinatario
        mail.Subject = assunto
        mail.HTMLBody = corpo_email_html
       
        if cc_list:
            mail.CC = "; ".join(cc_list)
       
        if anexo_path and os.path.exists(anexo_path):
            mail.Attachments.Add(str(anexo_path))
            if log_text: log_text.insert(tk.END, f"📎 Anexo adicionado: {anexo_path}\n")
        else:
            messagebox.showerror("Erro", f"Anexo não encontrado ou caminho inválido: {anexo_path}.")
            if log_text: log_text.insert(tk.END, f"❌ Erro: Anexo não encontrado ou caminho inválido.\n")
            return
       
        mail.Send()
        messagebox.showinfo("Sucesso", f"E-mail enviado com sucesso para: {destinatario}")
        if log_text: log_text.insert(tk.END, f"✅ E-mail enviado para: {destinatario}\n")
       
    except Exception as e:
        messagebox.showerror("Erro ao Enviar E-mail", f"ERRO: {e}\n\nVerifique se o Outlook está aberto e é o cliente de e-mail padrão do Windows.")
        if log_text: log_text.insert(tk.END, f"❌ ERRO AO ENVIAR O E-MAIL: {e}\n")
 
 
# === INTERFACE GRÁFICA ===
class EmailApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automatizador de E-mails ENADE - v2.0")
        self.geometry("600x600")
       
        self.caminho_pasta_pdfs = ""
        self.pdf_map = {}
       
        self.create_widgets()
 
    def create_widgets(self):
        main_frame = tk.Frame(self)
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)
 
        # Selecionar Pasta
        tk.Label(main_frame, text="Selecione a pasta com os PDFs:").pack(pady=(0, 5))
        select_frame = tk.Frame(main_frame)
        select_frame.pack(fill="x", pady=(0, 10))
        self.folder_path_label = tk.Label(select_frame, text="Nenhuma pasta selecionada", anchor="w")
        self.folder_path_label.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(select_frame, text="Selecionar Pasta", command=self.select_folder).pack(side="right")
       
        # E-mail do Destinatário
        tk.Label(main_frame, text="E-mail do Destinatário:").pack(pady=(0, 5))
        self.destinatario_entry = tk.Entry(main_frame)
        self.destinatario_entry.pack(fill="x", ipady=4)
       
        # E-mail de Cópia (CC)
        tk.Label(main_frame, text="E-mail(s) para Cópia (CC):").pack(pady=(10, 5))
        self.cc_entry = tk.Entry(main_frame)
        self.cc_entry.insert(0, ", ".join(copia_emails))
        self.cc_entry.pack(fill="x", ipady=4)
       
        # Nome da Cidade
        tk.Label(main_frame, text="Nome da Cidade:").pack(pady=(10, 5))
        self.cidade_entry = tk.Entry(main_frame)
        self.cidade_entry.pack(fill="x", ipady=4)
       
        # Botão de Envio
        tk.Button(main_frame, text="Enviar E-mail", command=self.start_email_thread).pack(pady=20)
       
        # Log do Processo
        tk.Label(main_frame, text="Log do Processo:").pack(pady=(0, 5))
        self.log_text = tk.Text(main_frame, height=10, state="disabled")
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config('green', foreground='green')
        self.log_text.tag_config('red', foreground='red')
 
    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.caminho_pasta_pdfs = Path(folder_path)
            self.folder_path_label.config(text=str(self.caminho_pasta_pdfs))
            self.log("📂 Pasta selecionada. Mapeando arquivos PDF...", 'green')
            self.pdf_map = self.pre_process_pdfs(self.caminho_pasta_pdfs)
            if self.pdf_map:
                self.log(f"✅ {len(self.pdf_map)} arquivos PDF encontrados e mapeados.", 'green')
            else:
                self.log(f"❌ Nenhum arquivo PDF encontrado na pasta.", 'red')
 
    def log(self, message, tag=None):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"{message}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
 
    def pre_process_pdfs(self, caminho_pasta):
        pdf_map = {}
        try:
            pdf_files = list(Path(caminho_pasta).glob("*.pdf")) + list(Path(caminho_pasta).glob("*.PDF"))
            for pdf in pdf_files:
                nome_arquivo = pdf.stem
                cidade = nome_arquivo.split('_')[-1].strip()
                if cidade:
                    cidade_normalizada = normalizar_texto(cidade)
                    pdf_map[cidade_normalizada] = pdf
        except Exception as e:
            self.log(f"❌ Erro ao processar PDFs: {e}", 'red')
        return pdf_map
 
    def start_email_thread(self):
        # Inicia o envio em uma thread separada para não travar a GUI
        threading.Thread(target=self.send_email).start()
 
    def send_email(self):
        destinatario = self.destinatario_entry.get().strip()
        cidade_input = self.cidade_entry.get().strip()
        cc_list = [email.strip() for email in self.cc_entry.get().strip().split(',') if email.strip()]
 
        if not destinatario or not cidade_input or not self.caminho_pasta_pdfs:
            self.log("❌ Por favor, preencha todos os campos e selecione uma pasta.", 'red')
            messagebox.showwarning("Atenção", "Por favor, preencha todos os campos e selecione uma pasta.")
            return
 
        cidade_normalizada = normalizar_texto(cidade_input)
        anexo_path = self.pdf_map.get(cidade_normalizada)
 
        if not anexo_path:
            self.log(f"❌ Erro: Anexo para '{cidade_input}' não encontrado no mapa de PDFs. Verifique o nome do arquivo.", 'red')
            messagebox.showerror("Erro", f"Anexo para '{cidade_input}' não encontrado. Por favor, verifique o nome do arquivo na pasta.")
            return
 
        assunto = f"ENADE_2025_DISTRIBUIÇÃO - {cidade_input}"
       
        self.log(f"🚀 Iniciando envio para: {destinatario}", 'green')
        self.log(f"Assunto: {assunto}", 'green')
        self.log(f"Anexo: {anexo_path.name}", 'green')
       
        enviar_email_com_anexo(
            destinatario=destinatario,
            assunto=assunto,
            corpo_html=corpo_email_html,
            anexo_path=anexo_path,
            cc_list=cc_list,
            log_text=self.log_text
        )
 
if __name__ == "__main__":
    app = EmailApp()
    app.mainloop()