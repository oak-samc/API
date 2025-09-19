import os
import time
import re
import unicodedata
import win32com.client as win32
from pathlib import Path

# === CONFIGURAÇÕES ===
# Verifique se o caminho da pasta está correto para o seu sistema.
caminho_pasta_pdfs = Path(r"C:\Users\victor.vasconcelos\Documents\ENVIAR E-MAIL PR")
copia_emails = ["alinemilacki@gmail.com", "yasmin.oliveira@cebraspe.org.br", "enade2025@cebraspe.org.br"] # Insira os 3 e-mails aqui

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
 
<p>Para garantir a qualidade e a excelência do nosso trabalho e cumprir os prazos estabelecidos, solicitamos a resposta a esse e-mail até o dia 19 de setembro de 2025, às 09:00h (horário de Brasília).</p>
 
<p>Em caso de dúvidas, entre em contato com o Cebraspe pelo e-mail <a href="mailto:enade2025@cebraspe.org.br">enade2025@cebraspe.org.br</a> ou telefone (61) 2109-5810.</p>
"""

# === FUNÇÕES ===

def normalizar_texto(texto):
    """Remove acentos e deixa tudo maiúsculo."""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto.strip().upper()

def pre_process_pdfs(caminho_pasta):
    """Mapeia PDFs com base no nome da cidade extraída do nome do arquivo."""
    
    if not caminho_pasta.is_dir():
        print(f"❌ ERRO: A pasta '{caminho_pasta}' não foi encontrada. Por favor, verifique o caminho.")
        return {}

    pdf_files = list(Path(caminho_pasta).glob("*.pdf")) + list(Path(caminho_pasta).glob("*.PDF"))
    pdf_map = {}

    print(f"\n📂 {len(pdf_files)} arquivos PDF encontrados na pasta '{caminho_pasta}':")
    for pdf in pdf_files:
        nome_arquivo = pdf.stem  # sem extensão
        
        # Pega a última parte do nome do arquivo após o último sublinhado
        cidade = nome_arquivo.split('_')[-1].strip()

        if cidade:
            cidade_normalizada = normalizar_texto(cidade)
            pdf_map[cidade_normalizada] = pdf
            print(f"✅ Arquivo: '{pdf.name}' -> Cidade extraída: '{cidade}' -> Normalizada: '{cidade_normalizada}'")
        else:
            print(f"⚠️ Arquivo ignorado (sem padrão): {pdf.name}")
    
    if not pdf_files:
        print("❌ NENHUM arquivo PDF encontrado. Verifique se a pasta está correta.")

    print("\n✅ Pré-processamento finalizado.")
    return pdf_map

def enviar_email_com_anexo(destinatario, assunto, corpo_html, anexo_path, cc_list=None):
    """Envia um e-mail via Outlook com anexo."""
    try:
        if not destinatario or "@" not in destinatario:
            print(f"❌ E-mail inválido: {destinatario}")
            return
        
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)

        mail.To = destinatario
        mail.Subject = assunto
        mail.HTMLBody = corpo_html

        if cc_list:
            mail.CC = "; ".join(cc_list)

        if anexo_path and os.path.exists(anexo_path):
            mail.Attachments.Add(str(anexo_path))
            print(f"📎 Anexo adicionado: {anexo_path}")
        else:
            print(f"❌ Erro: Anexo não encontrado ou caminho inválido: {anexo_path}. O e-mail não será enviado.")
            return

        mail.Send()
        print(f"✅ E-mail enviado com sucesso para: {destinatario}")

    except Exception as e:
        print(f"❌ ERRO AO ENVIAR O E-MAIL: {e}")
        print("Tente verificar se o Outlook está aberto e é o cliente de e-mail padrão do Windows.")

def main_automatizado():
    """Fluxo automatizado com entrada contínua do usuário."""
    print("\n🚀 Iniciando o envio de e-mails de forma automatizada...")
    pdf_map = pre_process_pdfs(caminho_pasta_pdfs)
    if not pdf_map:
        return

    print("\n🧭 Cidades disponíveis para envio:")
    for cidade in sorted(pdf_map.keys()):
        print(f"- {cidade}")

    print("\n" + "="*50)
    print("Digite 'sair' a qualquer momento para finalizar o processo.")
    print("="*50)

    while True:
        destinatario = input("\nDigite o e-mail do destinatário: ").strip()
        if destinatario.lower() == "sair":
            break

        cidade_input = input("Digite o nome da cidade para o anexo: ").strip()
        if cidade_input.lower() == "sair":
            break
        
        cidade_normalizada = normalizar_texto(cidade_input)
        anexo_path = pdf_map.get(cidade_normalizada)

        if not anexo_path:
            print(f"\n❌ Erro: Anexo para '{cidade_input}' (normalizado: '{cidade_normalizada}') não encontrado.")
            print("Por favor, verifique o nome do arquivo na pasta e tente novamente.")
            continue

        assunto = f"ENADE_2025_DISTRIBUIÇÃO - {cidade_input}"
        
        print("\n" + "="*50)
        print("Detalhes do envio:")
        print(f"  Destinatário: {destinatario}")
        print(f"  Assunto: {assunto}")
        print(f"  Anexo: {anexo_path.name}")
        print("="*50)
        
        enviar_email_com_anexo(
            destinatario=destinatario,
            assunto=assunto,
            corpo_html=corpo_email_html,
            anexo_path=anexo_path,
            cc_list=copia_emails
        )
        time.sleep(2) # Pequena pausa para evitar sobrecarga do Outlook

    print("\n🏁 Processo de envio de e-mail concluído.")


# === EXECUÇÃO ===
if __name__ == "__main__":
    main_automatizado()
