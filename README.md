# Sistema de Automação de E-mails ENADE 2025

## 📋 Descrição

Este projeto é um sistema de automação desenvolvido em Python para facilitar o envio de e-mails em massa com anexos PDF para coordenadores estaduais e de local do Exame Nacional de Desempenho dos Estudantes (ENADE 2025). O sistema automatiza o processo de distribuição de salas por município, permitindo que os coordenadores validem as informações de ensalamento.

## 🚀 Funcionalidades

- **Interface gráfica intuitiva** com tkinter para facilitar o uso
- **Seleção de pasta via GUI** para localizar arquivos PDF
- **Envio automatizado de e-mails** via Microsoft Outlook
- **Processamento inteligente de PDFs** com mapeamento automático por cidade
- **Template HTML personalizado** para validação de dados
- **Normalização de texto** para compatibilidade de nomes de cidades
- **Cópia para múltiplos destinatários** (CC) editável
- **Log em tempo real** do processo de envio na interface
- **Validação de campos** obrigatórios antes do envio
- **Validação de anexos** antes do envio
- **Processamento em background** para não travar a interface
- **Tratamento de erros** com mensagens informativas

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**
- **tkinter** - Interface gráfica do usuário (GUI)
- **win32com.client** - Integração com Microsoft Outlook
- **pathlib** - Manipulação de caminhos de arquivos
- **unicodedata** - Normalização de texto
- **threading** - Processamento em background
- **re** - Expressões regulares
- **os** - Operações do sistema operacional
- **time** - Controle de tempo

## 📦 Requisitos

### Dependências Python
```bash
pip install pywin32
```

**Nota**: O `tkinter` já vem incluído na instalação padrão do Python.

### Requisitos do Sistema
- Windows (obrigatório para integração com Outlook)
- Microsoft Outlook instalado e configurado
- Python 3.6 ou superior

## ⚙️ Configuração

### 1. Configuração de E-mails em Cópia
Edite a lista `copia_emails` no arquivo `import.py` para definir os destinatários em cópia:

```python
copia_emails = [
    "email1@exemplo.com",
    "email2@exemplo.com",
    "email3@exemplo.com"
]
```

### 2. Padrão de Nomenclatura dos Arquivos PDF
Os arquivos PDF devem seguir o padrão:
```
qualquer_nome_CIDADE.pdf
```

Exemplo:
- `distribuicao_salas_SAO_PAULO.pdf`
- `enade_BRASILIA.pdf`

## 🎯 Como Usar

1. **Clone ou baixe o projeto**
   ```bash
   git clone https://github.com/VicorVasconcelos/API.git
   cd API
   ```

2. **Instale as dependências**
   ```bash
   pip install pywin32
   ```

3. **Execute o script**
   ```bash
   python import.py
   ```

4. **Use a interface gráfica**:
   - **Selecione a pasta de PDFs**: Clique em "Selecionar Pasta" para escolher onde estão os arquivos PDF
   - **Digite o e-mail do destinatário**: Informe o e-mail de quem receberá a mensagem
   - **Configure os e-mails de cópia (CC)**: Os e-mails padrão já estarão preenchidos, mas você pode editá-los
   - **Digite o nome da cidade**: Informe a cidade correspondente ao PDF que será anexado
   - **Clique em "Enviar E-mail"**: O sistema encontrará automaticamente o arquivo correto e enviará
   - **Acompanhe o processo**: Use o log em tempo real para ver o status do envio

## 🔧 Funcionalidades Técnicas

### Interface Gráfica
- **Janela principal** com layout organizado e intuitivo
- **Seleção de pasta** via dialog nativo do sistema
- **Campos de entrada** validados antes do envio
- **Log em tempo real** com cores para diferentes tipos de mensagem
- **Processamento em background** usando threading para não travar a interface
- **Mensagens de feedback** usando messageboxes do tkinter

### Normalização de Texto
O sistema remove acentos e converte para maiúsculas para garantir compatibilidade:
```python
def normalizar_texto(texto):
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto.strip().upper()
```

### Mapeamento de PDFs
Extrai automaticamente o nome da cidade do arquivo PDF e cria um mapeamento:
```
BRASILIA.pdf → "BRASILIA"
sao_paulo.pdf → "SAO PAULO"
```

### Validação de Anexos
Verifica se o arquivo existe antes de enviar o e-mail, evitando envios sem anexo.

## 🛡️ Tratamento de Erros

- **Validação de e-mails** inválidos com messageboxes informativos
- **Verificação de existência** de arquivos PDF antes do envio
- **Validação de campos obrigatórios** na interface gráfica
- **Tratamento de exceções** do Outlook com mensagens detalhadas
- **Log colorido** para diferentes tipos de mensagem (sucesso, erro, informação)
- **Feedback visual** em tempo real durante o processamento

## 📝 Exemplo de Uso

### Interface Gráfica do Sistema

Ao executar o programa, uma janela intitulada **"Automatizador de E-mails ENADE - v2.0"** será aberta com os seguintes campos:

```
📂 Selecione a pasta com os PDFs: [Selecionar Pasta]
Nenhuma pasta selecionada

E-mail do Destinatário: [Campo de texto]

E-mail(s) para Cópia (CC): [Campo pré-preenchido com os e-mails padrão]

Nome da Cidade: [Campo de texto]

[Enviar E-mail]

Log do Processo:
📂 Pasta selecionada. Mapeando arquivos PDF...
✅ 3 arquivos PDF encontrados e mapeados.
🚀 Iniciando envio para: coordenador@exemplo.com
Assunto: ENADE_2025_DISTRIBUIÇÃO - Brasília
📎 Anexo adicionado: C:\PDFs\distribuicao_BRASILIA.pdf
✅ E-mail enviado para: coordenador@exemplo.com
```

### Fluxo de Trabalho
1. O usuário executa `python import.py`
2. A interface gráfica é aberta
3. O usuário clica em "Selecionar Pasta" e escolhe a pasta com os PDFs
4. O sistema mapeia automaticamente todos os arquivos PDF encontrados
5. O usuário preenche o e-mail do destinatário e nome da cidade
6. O usuário clica em "Enviar E-mail"
7. O sistema processa em background e mostra o progresso no log
8. Uma mensagem de sucesso ou erro é exibida

## 🤝 Contribuição

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 👨‍💻 Autor

**Victor Vasconcelos**
- GitHub: [@VicorVasconcelos](https://github.com/VicorVasconcelos)

## 📞 Suporte

Para dúvidas ou problemas:
- Abra uma [issue](https://github.com/VicorVasconcelos/API)
- Entre em contato pelo e-mail: victor.vasconcelos@cebraspe.org.br ou victorvasconcellos28@gmail.com
- Telefone: (61) 98438-5187

---

**⚠️ Nota Importante**: Este sistema foi desenvolvido especificamente para o ENADE 2025 e requer Microsoft Outlook instalado no Windows para funcionar corretamente.