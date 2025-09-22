# Sistema de Automação de E-mails

## 📋 Descrição

Sistema desenvolvido em Python para automatizar o envio de e-mails em massa com anexos PDF. Com interface gráfica intuitiva e integração direta com Microsoft Outlook.

## 🚀 Principais Funcionalidades

- **Interface gráfica intuitiva** com tkinter
- **Seleção de pasta via GUI** para localizar arquivos PDF
- **Envio automatizado** via Microsoft Outlook
- **Mapeamento automático de PDFs** por cidade
- **Template HTML personalizado** para validação
- **Log em tempo real** do processo de envio
- **Processamento em background** sem travar a interface
- **Validação completa** de campos e anexos

## 🛠️ Tecnologias Utilizadas

- **Python 3.x** - Linguagem principal
- **tkinter** - Interface gráfica
- **pywin32** - Integração com Outlook
- **threading** - Processamento em background
- **pathlib, unicodedata** - Manipulação de arquivos e texto

## 📦 Requisitos

### Sistema
- Windows (obrigatório para integração com Outlook)
- Microsoft Outlook instalado e configurado
- Python 3.6 ou superior

### Dependências
```bash
pip install pywin32
```

## ⚙️ Configuração

### 1. E-mails de Cópia
Edite a lista `copia_emails` no arquivo `import.py`:
```python
copia_emails = [
    "email1@exemplo.com",
    "email2@exemplo.com"
]
```

### 2. Padrão dos PDFs
Os arquivos devem seguir o formato: `qualquer_nome_CIDADE.pdf`

Exemplos:
- `distribuicao_salas_SAO_PAULO.pdf`
- `enade_BRASILIA.pdf`

## 🎯 Como Usar

1. **Instalação**
   ```bash
   git clone https://github.com/VicorVasconcelos/API.git
   cd API
   pip install pywin32
   ```

2. **Execução**
   ```bash
   python import.py
   ```

3. **Interface**
   - Selecione a pasta com os PDFs
   - Digite o e-mail do destinatário
   - Configure os e-mails de cópia (CC)
   - Informe o nome da cidade
   - Clique em "Enviar E-mail"
   - Acompanhe o progresso no log

## 🛡️ Funcionalidades Técnicas

- **Normalização de texto** para compatibilidade de nomes
- **Validação de anexos** antes do envio
- **Tratamento de erros** com mensagens detalhadas
- **Log colorido** para diferentes tipos de mensagem

## 👨‍💻 Autores

**Victor Vasconcelos** e **Samuel Almeida**
- GitHub: [@VicorVasconcelos](https://github.com/VicorVasconcelos)
- GitHub: [@oak-samc](https://github.com/oak-samc)

## 📞 Suporte

- E-mail: victor.vasconcelos@cebraspe.org.br
- Telefone: (61) 98438-5187
- Issues: [GitHub Issues](https://github.com/VicorVasconcelos/API/issues)

---

**⚠️ Nota**: Sistema desenvolvido especificamente para o ENADE 2025. Requer Microsoft Outlook no Windows.
