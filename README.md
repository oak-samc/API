# Sistema de Automação de E-mails ENADE 2025

## 📋 Descrição

Este projeto é um sistema de automação desenvolvido em Python para facilitar o envio de e-mails em massa com anexos PDF para coordenadores estaduais e de local do Exame Nacional de Desempenho dos Estudantes (ENADE 2025). O sistema automatiza o processo de distribuição de salas por município, permitindo que os coordenadores validem as informações de ensalamento.

## 🚀 Funcionalidades

- **Envio automatizado de e-mails** via Microsoft Outlook
- **Processamento inteligente de PDFs** com mapeamento por cidade
- **Template HTML personalizado** para validação de dados
- **Normalização de texto** para compatibilidade de nomes de cidades
- **Cópia para múltiplos destinatários** (CC)
- **Interface interativa** no terminal
- **Validação de anexos** antes do envio

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**
- **win32com.client** - Integração com Microsoft Outlook
- **pathlib** - Manipulação de caminhos de arquivos
- **unicodedata** - Normalização de texto
- **re** - Expressões regulares

## 📦 Requisitos

### Dependências Python
```bash
pip install pywin32
```

### Requisitos do Sistema
- Windows (obrigatório para integração com Outlook)
- Microsoft Outlook instalado e configurado
- Python 3.6 ou superior

## ⚙️ Configuração

### 1. Configuração da Pasta de PDFs
Edite a variável `caminho_pasta_pdfs` no arquivo `import.py`:

```python
caminho_pasta_pdfs = Path(r"C:\caminho\para\sua\pasta\de\pdfs")
```

### 2. Configuração de E-mails em Cópia
Edite a lista `copia_emails` para definir os destinatários em cópia:

```python
copia_emails = [
    "email1@exemplo.com",
    "email2@exemplo.com",
    "email3@exemplo.com"
]
```

### 3. Padrão de Nomenclatura dos Arquivos PDF
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

2. **Configure as variáveis** conforme descrito na seção de Configuração

3. **Execute o script**
   ```bash
   python import.py
   ```

4. **Siga as instruções interativas**:
   - Digite o e-mail do destinatário
   - Digite o nome da cidade correspondente ao PDF
   - O sistema encontrará automaticamente o arquivo correto
   - Digite `sair` para finalizar

## 📧 Template de E-mail

O sistema utiliza um template HTML completo que inclui:

- **Saudação personalizada** para coordenadores
- **Informações sobre o ENADE 2025** (data: 23 de novembro de 2025)
- **Tabela de validação** com campos para confirmação de:
  - Nome da instituição
  - Endereço completo
  - Número de salas e andares
  - Capacidade de candidatos
  - Agrupamento de blocos
  - Quantitativo de salas por bloco
  - Acessibilidade para atendimento especializado
- **Prazo para resposta**: 19 de setembro de 2025, às 09:00h
- **Contatos para dúvidas**

## 🔧 Funcionalidades Técnicas

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

- **Validação de e-mails** inválidos
- **Verificação de existência** de arquivos PDF
- **Tratamento de exceções** do Outlook
- **Mensagens informativas** para o usuário

## 📝 Exemplo de Uso

```
🚀 Iniciando o envio de e-mails de forma automatizada...

📂 3 arquivos PDF encontrados na pasta 'C:\PDFs\ENADE':
✅ Arquivo: 'distribuicao_BRASILIA.pdf' -> Cidade extraída: 'BRASILIA'
✅ Arquivo: 'distribuicao_SAO_PAULO.pdf' -> Cidade extraída: 'SAO PAULO'
✅ Arquivo: 'distribuicao_RIO_DE_JANEIRO.pdf' -> Cidade extraída: 'RIO DE JANEIRO'

🧭 Cidades disponíveis para envio:
- BRASILIA
- RIO DE JANEIRO
- SAO PAULO

Digite o e-mail do destinatário: coordenador@exemplo.com
Digite o nome da cidade para o anexo: Brasília

✅ E-mail enviado com sucesso para: coordenador@exemplo.com
```

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