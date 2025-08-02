# Google Sheets Setup Guide

Este arquivo explica como configurar a integração com Google Sheets.

## Pré-requisitos

1. **Conta Google Cloud**: Você precisa de uma conta no Google Cloud Platform
2. **Projeto Google Cloud**: Crie um projeto ou use um existente
3. **Google Sheets API**: Habilite a API do Google Sheets no seu projeto

## Configuração do Google Cloud

### 1. Criar Service Account

1. Acesse o [Google Cloud Console](https://console.cloud.google.com/)
2. Selecione seu projeto
3. Vá para "IAM & Admin" > "Service Accounts"
4. Clique em "Create Service Account"
5. Preencha:
   - **Service account name**: `sheets-reader`
   - **Description**: `Service account for reading Google Sheets`
6. Clique em "Create and Continue"
7. Pule as permissões opcionais e clique em "Done"

### 2. Gerar Chave JSON

1. Na lista de Service Accounts, clique no email da conta criada
2. Vá para a aba "Keys"
3. Clique em "Add Key" > "Create new key"
4. Selecione "JSON" e clique em "Create"
5. O arquivo será baixado automaticamente
6. **Renomeie o arquivo para `credentials.json`** e coloque na raiz do projeto

### 3. Habilitar Google Sheets API

1. No Google Cloud Console, vá para "APIs & Services" > "Library"
2. Procure por "Google Sheets API"
3. Clique nela e depois em "Enable"

## Configuração da Planilha

### 1. Compartilhar a Planilha

1. Abra sua planilha no Google Sheets
2. Clique em "Share" (Compartilhar)
3. Adicione o email do Service Account (encontrado no arquivo credentials.json no campo `client_email`)
4. Dê permissão de **Viewer** (Visualizador)

### 2. Obter o ID da Planilha

O ID da planilha está na URL:

```
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit#gid=0
```

Copie a parte `SPREADSHEET_ID` da URL.

## Variáveis de Ambiente

Adicione as seguintes variáveis ao seu arquivo `.env`:

```env
# Google Sheets Configuration
GOOGLE_SPREADSHEET_ID=your-google-spreadsheet-id
GOOGLE_SHEET_NAME=Sheet1
GOOGLE_SHEET_RANGE=A:Z
```

### Explicação das Variáveis:

- **GOOGLE_SPREADSHEET_ID**: ID da planilha (extraído da URL)
- **GOOGLE_SHEET_NAME**: Nome da aba/sheet (padrão: "Sheet1")
- **GOOGLE_SHEET_RANGE**: Intervalo de células a serem lidas (padrão: "A:Z" para todas as colunas)

## Executando o Código

### 1. Instalar Dependências

```bash
npm install
```

### 2. Executar o Google Sheets Reader

```bash
npm run dev:google-sheets
```

### 3. Testar a API

```bash
curl http://localhost:3001/google-sheets
```

## Estrutura dos Dados

O código automaticamente:

1. **Detecta cabeçalhos**: A primeira linha é tratada como cabeçalho
2. **Converte tipos**: Números são automaticamente convertidos
3. **Suporta formato brasileiro**: Vírgulas são convertidas para pontos decimais
4. **Remove linhas vazias**: Linhas completamente vazias são ignoradas

### Exemplo de Saída:

```json
{
  "message": "Google Sheets data retrieved successfully!",
  "totalRecords": 150,
  "data": [
    {
      "nome": "João Silva",
      "idade": 30,
      "salario": 5000.5
    },
    {
      "nome": "Maria Santos",
      "idade": 25,
      "salario": 4200.0
    }
  ]
}
```

## Troubleshooting

### Erro: "Credentials file not found"

- Verifique se o arquivo `credentials.json` está na raiz do projeto
- Certifique-se de que o arquivo não está em `.gitignore`

### Erro: "Permission denied"

- Verifique se a planilha foi compartilhada com o service account
- Confirme que o email do service account está correto

### Erro: "Spreadsheet not found"

- Verifique se o GOOGLE_SPREADSHEET_ID está correto
- Confirme que a planilha existe e está acessível

### Erro: "Sheet not found"

- Verifique se o nome da aba está correto na variável GOOGLE_SHEET_NAME
- Os nomes são case-sensitive

## Segurança

⚠️ **Importante**:

- Nunca commite o arquivo `credentials.json` no seu repositório
- Adicione `credentials.json` ao `.gitignore`
- Use variáveis de ambiente em produção ao invés de arquivos de credenciais
