# Quick Start Guide 🚀

## O que foi criado

Baseado no arquivo `server.ts` existente que lê planilhas do Microsoft SharePoint/Excel, foram criados os seguintes arquivos para integração com Google Sheets:

### 📁 Novos Arquivos:

1. **`src/google-sheets-reader.ts`** - Leitor principal para Google Sheets
2. **`src/integrated-server.ts`** - Servidor que integra ambas as fontes (Microsoft + Google)
3. **`credentials.json.example`** - Exemplo do arquivo de credenciais do Google
4. **`GOOGLE_SHEETS_SETUP.md`** - Guia completo de configuração
5. **`QUICK_START.md`** - Este guia rápido

### 🔧 Arquivos Modificados:

- **`package.json`** - Adicionados novos scripts de execução

---

## 🏃‍♂️ Como executar rapidamente

### 1. Configurar Google Sheets (5 minutos)

1. **Criar Service Account no Google Cloud**
2. **Baixar arquivo JSON de credenciais**
3. **Renomear para `credentials.json`** e colocar na raiz do projeto
4. **Compartilhar sua planilha** com o email do service account
5. **Adicionar variáveis ao `.env`**:
   ```env
   GOOGLE_SPREADSHEET_ID=seu-id-da-planilha
   GOOGLE_SHEET_NAME=Sheet1
   GOOGLE_SHEET_RANGE=A:Z
   ```

> 📖 **Guia detalhado**: Veja `GOOGLE_SHEETS_SETUP.md` para instruções completas

### 2. Executar o código

```bash
# Instalar dependências (já feito)
npm install

# Opção 1: Apenas Google Sheets (porta 3001)
npm run dev:google-sheets

# Opção 2: Apenas Microsoft Excel (porta 3000)
npm run dev

# Opção 3: Servidor integrado - ambos (porta 3002)
npm run dev:integrated
```

---

## 🌐 Endpoints disponíveis

### Google Sheets Server (porta 3001)

- `GET http://localhost:3001/google-sheets` - Retorna dados da planilha

### Microsoft Excel Server (porta 3000)

- `GET http://localhost:3000/` - Retorna dados do SharePoint

### Servidor Integrado (porta 3002)

- `GET http://localhost:3002/` - Página inicial
- `GET http://localhost:3002/google-sheets` - Dados do Google Sheets
- `GET http://localhost:3002/microsoft-excel` - Dados do Microsoft Excel
- `GET http://localhost:3002/compare` - Compara ambas as fontes
- `GET http://localhost:3002/health` - Status dos serviços

---

## 📊 Formato dos dados

Ambos os leitores retornam dados no formato JSON estruturado:

```json
{
  "message": "Data retrieved successfully!",
  "totalRecords": 150,
  "data": [
    {
      "coluna1": "valor1",
      "coluna2": 123.45,
      "coluna3": "valor3"
    }
  ]
}
```

### ✨ Funcionalidades automáticas:

- **Detecção de cabeçalhos**: Primeira linha vira nomes das colunas
- **Conversão de tipos**: Números são automaticamente convertidos
- **Formato brasileiro**: Vírgulas decimais viram pontos
- **Limpeza de dados**: Remove linhas vazias e espaços extras

---

## 🔍 Teste rápido

```bash
# Testar Google Sheets
curl http://localhost:3001/google-sheets

# Testar comparação
curl http://localhost:3002/compare
```

---

## 🛠️ Scripts disponíveis

```bash
# Desenvolvimento
npm run dev                    # Microsoft Excel (porta 3000)
npm run dev:google-sheets      # Google Sheets (porta 3001)
npm run dev:integrated         # Ambos integrados (porta 3002)

# Produção
npm run build                  # Compilar TypeScript
npm run start                  # Microsoft Excel
npm run start:google-sheets    # Google Sheets
npm run start:integrated       # Servidor integrado
```

---

## 🆚 Comparação: Microsoft vs Google

| Recurso                    | Microsoft Excel  | Google Sheets      |
| -------------------------- | ---------------- | ------------------ |
| **Autenticação**           | Azure AD / MSAL  | Service Account    |
| **Arquivo de credenciais** | `.env`           | `credentials.json` |
| **API**                    | Microsoft Graph  | Google Sheets API  |
| **Configuração**           | Mais complexa    | Mais simples       |
| **Limites**                | Baseado no plano | Generosos          |

---

## ❓ Troubleshooting rápido

### Google Sheets não funciona?

1. ✅ Arquivo `credentials.json` existe?
2. ✅ Planilha foi compartilhada com o service account?
3. ✅ ID da planilha está correto no `.env`?
4. ✅ Google Sheets API está habilitada?

### Microsoft Excel não funciona?

1. ✅ Variáveis do Azure estão no `.env`?
2. ✅ App tem permissões no SharePoint?
3. ✅ IDs do documento estão corretos?

---

## 📞 Próximos passos

1. **Testar com suas planilhas reais**
2. **Customizar a transformação de dados** em `transformGoogleSheetsData()`
3. **Adicionar endpoints específicos** para suas necessidades
4. **Implementar cache** para melhor performance
5. **Adicionar validação de dados** de entrada

---

**🎉 Pronto! Agora você tem dois leitores de planilhas funcionando em paralelo!**
