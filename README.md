# Read Excel Test

Uma aplicação Node.js/TypeScript que conecta ao Microsoft Graph API para ler dados de planilhas Excel armazenadas no SharePoint/OneDrive.

## 📋 Propósito

Esta aplicação demonstra como:

- Autenticar com o Microsoft Graph API usando Client Credentials Flow
- Conectar-se a planilhas Excel no SharePoint/OneDrive
- Extrair dados de células e tabelas específicas
- Processar informações de planilhas de forma programática

## 🚀 Funcionalidades

- **Autenticação Azure AD**: Usa credenciais de aplicação para acessar recursos do Microsoft 365
- **Leitura de Excel**: Acessa dados de planilhas específicas no SharePoint
- **API REST**: Servidor Fastify para demonstrar a integração
- **TypeScript**: Código tipado para melhor manutenibilidade

## ⚙️ Configuração

1. **Instale as dependências:**

   ```bash
   npm install
   ```

2. **Configure as variáveis de ambiente:**
   Crie um arquivo `.env` baseado no `.env.example` com:

   ```
   CLIENT_ID=seu_client_id_azure
   CLIENT_SECRET=seu_client_secret_azure
   TENANT_ID=seu_tenant_id_azure
   ```

3. **Execute a aplicação:**
   ```bash
   npm run dev
   ```

## 📊 Uso

A aplicação conecta automaticamente à planilha Excel configurada e exibe os dados no console quando iniciada. O servidor ficará disponível em `http://localhost:3000`.

## 🛠️ Tecnologias

- **Node.js** + **TypeScript**
- **Fastify** - Framework web rápido
- **@azure/msal-node** - Autenticação Microsoft
- **Axios** - Cliente HTTP para chamadas à API
- **Microsoft Graph API** - Acesso aos dados do Office 365
