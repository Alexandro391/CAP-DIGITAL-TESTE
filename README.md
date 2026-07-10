# Sistema-de-Atendimento-CAP

API REST para o sistema de chamados da Central de Apoio Pedagógico.

## Arquivos

```
cap-backend/
├── server.js       ← código principal
├── package.json
├── .gitignore
└── data/           ← criada automaticamente
    └── chamados.json
```

---

## Deploy no Render.com (passo a passo)

### 1. Crie um repositório no GitHub

1. Acesse github.com → **New repository**
2. Nome: `cap-backend` (privado ou público)
3. Faça upload dos arquivos: `server.js`, `package.json`, `.gitignore`

### 2. Crie o serviço no Render

1. Acesse **render.com** → **New** → **Web Service**
2. Conecte ao repositório `cap-backend`
3. Configure:

| Campo | Valor |
|---|---|
| **Name** | cap-backend |
| **Runtime** | Node |
| **Build Command** | `npm install` |
| **Start Command** | `npm start` |
| **Instance Type** | Free |

4. Na seção **Environment Variables**, adicione:

| Chave | Valor |
|---|---|
| `ADMIN_TOKEN` | sua-senha-secreta (escolha uma senha forte) |

5. Clique em **Create Web Service**

Após ~2 minutos, seu backend estará em:
`https://cap-backend.onrender.com`

---

## Configurar o dashboard

Abra o `dashboard.html` e preencha:

- **URL da API**: `https://cap-backend.onrender.com/api`
- **Token Administrativo**: a senha que você definiu em `ADMIN_TOKEN`

---

## Configurar o formulário (index.html)

Localize no `index.html` a linha com `API_URL` e substitua pelo endereço do Render:

```javascript
const API_URL = 'https://cap-backend.onrender.com/api/chamados';
```

---

## Rotas da API

| Método | Rota | Auth | Descrição |
|---|---|---|---|
| `POST` | `/api/chamados` | Não | Criar chamado (formulário público) |
| `GET` | `/api/chamados` | Bearer token | Listar chamados (dashboard) |
| `POST` | `/api/chamados/:protocolo/status` | Bearer token | Atualizar status |
| `DELETE` | `/api/chamados/:protocolo` | Bearer token | Remover chamado |
| `GET` | `/health` | Não | Status da API |

---

## Atenção — Free tier do Render

No plano gratuito, o serviço **hiberna após 15 minutos sem uso**.
A primeira requisição após a hibernação pode demorar ~30 segundos.

Para evitar isso, considere usar o plano pago ou um serviço de ping periódico
como **UptimeRobot** (gratuito) apontando para `https://cap-backend.onrender.com/health`.
