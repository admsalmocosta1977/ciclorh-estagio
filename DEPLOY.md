# Como colocar o sistema online no Render

Tempo estimado: 20 minutos. Não precisa saber programar.

---

## Passo 1 — Criar conta no GitHub

1. Acesse https://github.com e clique em **Sign up**
2. Crie uma conta gratuita com seu e-mail

---

## Passo 2 — Enviar o código para o GitHub

1. No GitHub, clique em **+** (canto superior direito) → **New repository**
2. Nome: `ciclorh-estagio` | Privado (Private) | clique em **Create repository**
3. Na página seguinte, clique em **uploading an existing file**
4. Arraste **todos os arquivos** da pasta `SISTEMA_ESTAGIO` (exceto `estagio.db` e `__pycache__`)
5. Clique em **Commit changes**

---

## Passo 3 — Criar conta no Render

1. Acesse https://render.com e clique em **Get Started for Free**
2. Faça login com sua conta do GitHub (botão "GitHub")

---

## Passo 4 — Criar o banco de dados PostgreSQL

1. No painel do Render, clique em **New +** → **PostgreSQL**
2. Preencha:
   - **Name:** `ciclorh-db`
   - **Region:** South America (São Paulo) — se disponível, senão Ohio
   - **Plan:** Free
3. Clique em **Create Database**
4. Aguarde ficar verde (Status: Available)
5. **Copie** o valor de **Internal Database URL** — você vai precisar no próximo passo

---

## Passo 5 — Criar o serviço web

1. Clique em **New +** → **Web Service**
2. Conecte o repositório `ciclorh-estagio` do GitHub
3. Preencha:
   - **Name:** `ciclorh-estagio`
   - **Region:** mesmo do banco
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --workers 2 --timeout 60`
   - **Plan:** Free
4. Em **Environment Variables**, adicione:
   - `DATABASE_URL` → cole a **Internal Database URL** copiada no Passo 4
   - `SECRET_KEY` → digite qualquer sequência aleatória (ex: `CicloRH@2026#Seguro`)
5. Clique em **Create Web Service**

---

## Passo 6 — Inicializar o banco

Após o deploy concluir (fica verde):

1. No painel do serviço, clique em **Shell**
2. Digite e pressione Enter:
   ```
   python -c "from app import init_db; init_db()"
   ```
3. Pronto! O banco foi criado com o usuário admin padrão.

---

## Passo 7 — Acessar o sistema

- A URL será algo como: `https://ciclorh-estagio.onrender.com`
- **Login inicial:**
  - Usuário: `salmo`
  - Senha: `ciclorh2026`
- **⚠️ Troque a senha imediatamente** em Admin → Usuários → Editar

---

## Passo 8 — Criar os outros usuários

1. Acesse **Admin → Usuários → Novo Usuário**
2. Crie 3 contas com perfil **Operador**
3. Passe o link e as credenciais para cada pessoa

---

## Passo 9 — Importar a planilha atual

1. Acesse **Admin → Importar Planilha**
2. Selecione o arquivo `PLANILHA SALMO.xlsm`
3. Clique em **Importar**
4. Verifique os dados importados em **Contratos**

---

## Observações importantes

| Item | Detalhe |
|------|---------|
| Cold start | Após 15 min sem uso, a primeira abertura leva ~30 segundos |
| Banco gratuito | Expira após 90 dias — crie um novo e atualize DATABASE_URL |
| HTTPS | Automático — conexão sempre criptografada |
| Backup | Faça exportações periódicas via Admin → Importar (em breve) |
| Limite de tamanho de upload | 100 MB (suficiente para a planilha) |

---

## Renovar o banco a cada 90 dias

1. Render → New PostgreSQL → `ciclorh-db-v2`
2. No Shell do serviço: `python -c "from app import init_db; init_db()"`
3. Re-importar a planilha ou fazer backup antes de trocar
4. Atualizar `DATABASE_URL` no serviço web com a nova URL
