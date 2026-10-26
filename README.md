# Painel de Baixas & Entregas

Dashboard de logística (React + Vite). Duas telas:
- `/` — painel gerencial completo (upload da planilha, KPIs, comparativos, financeiro)
- `/parceiro/:slug` — tela individual de cada transportadora (gerada automaticamente a partir do nome dela)

## Como os dados são compartilhados

Os dados carregados no painel gerencial ficam salvos num banco de dados (Upstash Redis, plano
gratuito) através de uma função serverless da própria Vercel (`/api/data`). Isso significa que
**qualquer pessoa que abrir o link — gerencial ou de qualquer parceiro — vê a mesma versão mais
recente dos dados**, não importa o navegador ou computador.

Para isso funcionar, você precisa criar uma conta gratuita no Upstash e configurar duas variáveis
de ambiente na Vercel — veja o passo 4 abaixo. **Sem isso, o upload continua funcionando, mas os
dados não ficam salvos entre uma visita e outra.**

## Passo a passo — GitHub

1. Instale o [Git](https://git-scm.com/downloads) e crie uma conta no [GitHub](https://github.com) se ainda não tiver.
2. Crie um repositório novo em github.com → botão **New repository** → dê um nome (ex: `painel-logistica`) → **não** marque "Add a README" → **Create repository**.
3. No seu computador, dentro da pasta deste projeto, rode:
   ```bash
   git init
   git add .
   git commit -m "Primeira versão do painel"
   git branch -M main
   git remote add origin https://github.com/SEU-USUARIO/painel-logistica.git
   git push -u origin main
   ```
   (troque `SEU-USUARIO` e o nome do repositório pelos seus)

## Passo a passo — Vercel

1. Crie uma conta em [vercel.com](https://vercel.com) — dá pra entrar direto com a conta do GitHub.
2. No painel da Vercel, clique em **Add New → Project**.
3. Selecione o repositório `painel-logistica` (a Vercel pede permissão para ler seus repositórios do GitHub na primeira vez).
4. **Antes de clicar em Deploy**, configure o banco de dados compartilhado (veja a seção abaixo) e adicione as variáveis de ambiente. Se preferir, pode fazer isso depois do primeiro deploy também — só vai precisar redeployar.
5. A Vercel detecta sozinha que é um projeto **Vite** — não precisa mudar Build Command (`npm run build`) nem Output Directory (`dist`).
6. Clique em **Deploy** e aguarde (1-2 minutos).
7. Ao terminar, você tem um link tipo `https://painel-logistica.vercel.app` — esse é o painel gerencial.
8. Os links de cada parceiro ficam em `https://painel-logistica.vercel.app/parceiro/NOME-DA-TRANSPORTADORA` — pegue o link certo direto na aba **"Portal do Parceiro"** dentro do painel (tem botão de copiar).

## Configurando o banco de dados compartilhado (Upstash — gratuito)

1. Crie uma conta grátis em [upstash.com](https://upstash.com) (dá pra entrar com GitHub).
2. No painel da Upstash, clique em **Create Database**.
   - Nome: qualquer um (ex: `painel-logistica`)
   - Tipo: **Regional** (mais barato/simples que Global, e é suficiente aqui)
   - Região: escolha uma perto de onde a Vercel vai hospedar (ex: `us-east-1` se a Vercel estiver nos EUA — a região exata não é crítica para este uso)
   - Clique em **Create**
3. Na página do banco criado, role até **REST API** e copie dois valores:
   - `UPSTASH_REDIS_REST_URL`
   - `UPSTASH_REDIS_REST_TOKEN`
4. Na Vercel, entre no seu projeto → **Settings → Environment Variables** e adicione as duas:
   - Nome: `UPSTASH_REDIS_REST_URL` — Valor: (o que você copiou)
   - Nome: `UPSTASH_REDIS_REST_TOKEN` — Valor: (o que você copiou)
   - Marque para os três ambientes (Production, Preview, Development)
5. Vá em **Deployments**, abra os "..." do último deploy e clique em **Redeploy** (as variáveis de ambiente só entram em vigor num novo deploy).

Pronto — a partir daí, todo upload feito no painel gerencial fica visível para qualquer pessoa que
acessar o link, seja o painel gerencial ou a tela de um parceiro específico.

### Limite de tamanho dos dados

O plano gratuito da Upstash tem um limite de tamanho por requisição (na casa de algumas centenas
de KB). Para o volume de notas que vimos até agora isso não é problema, mas se a planilha crescer
muito (dezenas de milhares de notas pendentes de uma vez), pode ser necessário reduzir os campos
salvos por nota ou passar para um plano pago. Se isso acontecer, me avise.

## Atualizações futuras

Sempre que você editar o código e quiser publicar de novo:
```bash
git add .
git commit -m "Descrição da mudança"
git push
```
A Vercel publica automaticamente a cada `push` na branch `main`.

## Rodando localmente antes de publicar (opcional)

A função `/api/data` só funciona hospedada na Vercel (ou com `vercel dev`, que simula o ambiente
da Vercel localmente). Rodando só com `npm run dev`, o upload vai falhar ao tentar salvar — isso é
esperado, é só para conferir a interface antes de publicar.

```bash
npm install
npm run dev
```
Abre em `http://localhost:5173`.
