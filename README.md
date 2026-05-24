# Conecta General Bruce

**Serviços, produtos e talentos bem pertinho de você.**

Plataforma comunitária gratuita que conecta moradores, prestadores de serviço e pequenos empreendedores da Rua General Bruce e arredores.

---

## Sobre o projeto

Iniciativa de extensão universitária do curso de **Análise e Desenvolvimento de Sistemas**, alinhada à Agenda 2030 da ONU (ODS 8 e ODS 11).

O Conecta General Bruce reúne em um só lugar os serviços e produtos oferecidos por pessoas da comunidade local, facilitando contratações entre vizinhos e fortalecendo a economia da região.

---

## Estrutura de arquivos

```
/
├── index.html          ← Site completo (HTML + CSS + JS em arquivo único)
├── condominio.jpg      ← Imagem de fundo da seção hero
├── og-image.jpg        ← Imagem de preview para redes sociais (1200×630 px)
├── favicon.ico         ← Ícone da aba do navegador
├── vercel.json         ← Configuração de redirecionamento da URL antiga
└── README.md           ← Este arquivo
```

---

## Como rodar localmente

1. Baixe ou clone este repositório
2. Abra o arquivo `index.html` diretamente no navegador (duplo clique)
3. O site funciona sem servidor — exceto o formulário de cadastro, que depende do Google Apps Script

---

## Backend (Google Apps Script)

O formulário de cadastro envia dados para uma planilha Google Sheets via Apps Script.

- Script: `Codigo_AppsScript.gs`
- A URL do script fica na variável `APPS_SCRIPT_URL` no topo do `<script>` do `index.html`
- Novos cadastros chegam com status **PENDENTE** e só aparecem no site após aprovação manual na planilha

---

## Deploy (Vercel)

- **URL principal:** https://conectageneralbruce.vercel.app
- **URL legada (redireciona):** vitrinedomoradacarioca-78pi3ivts-bereniceahb-tis-projects.vercel.app
- O arquivo `vercel.json` configura o redirecionamento 301 da URL antiga para a nova

---

## Tecnologias

- HTML5, CSS3 e JavaScript puro (sem frameworks)
- Google Apps Script + Google Sheets (backend/banco de dados)
- Vercel (hospedagem)

---

## Alinhamento ODS

| ODS | Descrição |
|-----|-----------|
| 🎯 ODS 8 | Trabalho decente e crescimento econômico |
| 🏙️ ODS 11 | Cidades e comunidades sustentáveis |

---

*Feito com 💙 pela comunidade — 2026*
