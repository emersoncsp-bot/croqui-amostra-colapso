# Croqui · Amostra de Colapso

Aplicação React (Vite) para gerar croquis de tubo com a amostra de colapso e as
amostras adicionais (tensão residual e tração). Permite **inserção manual** ou
**importação de planilha** (.xlsx/.xls/.csv) — cada linha/IPPN gera uma página A4
em paisagem, pronta para imprimir ou salvar em PDF.

## Requisitos

- Node.js 18+ (recomendado 20+)

## Rodar localmente

```bash
npm install
npm run dev
```

Abra o endereço exibido no terminal (por padrão http://localhost:5173).

Para gerar a versão de produção:

```bash
npm run build      # gera a pasta dist/
npm run preview    # serve o build localmente para conferência
```

## Deploy na Vercel

A Vercel detecta o Vite automaticamente. Há dois caminhos:

### A) Pelo painel (recomendado)

1. Suba este projeto para um repositório no GitHub/GitLab/Bitbucket.
2. Em https://vercel.com → **Add New… → Project** e importe o repositório.
3. A Vercel preenche sozinha:
   - **Framework Preset:** Vite
   - **Build Command:** `npm run build`
   - **Output Directory:** `dist`
4. Clique em **Deploy**.

### B) Pela CLI

```bash
npm i -g vercel
vercel          # primeiro deploy (responda às perguntas)
vercel --prod   # publica em produção
```

## Importação de planilha

O leitor reconhece as colunas por nome (tolerante a acentos/maiúsculas/espaços).
Colunas usadas:

| Campo                       | Coluna da planilha               |
| --------------------------- | -------------------------------- |
| OD                          | `OD`                             |
| WT                          | `WT`                             |
| Grau do aço                 | `Grau do aço`                    |
| Comprimento do tubo         | `Comprimento do tubo MES`        |
| Pressão de colapso (Psi)    | `Pressão Spec (Psi)`             |
| Pedido/Item                 | `Pedido/Item`                    |
| Ordem de Produção           | `Ordem de Produção`              |
| Posição do colapso (mm)     | `Posição High Collapse(mm) ...`  |
| Posição (Pé/Meio/Ponta)     | `Posição`                        |
| IPPN                        | `IPPN`                           |
| Tipo-amostra                | `Tipo-amostra`                   |

O **tamanho da amostra de colapso** é definido por OD + pressão:

- OD 114,3–273,1 (qualquer pressão): 2800 mm
- OD 298,4–355,6 e pressão ≤ 11600 psi: 3600 mm
- OD > 355,6 (qualquer pressão): 5400 mm
- OD 273,1–473,1 e pressão > 11600 psi: 5400 mm

A tensão residual = 3 × OD (arredondado à dezena superior) e a tração = 380 mm.

## Estrutura

```
.
├── index.html
├── package.json
├── vite.config.js
└── src
    ├── App.jsx     # toda a aplicação (UI + croqui em SVG + importação)
    ├── main.jsx
    └── index.css
```
