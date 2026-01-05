# Conversor XML para Excel ⚡️

**Projeto:** Vite + React + TypeScript

Um aplicativo leve para importar arquivos XML, visualizar os dados em tabela, detectar/remover duplicatas e exportar para Excel (`.xlsx`). Ideal para transformar relatórios XML em planilhas editáveis. 🔧

---

## 🚀 Recursos principais

- Upload de arquivos XML via *drag & drop* ou seletor de arquivos
- Parser de XML para JSON (extração de campos relevantes)
- Visualização em tabela com detecção de duplicatas
- Exportação para Excel (`.xlsx`) usando a biblioteca `xlsx`
- UI responsiva com componentes reutilizáveis

---

## 🧭 Tecnologias

- Vite
- React
- TypeScript
- Tailwind CSS
- XLSX (exportação para Excel)

---

## 🔧 Requisitos

- Node.js (versão LTS recomendada)
- npm (ou pnpm/yarn)

---

## 📦 Instalação

```bash
# clonar o repositório
git clone <URL-do-repositório>
cd "conversor XML"

# instalar dependências
npm install
```

---

## ▶️ Scripts úteis

- `npm run dev` — Inicia o servidor de desenvolvimento (Vite)
- `npm run build` — Gera a build de produção
- `npm run build:dev` — Build em modo development
- `npm run preview` — Pré-visualiza a build gerada
- `npm run lint` — Executa o ESLint

---

## ✅ Como usar

1. Execute `npm run dev`.
2. Abra o navegador em `http://localhost:5173`.
3. Faça upload do arquivo XML (arrastar ou clicar no seletor).
4. Revise os dados na tabela, remova duplicatas se necessário.
5. Clique em **Exportar** para gerar o arquivo `.xlsx`.

> Dica: a interface contém botões para localizar e resolver duplicatas antes da exportação.

---

## 🗂 Estrutura do projeto (resumida)

- `src/components/` — componentes da UI (upload, tabela, botões)
- `src/lib/` — utilitários (parser XML, exportação para Excel)
- `src/pages/` — páginas (Index, NotFound)
- `public/` — arquivos estáticos

---

## 🤝 Contribuindo

Contribuições são bem-vindas! Você pode:

1. Abrir uma issue descrevendo o problema ou a feature.
2. Criar um *fork* e enviar um pull request com as mudanças.

Por favor siga as regras de estilo de código do projeto e adicione testes/descrições quando relevante.

---

## 📝 Licença

Sem licença especificada neste repositório. Se desejar, adicione um arquivo `LICENSE` (por exemplo, MIT) para tornar a licença explícita.

---

## ✉️ Contato

Se precisar de ajuda, abra uma issue ou deixe uma mensagem no repositório.

---

**Bom trabalho!** Se quiser, eu posso também: adicionar um arquivo `LICENSE`, ajustar o texto para um README mais curto, ou incluir instruções para Docker/CI/CD. 🚀