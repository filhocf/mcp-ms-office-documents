# MEMORY.md — mcp-ms-office-documents

## Estado Atual

- **Fork**: filhocf/mcp-ms-office-documents
- **Upstream**: ForLegalAI/mcp-ms-office-documents
- **Branch**: master
- **10 issues criadas** (#1-#10)
- **Upstream aprovado**: #55 (templates DOCX+PPTX), #56 (PyPI publish)
- **Próximo**: Sprint 0 (#1) — pyproject.toml + CI/CD

## Contexto Técnico

- Stack: Python + FastMCP 3.2 + python-docx + python-pptx + openpyxl
- Testes: 285 passando (pytest, 6s)
- CI: GitHub Actions (ci.yml) — atualmente só workflow_dispatch
- Roda em: Streamable HTTP porta 8958
- Venv local: ~/git/mcp-ms-office-documents/.venv

## Upstream Maintainer

- User: dvejsada
- Receptivo: sim (respondeu <12h)
- Preferência: tool schema limitado (evitar LLMs errando com modelos menores)
- Aprovações: templates DOCX (#55.1), templates PPTX (#55.3), PyPI (#56)
- Pendente: XLSX formatting (#55.2) — proposta de rules array enviada

## Decisões

- Implementar tudo no fork primeiro, submeter upstream depois
- Cada feature = 1 issue + 1 branch + 1 PR
- Se upstream não aceitar algo, mantemos no fork
