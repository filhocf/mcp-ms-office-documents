# Roadmap — mcp-ms-office-documents (fork filhocf)

## Vision

Transform mcp-ms-office-documents from a document **creation** tool into a complete document **lifecycle** tool: create, read, edit, convert, and merge — across all Office formats (DOCX, XLSX, PPTX, PDF).

## Sprints

| Sprint | Issue | Status | Description |
|--------|-------|--------|-------------|
| 0 | #1 | 🔲 Queue | PyPI publish + CI/CD automation |
| 1 | #2 | 🔲 Queue | Read/open existing documents (DOCX, XLSX, PPTX) |
| 2 | #3 | 🔲 Queue | Edit existing documents (DOCX, XLSX, PPTX) |
| 3 | #4 | 🔲 Queue | PDF generation and conversion |
| 4 | #5 | 🔲 Queue | Advanced DOCX (images, headers, layout, lists) |
| 5 | #6 | 🔲 Queue | Excel charts (line, bar, pie, waterfall) |
| 6 | #7 | 🔲 Queue | Advanced PPTX (edit, shapes, placeholders) |
| 7 | #8 | 🔲 Queue | Document merge and format conversion |
| — | #9 | 🔲 Queue | Templates with conditionals and loops (DOCX + PPTX) |
| — | #10 | 🔲 Queue | XLSX conditional formatting |

## Dependencies

```
S0 (PyPI) ──review──→ S1 (Read) ──→ S2 (Edit) ──→ S4 (DOCX adv) ──→ S7 (Merge)
                                                                    ↗
S3 (PDF) ──────────────────────────────────────────────────────────→ S7 (Merge)

S5 (Charts)     — independent
S6 (PPTX adv)   — independent
#9 (Templates)  — independent
#10 (XLSX fmt)  — independent
```

## Parallel tracks

6 independent tracks can run simultaneously:
- S0 (PyPI) + S3 (PDF) + S5 (Charts) + S6 (PPTX) + #9 (Templates) + #10 (XLSX fmt)

## Upstream relationship

- Fork: `filhocf/mcp-ms-office-documents`
- Upstream: `ForLegalAI/mcp-ms-office-documents`
- Strategy: implement in fork, submit PRs to upstream when ready
- Approved upstream: #55 (items 1, 3), #56 (PyPI)

## Conventions

- Each feature = 1 issue + 1 branch + 1 PR
- Branch naming: `feat/{issue-number}-{short-name}`
- Tests required for every new tool
- CI must pass before merge
