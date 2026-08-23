# Spec-to-Deck Pipeline

Rerunnable spec-JSON → PPTX → PDF deck pipeline (pptxgenjs). Part of the **Paul's Tradecraft** portfolio work.

## How it works

- `spec-*.json` — declarative deck spec (slides, layouts, copy)
- `make-pptx.mjs <spec.json>` — renders the spec to a `.pptx` (v4 editorial design system, McKinsey-style)
- PPTX → PDF — local LibreOffice (`soffice --headless --convert-to pdf`) or PowerPoint COM
- `render-pdf.ps1` — PDF → PNG preview (optional)

## Layouts (v4)

| layout | purpose |
|---|---|
| `cover` / `section` | 封面 / 章節頁（深墨藍） |
| `statement` | 論點頁：action title + bullets |
| `statementSplit` | 論點 + 深色 insight panel（左文右卡） |
| `statementCards` | 卡片化：bullet 含「：/→」自動拆「主詞 + 說明」 |
| `statRow` | 統計列（number band）：大數字 + 標籤 + 注釋 |
| `table` | 機構對比表：hairline 分隔、狀態欄強調色 |
| `quote` | 引句頁（pull quote，深墨藍底） |
| `onePager` | 單頁總覽 |
| `closing` | 收尾 CTA |

## Rerun

```
npm install
node make-pptx.mjs spec-executive-capability.json
# output: Paul-Tradecraft-Executive-Capability-Deck.pptx
```

## Sample outputs

- `Paul-Tradecraft-Executive-Capability-Deck.pdf` — flagship Executive Capability deck (9 slides, v4), rendered from `spec-executive-capability.json`.
- `sample-output.pdf` — legacy AI-Collaboration One-Pager deck, rendered from `spec-ai-collab-onepager.json`.
