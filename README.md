# Spec-to-Deck Pipeline

Rerunnable spec-JSON-to-PPTX-to-PDF deck pipeline (pptxgenjs). Part of the **DSH Guide Deck** portfolio work.

## How it works

- spec-*.json — declarative deck spec (slides, sections, copy)
- make-pptx.mjs <spec.json> — renders the spec to a .pptx
- ender-pdf.ps1 — optional PPTX-to-PDF export (local LibreOffice/Office)

## Rerun

npm install
node make-pptx.mjs spec-ai-collab-onepager.json
# output: Paul-Tradecraft-AI-Collaboration-OnePager.pptx

## Sample output

sample-output.pdf — the AI-Collaboration One-Pager deck rendered from spec-ai-collab-onepager.json.
