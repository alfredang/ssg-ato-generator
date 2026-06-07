# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Streamlit app that generates a customised **Standard Operating Procedure (SOP) Word document** for SkillsFuture Singapore (SSG) Registered Training Partners. The user supplies company details (manually or by uploading supporting documents for auto-extraction); the app fills a **placeholder-only template**, regenerates branded diagrams, and outputs `SOP_<Company Name>.docx`. It also includes an SSG RTP registration guide and a supporting-documents checklist.

## Commands

```bash
# Run the app (Graphviz `dot` must be on PATH — it lives in /opt/homebrew/bin)
export PATH="/opt/homebrew/bin:$PATH"
uv run streamlit run app.py --server.port 8501 --server.headless true

# Rebuild the SOP template from scratch (after editing build_template.py)
uv run python build_template.py        # -> writes sop_template.docx

# Regenerate all diagrams (org charts + flowcharts) into a folder
uv run python flowcharts.py /tmp/diagrams

# Regenerate the static RTP application-flow image used by the guide page
uv run python -c "import flowcharts; flowcharts.rtp_application('assets')"

# Sync dependencies
uv sync
```

There is **no test suite, linter, or build step**. To verify a change, render a generated `.docx` to PDF with LibreOffice and inspect it visually:

```bash
/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to pdf --outdir /tmp/out SOP_<name>.docx
```

## System dependencies (not pip-installable)

- **Graphviz `dot`** — required to render diagrams. Installed via Homebrew (`/opt/homebrew/bin/dot`). `flowcharts.py` prepends `/opt/homebrew/bin` to `PATH` at import, but a shell launching the app should also export it.
- **LibreOffice** (`soffice`) — only used for visual verification/rendering, not at runtime.
- `pdftotext` (poppler) — used by `extraction.py` as the primary PDF text extractor, with `pypdf` as fallback.

## Architecture

The generation pipeline is **template + markers + per-company diagram regeneration**. The shipped template contains *no* real company or personal data — only tokens.

- **`build_template.py`** — authors `sop_template.docx` from scratch with `python-docx`. Produces: cover page, version-control table, an auto-updating Word TOC field (`updateFields` is enabled so Word builds it on open), all SOP sections (A–D), and a company-name footer. Sections embed two kinds of markers:
  - text placeholders: `{{COMPANY_NAME}}`, `{{OWNER_NAME}}`, `{{UEN}}`, `{{ADDRESS}}`, `{{WEBSITE}}`, `{{CONTACT_NAME}}`, `{{CONTACT_TEL}}`, `{{CONTACT_EMAIL}}`, `{{DATE}}`, and `{{LOGO}}`.
  - diagram markers: paragraphs containing `[[DIAGRAM:<key>]]`.

- **`flowcharts.py`** — builds every diagram as a Graphviz PNG, **parameterised by `owner` name** so the template stays generic and the owner is injected only at generation time. `DIAGRAMS` maps each `[[DIAGRAM:key]]` to its builder (`org_management`, `org_training`, `enquiry`, `funding`, `course_confirmation`, `refund`, `trainer_performance`). `build_all(out_dir, owner)` renders them all. `rtp_application()` is a separate static guide diagram (not in `DIAGRAMS`). Flowcharts were recreated from the original InnoHat and HighSpark SOPs; shared style helpers (`_proc`, `_decision`, `_terminal`) keep them consistent.

- **`sop_generator.py`** — the fill engine used by the app. `generate(info, logo_bytes, show_logo)`: loads the template, calls `flowcharts.build_all` with the owner, replaces `{{...}}` tokens across paragraphs/tables/headers/footers, inserts the logo at `{{LOGO}}` (or removes it when toggled off), and swaps each `[[DIAGRAM:key]]` marker for its regenerated PNG (sized via `_fit_box` to avoid overflow). Returns a `BytesIO`. `safe_filename()` derives `SOP_<Company>.docx`.

- **`extraction.py`** — best-effort parsing of uploaded ACRA/BizFile business profiles (PDF or DOCX) to pre-fill the form (company name, UEN, registered address, email, tel). Scanned/image PDFs yield no text — it returns notes telling the user to enter details manually.

- **`app.py`** — Streamlit UI with a sidebar of three pages: **Generate SOP** (manual entry or upload-and-extract, logo toggle, download), **Documents Checklist** (interactive 5-set checklist + Training Record `.xlsx` download from `assets/`), **SSG RTP Registration Guide** (visual summary + application-flow image). A fixed footer credits Tertiary Infotech Academy.

- **`assets/`** — `Training_Record_Template.xlsx` (offered for download) and `rtp_application.png` (the static guide diagram).

## Key constraints when editing

- **The template must never contain real company/personal data** — keep everything as `{{placeholders}}` or `[[DIAGRAM:markers]]`. Owner/company names enter diagrams only via `flowcharts.build_all(owner=...)` at generation time.
- If you add a new diagram: add its builder to `flowcharts.DIAGRAMS`, then add a matching `[[DIAGRAM:key]]` marker in `build_template.py` and re-run it.
- If you add a new placeholder: add the token in `build_template.py` and the key to `PLACEHOLDER_KEYS` in `sop_generator.py`.
- The TOC is a Word field; it renders empty until Word/LibreOffice updates fields on open (this is expected, surfaced to the user in the UI).
