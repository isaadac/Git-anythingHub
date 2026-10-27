# Atlas Office

A free office suite for **Windows, macOS and Linux** (and any modern browser). It has three apps that work with Microsoft Office file formats:

| App | Like | Opens | Saves |
|---|---|---|---|
| **Atlas Write** | Word | `.docx`, `.html`, `.md`, `.txt` | `.docx`, `.html`, `.md`, `.txt`, PDF |
| **Atlas Sheets** | Excel | `.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv` | `.xlsx`, `.xls`, `.ods`, `.csv`, PDF |
| **Atlas Slides** | PowerPoint | `.pptx`, `.aslides` | `.pptx`, `.aslides`, PDF |

## Features

**Write**: ribbon UI, fonts and sizes, bold/italic/underline/strikethrough, sub/superscript, text and highlight colours, styles (Title, Headings, Quote, Code), bullets and numbering, alignment, line spacing, shading, format painter, tables (add/remove rows and columns), pictures, links, page breaks, text boxes, symbols, date and time, table of contents, header and footer, comments, find and replace, word count, read aloud, page setup (margins, orientation, Letter/A4/Legal, columns, page colour), navigation pane, print/web/focus/dark views, zoom.

**Sheets**: a grid that grows as you scroll, multiple sheets (add, rename, duplicate, reorder, delete), a formula bar with function autocomplete, **120+ Excel functions** (SUM, AVERAGE, IF/IFS/IFERROR, SUMIF(S)/COUNTIF(S)/AVERAGEIF(S), VLOOKUP/HLOOKUP/XLOOKUP/INDEX/MATCH, text, date, statistics, PMT/FV/PV/NPV, …), cross-sheet references, relative and absolute references that adjust on copy, fill and insert/delete, circular-reference detection, number formats (currency, percent, dates, …), fonts, colours, borders, alignment, wrap, merge and centre, AutoSum, fill down/right, sort, filter, remove duplicates, text to columns, column and line/pie/bar charts, freeze panes, a status-bar sum/average/count, and Excel keyboard shortcuts.

**Slides**: slide layouts, 6 themes, text boxes, 8 shape types, pictures, tables, drag/resize with snap guides, ordering and alignment, backgrounds (colour or picture), speaker notes, transitions, drag-to-reorder thumbnails, and a full-screen slide show.

All apps have undo/redo, unsaved-changes prompts, recent files, print and PDF export.

## Install

Download the installer for your system from the repository's **Releases** page (or from the latest **Actions → Atlas Office** run's artifacts):

- **Windows**: `Atlas Office Setup x.y.z.exe` (installer) or `Atlas Office-x.y.z-portable.exe` (no install needed)
- **macOS**: `Atlas Office-x.y.z-universal.dmg` (Intel and Apple Silicon)
- **Linux**: `Atlas Office-x.y.z.AppImage` (any distro: `chmod +x` and run) or `atlas-office_x.y.z_amd64.deb` (Debian/Ubuntu)

The builds are not code-signed yet. On Windows, SmartScreen may warn you ("More info → Run anyway"). On macOS, right-click the app and choose **Open** the first time.

**Phones, tablets and Chromebooks**: serve the `office-suite` folder from any static web host and open `src/renderer/index.html`. The same apps run in the browser. Saving downloads the file.

## Develop

```bash
cd office-suite
npm install
npm start          # run the desktop app
npm test           # formula engine tests
npm run dist       # build installers for the current OS into dist/
npm run dist:win   # or :mac / :linux
```

To publish a release with installers for all three systems, push a tag such as `office-v1.0.0`. `.github/workflows/office-suite.yml` builds on Windows, macOS and Linux runners and attaches the files to a GitHub Release.

## How it's built

- `src/main/`: Electron main process (windows, native menus, file dialogs, file associations, recent files) and a minimal, context-isolated preload bridge.
- `src/renderer/`: the three apps as plain HTML/CSS/JS with no build step. `js/platform.js` hides the desktop/web difference.
- `src/renderer/js/formula.js`: the spreadsheet engine (tokenizer, parser, evaluator, reference rewriting), with tests in `test/`.
- File formats: [docx](https://github.com/dolanmiu/docx) and [mammoth](https://github.com/mwilliamson/mammoth.js) for Word, [SheetJS](https://sheetjs.com) for spreadsheets, [PptxGenJS](https://github.com/gitbrent/PptxGenJS) and [JSZip](https://stuk.github.io/jszip/) for PowerPoint.

## Security and large files

- Documents are treated as untrusted. Opened and pasted content is sanitized (no scripts, frames, forms or remote content), and every page runs under a strict Content Security Policy.
- The renderer runs sandboxed with no Node.js access. It can only read or overwrite files the user picked in a dialog, opened from the OS or reopened from the recent list.
- Spreadsheet files are parsed in a throwaway Web Worker with a time limit, so a malicious file can't affect the app or freeze it. This contains the known SheetJS advisories (CVE-2023-30533 and CVE-2024-22363) in the npm build of `xlsx` 0.18.5. Upgrading to SheetJS 0.20.3 (`npm i https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`) removes them entirely.
- Sheets renders only the visible part of the grid, so workbooks with hundreds of thousands of rows open and scroll smoothly. Formula chains of any length (such as running totals down 100,000 rows) calculate without hitting the JavaScript stack limit.

## Known limitations

Atlas Office covers the everyday features of Word, Excel and PowerPoint, not all of them. Some things don't survive a round trip through Office files:

- **Write**: no live pagination (page count is estimated, and page breaks are explicit), and no tracked changes. Comments are kept as highlights only. Complex Word layouts (floating objects, section-specific headers) are simplified when opened.
- **Sheets**: cell styling and charts are saved inside the `.xlsx` in a form only Atlas reads back. Excel shows the values, formulas, number formats, column widths, row heights and merged cells, but not the colours or charts. Pivot tables, macros and conditional formatting are not supported.
- **Slides**: imported `.pptx` files keep text, pictures, shapes, tables, backgrounds and notes, but placeholder positions that come only from the slide master use default positions, and animations aren't supported. Save as `.aslides` to keep a presentation exactly as it is.
