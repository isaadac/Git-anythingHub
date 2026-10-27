# Atlas Office

A free office suite for **Windows, macOS and Linux** (and any modern browser). It has three apps that work with Microsoft Office file formats:

| App | Like | Opens | Saves |
|---|---|---|---|
| **Atlas Write** | Word | `.docx`, `.html`, `.md`, `.txt` | `.docx`, `.html`, `.md`, `.txt`, PDF |
| **Atlas Sheets** | Excel | `.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv` | `.xlsx`, `.xls`, `.ods`, `.csv`, PDF |
| **Atlas Slides** | PowerPoint | `.pptx`, `.aslides` | `.pptx`, `.aslides`, PDF |

## Features

**Write**: real pages like Word's Print Layout (text moves to the next page at the page edge; printing and PDF match the screen), headers and footers on every page with page numbers ("Page 1 of 3"), a different first page, **Track Changes** (accept/reject, All Markup / No Markup / Original) and **comments** with a comments pane, both saved as real Word revisions and comments in `.docx`, **footnotes and endnotes** (footnotes at the bottom of their page), **citations** in APA, MLA or Chicago style with a source manager and an automatic **bibliography**, all saved as real Word footnotes, endnotes and citation fields, **templates** (resume, cover letter, business letter, report, memo, meeting minutes, MLA essay, newsletter, invoice) from Home or File → New, **mail merge** from a CSV or spreadsheet (insert fields, preview each recipient, Finish & Merge into one document; fields are saved as Word MERGEFIELDs), ribbon UI, fonts and sizes, bold/italic/underline/strikethrough, sub/superscript, text and highlight colours, styles (Title, Headings, Quote, Code), bullets and numbering, alignment, line spacing, shading, format painter, tables (add/remove rows and columns), pictures, links, page breaks, text boxes, symbols, date and time, **table of contents** with page numbers and dot leaders (click an entry to jump to it; Update Table rebuilds it; saved as a real Word TOC field, and Word's own contents opens as a live one), **PDF export with clickable contents and bookmarks** for every heading, header and footer, comments, find and replace, word count, read aloud, page setup (margins, orientation, Letter/A4/Legal, columns, page colour), navigation pane, print/web/focus/dark views, zoom.

**Sheets**: a grid that grows as you scroll, **PivotTables** (rows, columns, values with Sum/Count/Average/Max/Min, report filters, subtotals and grand totals, a PivotTable Fields pane and Refresh All; other apps see the summary values), **conditional formatting** (highlight rules, text and duplicates, top/bottom and above/below average, formula rules, data bars, colour scales and icon sets, with a rules manager; saved as real Excel conditional formats and read from Excel files), multiple sheets (add, rename, duplicate, reorder, delete), a formula bar with function autocomplete, **120+ Excel functions** (SUM, AVERAGE, IF/IFS/IFERROR, SUMIF(S)/COUNTIF(S)/AVERAGEIF(S), VLOOKUP/HLOOKUP/XLOOKUP/INDEX/MATCH, text, date, statistics, PMT/FV/PV/NPV, …), cross-sheet references, relative and absolute references that adjust on copy, fill and insert/delete, circular-reference detection, number formats (currency, percent, dates, …), fonts, colours, borders, alignment, wrap, merge and centre, AutoSum, fill down/right, sort, filter, remove duplicates, text to columns, column and line/pie/bar charts, freeze panes, a status-bar sum/average/count, and Excel keyboard shortcuts.

**Slides**: slide layouts, 6 themes, text boxes, 8 shape types, pictures, tables, drag/resize with snap guides, ordering and alignment, backgrounds (colour or picture), speaker notes, transitions, drag-to-reorder thumbnails, and a full-screen slide show.

All apps have undo/redo, unsaved-changes prompts, recent files, print and PDF export, plus:

- **AutoSave**: the switch in the title bar saves to the file a few seconds after each change (desktop app; `.docx`/`.html`, `.xlsx`/`.ods`, `.pptx`/`.aslides`).
- **Crash recovery**: unsaved changes are kept as a draft every 30 seconds. If the app crashes or the computer loses power, Home lists the recovered documents next time. If a single window crashes, it offers to reopen with your changes.

## Install

Download the installer for your system from the repository's **Releases** page (or from the latest **Actions → Atlas Office** run's artifacts):

- **Windows**: `Atlas-Office-Setup-x.y.z-x64.exe` (installer; `-arm64` for ARM PCs) or `Atlas-Office-x.y.z-portable.exe` (no install needed)
- **macOS**: `Atlas-Office-x.y.z-universal.dmg` (Intel and Apple Silicon)
- **Linux**: `Atlas-Office-x.y.z-x86_64.AppImage` (any distro: `chmod +x` and run) or `atlas-office_x.y.z_amd64.deb` (Debian/Ubuntu)

Builds are code-signed once signing certificates are added to the repository (see [SIGNING.md](SIGNING.md)). Until then, Windows SmartScreen may warn you ("More info → Run anyway"), and on macOS you right-click the app and choose **Open** the first time.

**Phones, tablets and Chromebooks**: serve the `office-suite` folder from any static web host and open `src/renderer/index.html`. The same apps run in the browser. Saving downloads the file.

### Updates

Installed copies (Windows installer, macOS app, Linux AppImage and `.deb`) check the repository's latest GitHub Release every 6 hours, download new versions in the background and offer to restart. Use **Help → Check for Updates…** to check now, or untick **Help → Update Automatically** to turn this off. The portable `.exe` doesn't update itself. On macOS, updates need a signed build (see [SIGNING.md](SIGNING.md)).

To ship an update, raise `version` in `package.json` and push a tag such as `office-v1.0.1`.

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

- **Write**: a single table or block taller than a page isn't split across pages, and multi-column layouts flow without page breaks. Complex Word layouts (floating objects, section-specific headers) are simplified when opened.
- **Sheets**: cell styling and charts are saved inside the `.xlsx` in a form only Atlas reads back. Excel shows the values, formulas, number formats, column widths, row heights and merged cells, but not the colours or charts. Pivot tables, macros and conditional formatting are not supported.
- **Slides**: imported `.pptx` files keep text, pictures, shapes, tables, backgrounds and notes, but placeholder positions that come only from the slide master use default positions, and animations aren't supported. Save as `.aslides` to keep a presentation exactly as it is.
