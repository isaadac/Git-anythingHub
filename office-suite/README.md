# Atlas Office

A free office suite for **Windows, macOS and Linux** (and any modern browser). It has three apps that work with Microsoft Office file formats:

| App | Like | Opens | Saves |
|---|---|---|---|
| **Atlas Write** | Word | `.docx`, `.html`, `.md`, `.txt` | `.docx`, `.html`, `.md`, `.txt`, PDF |
| **Atlas Sheets** | Excel | `.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv` | `.xlsx`, `.xls`, `.ods`, `.csv`, PDF |
| **Atlas Slides** | PowerPoint | `.pptx`, `.aslides` | `.pptx`, `.aslides`, PDF |

## Features

**Write**: real pages like Word's Print Layout (text moves to the next page at the page edge; printing and PDF match the screen), headers and footers on every page with page numbers ("Page 1 of 3"), a different first page, **Track Changes** (accept/reject, All Markup / No Markup / Original) and **comments** with a comments pane, both saved as real Word revisions and comments in `.docx`, **footnotes and endnotes** (footnotes at the bottom of their page), **citations** in APA, MLA or Chicago style with a source manager and an automatic **bibliography**, all saved as real Word footnotes, endnotes and citation fields, **templates** (resume, cover letter, business letter, report, memo, meeting minutes, MLA essay, newsletter, invoice) from Home or File → New, **mail merge** from a CSV or spreadsheet (insert fields, preview each recipient, Finish & Merge into one document; fields are saved as Word MERGEFIELDs), ribbon UI, fonts and sizes, bold/italic/underline/strikethrough, sub/superscript, text and highlight colours, styles (Title, Headings, Quote, Code), bullets and numbering, alignment, line spacing, shading, format painter, tables (add/remove rows and columns), pictures, links, page breaks, text boxes, symbols, date and time, **table of contents** with page numbers and dot leaders (click an entry to jump to it; Update Table rebuilds it; saved as a real Word TOC field, and Word's own contents opens as a live one), **PDF export with clickable contents and bookmarks** for every heading, header and footer, comments, find and replace, word count, read aloud, page setup (margins, orientation, Letter/A4/Legal, columns, page colour), navigation pane, print/web/focus/dark views, zoom.

**Sheets**: a grid that grows as you scroll, **notes** on cells (red corner, hover, Show All, Shift+F2) and **protected sheets** (lock/unlock cells, optional password, allowed actions; Excel's protection and notes both ways), **charts** (column, bar, line, area, pie, doughnut, scatter and combo; saved as real Excel charts that Excel, LibreOffice and Google Sheets can edit, and Excel charts open as Atlas charts), **fill handle** (drag in any direction to continue numbers, dates, "Item 1"-style text, day and month names, or copy with Ctrl; double-click to fill down beside data; Fill → Series for linear, growth and date series), **PivotTables** (rows, columns, values with Sum/Count/Average/Max/Min, report filters, subtotals and grand totals, a PivotTable Fields pane and Refresh All; other apps see the summary values), **data validation** (lists with in-cell dropdowns, including lists from a range or another sheet; whole/decimal numbers, dates, times, text length and custom formulas; input messages; Stop/Warning/Information alerts; Circle Invalid Data; read from and saved as Excel validation), **conditional formatting** (highlight rules, text and duplicates, top/bottom and above/below average, formula rules, data bars, colour scales and icon sets, with a rules manager; saved as real Excel conditional formats and read from Excel files), multiple sheets (add, rename, duplicate, reorder, delete), a formula bar with function autocomplete, **270+ Excel functions** including **dynamic arrays** that spill like Excel 365 (FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, XMATCH, TEXTSPLIT, VSTACK/HSTACK, TAKE/DROP, LET, A1# spill references, array maths and #SPILL!), INDIRECT/OFFSET, statistics (CORREL, PERCENTILE, NORM.DIST…), finance (IRR, XIRR, RATE, NPER, IPMT/PPMT, DDB…), and newer functions saved with the names Excel expects (SUM, AVERAGE, IF/IFS/IFERROR, SUMIF(S)/COUNTIF(S)/AVERAGEIF(S), VLOOKUP/HLOOKUP/XLOOKUP/INDEX/MATCH, text, date, statistics, PMT/FV/PV/NPV, …), cross-sheet references, relative and absolute references that adjust on copy, fill and insert/delete, circular-reference detection, number formats (currency, percent, dates, …), fonts, colours, borders, alignment, wrap, merge and centre, AutoSum, fill down/right, sort, filter, remove duplicates, text to columns, column and line/pie/bar charts, freeze panes, a status-bar sum/average/count, and Excel keyboard shortcuts.

**Slides**: **rotate, group and connect** (rotation handle with 15° snapping, rotate 90° and flip; Shift+click multi-select, Group/Ungroup with scaling; straight and elbow arrow connectors that stay attached when shapes move; saved as real PowerPoint connectors, and PowerPoint groups and connectors open as such), **video and audio** (MP4/WebM/MOV and MP3/WAV/M4A/OGG; poster frame, start automatically or on click, loop, hide the audio icon; embedded in .pptx and read from PowerPoint files), **Slide Master view** (master shapes and background on every slide, editable built-in layouts and custom layouts with placeholders, Reset slide, Header & Footer with slide numbers and date, hide background graphics per slide), much better PowerPoint import (placeholders inherit their position and text size from layouts and the master, group shapes keep their scaling, the deck's masters, layouts, backgrounds, theme colours and fonts come across), **Presenter View** (current slide, next slide or build, notes with zoom, timer and clock on your screen; the slides full screen on the second display, or in a second window), **animations** (entrance, emphasis and exit effects with direction, On Click / With / After Previous, duration, delay and order; builds play in the slide show and are saved as PowerPoint animations), **transitions** (fade, push, wipe, split, cover, uncover, zoom with duration and auto-advance, saved in .pptx), slide layouts, 6 themes, text boxes, 8 shape types, pictures, tables, drag/resize with snap guides, ordering and alignment, backgrounds (colour or picture), speaker notes, transitions, drag-to-reorder thumbnails, and a full-screen slide show.

All apps have undo/redo, unsaved-changes prompts, recent files, print and PDF export, plus:

- **Cloud storage and sharing**: connect **Google Drive, OneDrive, Dropbox** or a **WebDAV** server (Nextcloud, ownCloud) on Home, then File → Open from Cloud / Save to Cloud. Ctrl+S and AutoSave write straight back to the cloud file, and **Share** (title bar) makes a view or edit link or invites people by email. Google Docs, Sheets and Slides files open as their Office equivalents; saving creates an Office copy beside them. Sign-in uses OAuth with PKCE in your normal browser (desktop) or a pop-up (web); tokens are kept encrypted with the OS keychain on the desktop (in the web version, in the browser's local storage). Each service needs an app registration of your own: put its client ID in Home → Cloud → Settings, with the redirect address shown there.
- **Spelling, grammar and languages**: red underlines as you type, with suggestions and *Add to Dictionary* on right-click (desktop). **Review → Spelling & Grammar** (F7) opens the Editor pane in Write, Sheets (text cells) and Slides (text boxes and speaker notes): it walks through misspellings and grammar slips — repeated words, extra spaces, spacing around punctuation, sentence capitals, "a"/"an", "i", could of, then/than, very long sentences — with Change, Change All, Ignore Once, Ignore All and Add to Dictionary. **Review → Language** (or the language in Write's status bar) sets the proofing language of a document from 50+ languages; Write saves it with the document. English grammar rules apply to English text; the others to every language.
- **Assistant (Claude)**: **✨ Assistant** in the title bar opens a pane to ask Claude about the open document, with one-click actions. Write: summarize, improve writing, fix spelling and grammar, shorten, make more formal, continue writing, with *Replace selection* and *Insert* (Markdown replies become headings, lists and tables). Sheets: write or explain a formula, analyze the selection, find mistakes; *Put in sheet* writes a formula into the active cell or pastes a table. Slides: draft slides from a topic, speaker notes, tighten a text box; *Add as slides* turns an outline into slides with bullets and notes. Uses your own Anthropic API key (⚙ in the pane), kept encrypted by the desktop app and never given to the pages; choose Claude Opus 5.5, Sonnet 5.5 or Haiku 5.5. Desktop app only.
- **Live editing**: click **Live** in the title bar to edit a document together in real time. Start a session on this computer (people on the same network join with the link) or on your own server (`npm run live-server`, or `node src/main/collab-server.js --port 8787` behind any TLS proxy for `wss://`), then send the invite link. Everyone sees the others' changes as they type, their cursors and selections with their names, and who is connected; edits made while the connection drops are sent when it comes back, and Undo takes back only your own changes. Write syncs paragraph by paragraph, Sheets cell by cell, Slides shape by shape; when two people change the same paragraph, cell or shape at the same moment, the later change wins.
- **AutoSave**: the switch in the title bar saves to the file a few seconds after each change (desktop app or cloud documents; `.docx`/`.html`, `.xlsx`/`.ods`, `.pptx`/`.aslides`).
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
- **Sheets**: cell colours and fonts are saved inside the `.xlsx` in a form only Atlas reads back; Excel shows the values, formulas, number formats, charts, conditional formats, column widths, row heights and merged cells. Other apps see a pivot table's values, not a live pivot. Macros are not supported.
- **Slides**: SmartArt, 3D models, Morph and other effects Atlas doesn't have are dropped when a `.pptx` opens. Save as `.aslides` to keep a presentation exactly as it is.
- **Live editing**: in Write, comments, footnotes and citation sources aren't shared live yet, and Undo during a session uses the browser's own history, which can also take back other people's edits to the same paragraph. The relay server keeps sessions in memory only; anyone with the link can join, so share it only with people you trust, and use `wss://` across the internet.
- **Cloud**: the web version can only reach WebDAV servers that allow cross-origin requests (Nextcloud needs a CORS app); the desktop app has no such limit. Files over 250 MB can't be saved to OneDrive.
