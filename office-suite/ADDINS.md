# Isadek Office add-ins

An add-in adds buttons (on the **Insert** tab, in the **Add-ins** group) and a pane to Isadek
Write, Sheets or Slides. Install one with **Insert → Add-ins → Install from File…**; the same
button turns add-ins on or off and removes them. Two samples are in [`addins/`](addins/).

## The file

An add-in is one JSON file, usually named `something.isadek-addin`:

```json
{
  "id": "com.example.word-goal",
  "name": "Word Goal",
  "version": "1.0.0",
  "author": "Example Ltd",
  "description": "Shows how far the document is toward a number of words.",
  "apps": ["write"],
  "permissions": ["read"],
  "commands": [{ "id": "show", "label": "Word Goal", "icon": "target" }],
  "script": "isadek.onCommand('show', async () => { … })"
}
```

- `id`: letters, digits, dots and dashes. Installing a file with the same id updates it.
- `apps`: any of `write`, `sheets`, `slides`. A command can list its own `apps`.
- `permissions`: `read` (read the document) and `write` (change it). Calls outside these fail.
- `commands`: up to 8 buttons. `icon` is a [Lucide](https://lucide.dev) name from the set the app
  ships (`puzzle`, `target`, `calculator`, `wand-2`, `list-checks`, `chart-bar`, `languages`,
  `timer`, `sparkles`, `hash`, `book-open`, `clipboard-list`, `calendar-days`, `quote`, `smile`);
  others show a puzzle piece.
- `script`: the add-in's JavaScript (up to 1 MB).

## Where it runs

Each add-in runs in its own Web Worker with no network (`fetch`, `XMLHttpRequest`, `WebSocket`…),
no storage, no page and no other add-ins. It reaches the document only through `isadek`, and
every call is checked against its permissions. A command that runs for more than a minute is
stopped. What an add-in shows is built from the blocks below as plain text and controls, never
as HTML.

## The `isadek` API

All calls return promises.

| Call | Apps | Permission | Result |
| --- | --- | --- | --- |
| `isadek.onCommand(id, fn)` | all | — | `fn()` runs when the button with that id is clicked |
| `isadek.onAction(id, fn)` | all | — | `fn(values)` runs when a pane button is clicked; `values` holds the pane's inputs by id |
| `isadek.document.getText()` | all | read | the document's text (Sheets: cells as `A1<tab>value` lines) |
| `isadek.document.getSelection()` | all | read | the selected text (Sheets: the selected address) |
| `isadek.document.insertText(text)` | write | write | types the text at the cursor |
| `isadek.document.getSelectedRange()` | sheets | read | `{ address, sheet, values }` |
| `isadek.document.getValues("A1:B3")` | sheets | read | rows of values (numbers, text, booleans, error codes) |
| `isadek.document.setValues("C1", [[1, 2], ["=A1*2", "x"]])` | sheets | write | writes from that cell; text starting with `=` is a formula; one undo step |
| `isadek.document.getSlides()` | slides | read | `[{ index, title, text, notes }]` |
| `isadek.document.addSlide(title, bullets)` | slides | write | adds a slide after the current one |
| `isadek.ui.toast(text)` | all | — | a short message |
| `isadek.ui.showPane(title, blocks)` | all | — | shows the add-in's pane |
| `isadek.ui.closePane()` | all | — | |
| `isadek.ui.form(title, fields, okLabel)` | all | — | a dialog; resolves to the values, or `null` if cancelled |

Pane blocks: `{ type: "heading" | "text", text }`, `{ type: "list", items }`,
`{ type: "progress", value: 0…1, label }`, `{ type: "input", id, label, value, inputType: "text" | "number" }`,
`{ type: "button", id, label, primary }`.

Form fields: `{ name, label, type: "text" | "number" | "textarea" | "select", value, options }`.

`isadek.app` is the app it runs in; `isadek.addin` is `{ id, name, version }`.
