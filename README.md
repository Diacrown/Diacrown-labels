# Diacrown Label Sheet Generator

A small in-house tool that turns an ERP (EasyGems) invoice export directly
into a print-ready Word label sheet — no Excel cleanup, no Word mail merge,
no manual formatting.

Open the page, drop in the `.xlsx` export, download the finished `.docx`.
Everything runs **client-side in the browser** — no file is ever uploaded
to a server, which makes this safe to host as a public static page even
though the data it processes is internal.

## How it works

- Reads the uploaded Excel file in the browser using [SheetJS](https://sheetjs.com/)
- Matches columns by name (case-insensitive, tolerant of minor naming differences)
- Recreates the exact label layout used in the original `STICKER_NEW.docx`
  template: 5 labels across × 8 down (40 per page), sized for 39×35&nbsp;mm
  laser label stock on A4
- Builds the `.docx` file's XML directly and zips it with [JSZip](https://stuk.github.io/jszip/)
- Triggers a normal browser download of the finished Word file

No backend, no build step, no dependencies to install. It's a single
`index.html` file.

## Excel column requirements

Any sheet with these columns works (raw ERP names or the "cleaned up"
friendly names both match — matching is case-insensitive):

| Field | Accepted column names |
|---|---|
| SSP | `SSP`, `FSSP`, `SSP Any`, `Lot #` |
| Item | `Item#`, `Item`, `Item Number` |
| Stone Type | `StnTyp`, `Stone Type`, `Gem` |
| Shape | `Shp`, `Shape` |
| Color | `Col`, `Color`, `Colour` |
| Weight | `Units`, `Wt`, `Weight`, `Carat(s)`, `Cts` |
| Lab + Report# | `Lab` + `Report#` (separately), or one combined `Lab Report Number` column |
| Dimensions | `L`, `W`, `D` (or `Length`, `Width`, `Depth`) |

Extra columns in the sheet are ignored — no need to delete anything first.
Every data row becomes one label.

## Deploying on GitHub Pages

1. Create a new repository (e.g. `diacrown-labels`) and add `index.html`
   to the root.
2. Push it to GitHub.
3. In the repo: **Settings → Pages → Build and deployment → Source:**
   `Deploy from a branch`, branch `main`, folder `/ (root)`.
4. GitHub will publish it at `https://<your-username>.github.io/diacrown-labels/`
   (takes a minute or two on first deploy).
5. Share that link with your team — that's the whole tool.

The page needs an internet connection the first time it loads, just to
fetch the two small libraries (SheetJS and JSZip) from cdnjs. Everything
after that — reading the file, building the document — happens locally
in the browser.

## Files in this repo

- `index.html` — the tool itself. This is the only file GitHub Pages needs.
- `scripts/generate_labels.py` — optional command-line version of the same
  logic, useful for bulk/automated runs outside the browser. Not required
  for the web page to work.

## Updating the label layout

All of the layout logic (column widths, fonts, spacing, page size) lives
in the `<script>` block near the bottom of `index.html`, inside the
constants at the top (`COL_WIDTH_TW`, `ROW_HEIGHT_TW`, etc.) and the
`buildCellXml` / `buildTableXml` functions. If your label stock or layout
ever changes, that's the section to edit.
