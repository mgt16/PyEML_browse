# EML Archive Explorer

A small desktop app for browsing `.eml` files on disk. It scans the **current working directory** recursively, lists messages in a table, and shows the selected message body with optional attachment handling.

### Outlook for Mac (`.olm`) → `.eml`

This explorer only reads **`.eml`** files. If you exported mail from **Outlook for Mac** as an **`.olm`** archive, convert it to `.eml` first. A free, open-source option is **[OLM Convert](https://github.com/PeterWarrington/olm-convert)** ([web app](https://www.lilpete.me/olm-convert)): it turns `.olm` into a folder of `.html` or `.eml` messages. Use **`eml`** output, then point this app at the directory that contains those files (or copy them under your project tree) and run as usual.

## Requirements

- **Python 3** with the standard library (uses `tkinter` for the UI; on Linux you may need the `python3-tk` package).

## Run

From this project directory (so the bundled sample mail is included in the scan):

```bash
python3 PythonEmail_browse.py
```

## What it does

### Message list

- **Folder** — Relative path to the folder containing the `.eml` (subfolders included), with `/` separators; files in the launch directory show as `/`.
- **Date** — Parsed `Date` header, shown as `YYYY-MM-DD HH:MM`, or `Unknown` if missing or invalid.
- **From** — Address from the `From` header (parsed).
- **Subject** — Subject line, or `(No Subject)` if absent.
- **Attachments** — `Yes` or `No`, based on whether the message has non-multipart parts with a filename or `Content-Disposition: attachment`.

### Sorting

Click any column header to sort by that column. The active column shows **▲** (ascending) or **▼** (descending). Click again on the same column to reverse order. Date sorting uses the real message date, not the displayed string.

### Filters

Each column has a filter field. Typing text **narrows the list** to rows where that column’s value contains the text (case-insensitive substring). Empty fields are ignored. Filters combine with **AND** logic (all non-empty filters must match).

### Search

The **Search (subject, sender, body, attachment names)** field matches a **case-insensitive substring** across the message **subject**, **From** header (display name and address), **body** (same plain/HTML handling as the preview), and **filenames** of parts treated as attachments (same parts as in the preview’s attachment list). It applies together with the column filters (**AND**): every non-empty filter and the search term (if any) must match.

### Preview and attachments

Selecting a row loads that `.eml` file. If a **`text/html`** part exists, it is shown **rendered** in the preview (headings, bold/italic, lists, block quotes, links you can click, preformatted text) similar to a typical mail client—not as raw HTML. Images, embedded media, and iframes are not inlined from the web, so that markup is **stripped** before rendering to avoid blank space where those assets would appear. If there is only **`text/plain`**, the body is shown as normalized plain text. Search/indexing still uses a plain-text form of the message (cleaned HTML plus any plain alternative). If there are attachments, names are listed below the preview; **Save All** lets you pick a folder and write attachment payloads to disk.

### Scan behavior

Loading runs in a **background thread** so the window stays responsive. Every `.eml` file under the process current directory is considered.

## Sample data

The repo includes **`ExampleEmails/Inbox/`** and **`ExampleEmails/Sent/`** with a few test messages (with and without attachments) for local testing.

## Tests

From the project root:

```bash
python3 -m unittest tests.test_email_browse -v
```

Tests cover folder path display, loading `.eml` metadata, attachment detection, HTML stripping, body extraction, column filters, global search, and sorting (no GUI required).
