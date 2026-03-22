"""
Browse folders of .eml files. Outlook for Mac exports use .olm archives, not .eml;
convert with OLM Convert (https://github.com/PeterWarrington/olm-convert) to .eml,
then open this app from the folder that contains the converted messages.
"""
import html
import os
import re
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import queue
from html.parser import HTMLParser
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime

COLUMNS = ("Folder", "Date", "From", "Subject", "Attachments")

_FIELD_KEYS = {
    "Folder": "folder",
    "Date": "date",
    "From": "from",
    "Subject": "subject",
    "Attachments": "attach_label",
}


def field_key_for(col):
    return _FIELD_KEYS[col]


def folder_display_for_path(path, root_path):
    rel_path = os.path.relpath(path, root_path)
    rel_dir = os.path.dirname(rel_path)
    if not rel_dir or rel_dir == ".":
        return "/"
    return rel_dir.replace(os.sep, "/")


def sort_key_for_row(d, col):
    if col == "Date":
        return (d.get("sort_key") or 0,)
    if col == "Attachments":
        return (0 if d.get("has_attachments") else 1,)
    s = str(d.get(field_key_for(col), "")).lower()
    return (s,)


def filter_emails(emails, filters, global_search=""):
    """filters: mapping column name -> filter string (substring, case-insensitive).
    global_search: if non-empty, row must contain this substring in search_blob."""
    gs = (global_search or "").strip().lower()
    out = []
    for d in emails:
        ok = True
        for col in COLUMNS:
            needle = (filters.get(col) or "").strip().lower()
            if not needle:
                continue
            hay = str(d.get(field_key_for(col), "")).lower()
            if needle not in hay:
                ok = False
                break
        if ok and gs and gs not in d.get("search_blob", ""):
            ok = False
        if ok:
            out.append(d)
    return out


def filtered_sorted_emails(emails, sort_column, sort_reverse, filters, global_search=""):
    rows = filter_emails(emails, filters, global_search=global_search)
    rows.sort(key=lambda x: sort_key_for_row(x, sort_column), reverse=sort_reverse)
    return rows


class HTMLToPlainText(HTMLParser):
    """Turn HTML email into readable plain text (block breaks, no scripts, normalized space)."""

    _BLOCK = frozenset(
        {
            "p",
            "div",
            "section",
            "article",
            "header",
            "footer",
            "main",
            "tr",
            "li",
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
            "blockquote",
            "pre",
            "table",
            "form",
        }
    )

    def __init__(self):
        super().__init__()
        self.reset()
        self.fed = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        if t in ("script", "style", "noscript", "template"):
            self._skip += 1
            return
        if self._skip:
            return
        if t == "br":
            self.fed.append("\n")
        elif t == "hr":
            self.fed.append("\n---\n")
        elif t in self._BLOCK:
            if self.fed and not self.fed[-1].endswith("\n"):
                self.fed.append("\n")
        elif t in ("td", "th"):
            self.fed.append(" ")

    def handle_endtag(self, tag):
        t = tag.lower()
        if t in ("script", "style", "noscript", "template"):
            self._skip = max(0, self._skip - 1)
            return
        if self._skip:
            return
        if t in self._BLOCK or t in ("p", "table", "tbody", "thead", "tfoot"):
            self.fed.append("\n")

    def handle_data(self, d):
        if self._skip:
            return
        self.fed.append(d)

    def get_data(self):
        return "".join(self.fed)


def normalize_plain_body(text):
    """Collapse noisy whitespace so plain and HTML-derived text read like a normal message."""
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = []
    for line in text.split("\n"):
        line = re.sub(r"[\t\f\v ]+", " ", line).strip()
        lines.append(line)
    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def clean_html(html_content):
    s = HTMLToPlainText()
    try:
        s.feed(html_content)
        s.close()
        raw = html.unescape(s.get_data())
        return normalize_plain_body(raw)
    except Exception:
        return normalize_plain_body(html_content)


def configure_mail_preview_tags(text_widget):
    """Tk Text tags that approximate a typical mail client (Outlook/Apple Mail–like)."""
    base = ("Segoe UI", 11)
    text_widget.tag_configure("mail_base", font=base, foreground="#1a1a1a")
    text_widget.tag_configure("mail_b", font=(base[0], base[1], "bold"))
    text_widget.tag_configure("mail_i", font=(base[0], base[1], "italic"))
    text_widget.tag_configure("mail_bi", font=(base[0], base[1], "bold italic"))
    text_widget.tag_configure("mail_u", underline=True)
    text_widget.tag_configure("mail_h1", font=(base[0], 20, "bold"), spacing1=0, spacing3=0)
    text_widget.tag_configure("mail_h2", font=(base[0], 16, "bold"), spacing1=0, spacing3=0)
    text_widget.tag_configure("mail_h3", font=(base[0], 14, "bold"), spacing1=0, spacing3=0)
    text_widget.tag_configure("mail_h4", font=(base[0], 12, "bold"), spacing1=0, spacing3=0)
    text_widget.tag_configure("mail_small", font=(base[0], 9))
    text_widget.tag_configure("mail_pre", font=("Consolas", 10))
    text_widget.tag_configure("mail_quote", font=base, foreground="#4a4a4a", lmargin1=12, lmargin2=12)
    text_widget.tag_configure("mail_link", font=base, foreground="#0b57d0", underline=True)


def _mail_pick_style_tag(stack):
    """Map semantic stack to one Tk tag name (simple precedence)."""
    if "pre" in stack:
        return "mail_pre"
    if "h1" in stack:
        return "mail_h1"
    if "h2" in stack:
        return "mail_h2"
    if "h3" in stack:
        return "mail_h3"
    if "h4" in stack:
        return "mail_h4"
    if "h5" in stack:
        return "mail_h3"
    if "h6" in stack:
        return "mail_h4"
    if "small" in stack:
        return "mail_small"
    b = "b" in stack
    i = "i" in stack
    if b and i:
        return "mail_bi"
    if b:
        return "mail_b"
    if i:
        return "mail_i"
    return "mail_base"


_SKIP_HTML_SUBTREE = frozenset(
    {
        "script",
        "style",
        "noscript",
        "template",
        "head",
        "iframe",
        "object",
        "embed",
        "video",
        "audio",
        "canvas",
        "svg",
        "picture",
        "map",
    }
)


def strip_html_asset_placeholders(html_src):
    """Remove tags for images/media/objects not present as body parts in .eml (avoids empty gaps in preview)."""
    if not html_src:
        return html_src
    s = html_src
    for tag in ("iframe", "object", "embed", "video", "audio", "canvas", "svg", "picture", "map"):
        s = re.sub(rf"<{tag}\b[^>]*>.*?</{tag}>", "", s, flags=re.I | re.S)
    for tag in ("img", "input", "source", "track", "area", "base"):
        s = re.sub(rf"<{tag}\b[^/>]*/?>", "", s, flags=re.I)
    s = re.sub(r"<img\b[^>]*>", "", s, flags=re.I)
    return s


def remove_empty_html_containers(html_src):
    """Drop empty wrappers left after asset removal (common source of extra blank lines)."""
    if not html_src:
        return html_src
    s = html_src
    for _ in range(10):
        prev = s
        for tag in ("div", "p", "span", "td", "th", "section", "article", "center", "footer", "header", "li", "main", "nav"):
            s = re.sub(rf"<{tag}\b[^>]*>\s*</{tag}>", "", s, flags=re.I)
        s = re.sub(
            r"<div\b[^>]*>\s*(?:<br\s*/?>|\s)*\s*</div>",
            "",
            s,
            flags=re.I,
        )
        s = re.sub(
            r"<p\b[^>]*>\s*(?:<br\s*/?>|\s)*\s*</p>",
            "",
            s,
            flags=re.I,
        )
        if s == prev:
            break
    return s


def prepare_html_for_mail_preview(html_src):
    """Strip external-asset markup and empty shells before feeding the renderer."""
    s = strip_html_asset_placeholders(html_src)
    return remove_empty_html_containers(s)


class HTMLMailRenderer(HTMLParser):
    """Render HTML email body into a Tk Text widget with client-like typography."""

    _BLOCK = frozenset(
        {
            "p",
            "div",
            "section",
            "article",
            "header",
            "footer",
            "main",
            "center",
            "table",
            "tbody",
            "thead",
            "tfoot",
            "tr",
            "form",
        }
    )

    def __init__(self, text_widget):
        super().__init__()
        self.tw = text_widget
        self._skip = 0
        self._stack = []
        self._link_stack = []
        self._link_tag_seq = 0
        self._last_was_block_break = True
        self._trailing_newline_run = 0

    def _semantic_push(self, name):
        self._stack.append(name)

    def _semantic_pop(self, name):
        if name in ("b", "strong"):
            key = "b"
        elif name in ("i", "em"):
            key = "i"
        else:
            key = name
        for i in range(len(self._stack) - 1, -1, -1):
            if self._stack[i] == key:
                self._stack = self._stack[:i]
                return

    def _insert(self, text):
        if not text:
            return
        if "\n" in text:
            text = re.sub(r"\n{3,}", "\n\n", text)
        if text == "\n":
            if self._trailing_newline_run >= 2:
                return
            st = _mail_pick_style_tag(self._stack)
            tags = [st]
            if "u" in self._stack:
                tags.append("mail_u")
            if "blockquote" in self._stack:
                tags.append("mail_quote")
            if self._link_stack:
                href, ltag = self._link_stack[-1]
                if ltag:
                    tags.append(ltag)
                elif href:
                    tags.append("mail_link")
            self.tw.insert(tk.END, "\n", tuple(tags))
            self._trailing_newline_run += 1
            return
        st = _mail_pick_style_tag(self._stack)
        tags = [st]
        if "u" in self._stack:
            tags.append("mail_u")
        if "blockquote" in self._stack:
            tags.append("mail_quote")
        if self._link_stack:
            href, ltag = self._link_stack[-1]
            if ltag:
                tags.append(ltag)
            elif href:
                tags.append("mail_link")
        t = tuple(tags)
        self.tw.insert(tk.END, text, t)
        tail = len(text) - len(text.rstrip("\n"))
        self._trailing_newline_run = min(2, tail) if tail else 0

    def _soft_break(self):
        if not self._last_was_block_break:
            self._insert("\n")
            self._last_was_block_break = True

    def _block_break(self):
        self._insert("\n")
        self._last_was_block_break = True

    def handle_starttag(self, tag, attrs):
        ad = {k.lower(): v for k, v in attrs}
        t = tag.lower()
        if t in _SKIP_HTML_SUBTREE:
            self._skip += 1
            return
        if t in ("img", "input", "source", "track", "area", "base", "link", "meta"):
            return
        if self._skip:
            return
        if t in ("b", "strong"):
            self._semantic_push("b")
        elif t in ("i", "em"):
            self._semantic_push("i")
        elif t == "u":
            self._semantic_push("u")
        elif t in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._soft_break()
            self._semantic_push(t)
        elif t == "small":
            self._semantic_push("small")
        elif t == "blockquote":
            self._block_break()
            self._semantic_push("blockquote")
        elif t == "pre":
            self._block_break()
            self._semantic_push("pre")
        elif t == "a":
            href = (ad.get("href") or "").strip()
            ltag = None
            if href:
                ltag = f"mail_link_{self._link_tag_seq}"
                self._link_tag_seq += 1
                self.tw.tag_configure(
                    ltag,
                    font=("Segoe UI", 11),
                    foreground="#0b57d0",
                    underline=True,
                )

                def open_url(_e, u=href):
                    webbrowser.open(u)

                self.tw.tag_bind(ltag, "<Button-1>", open_url)
                self.tw.tag_bind(ltag, "<Enter>", lambda _e: self.tw.config(cursor="hand2"))
                self.tw.tag_bind(ltag, "<Leave>", lambda _e: self.tw.config(cursor=""))
            self._link_stack.append((href, ltag))
        elif t == "br":
            self._insert("\n")
            self._last_was_block_break = True
        elif t == "hr":
            self._block_break()
            self._insert("—" * 28 + "\n")
            self._last_was_block_break = True
        elif t in ("ul", "ol"):
            self._block_break()
            self._semantic_push(t)
        elif t == "li":
            self._block_break()
            self._insert("• " if "ul" in self._stack else "◦ ")
            self._last_was_block_break = False
        elif t in self._BLOCK:
            self._block_break()
        elif t in ("td", "th"):
            self._insert("  ")

    def handle_endtag(self, tag):
        t = tag.lower()
        if t in _SKIP_HTML_SUBTREE:
            self._skip = max(0, self._skip - 1)
            return
        if self._skip:
            return
        if t in ("b", "strong"):
            self._semantic_pop("strong")
        elif t in ("i", "em"):
            self._semantic_pop("em")
        elif t == "u":
            self._semantic_pop("u")
        elif t in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._semantic_pop(t)
            self._block_break()
        elif t == "small":
            self._semantic_pop("small")
        elif t == "blockquote":
            self._semantic_pop("blockquote")
            self._block_break()
        elif t == "pre":
            self._semantic_pop("pre")
            self._block_break()
        elif t == "a":
            if self._link_stack:
                self._link_stack.pop()
        elif t in ("p", "div", "section", "article", "tr", "table"):
            self._block_break()
        elif t in ("ul", "ol"):
            self._semantic_pop(t)
            self._block_break()
        elif t == "li":
            self._block_break()

    def handle_data(self, d):
        if self._skip:
            return
        d = html.unescape(d)
        if not d:
            return
        d = re.sub(r"\n{3,}", "\n\n", d)
        self._insert(d)
        self._last_was_block_break = False


def render_html_mail_body(text_widget, html_source):
    """Fill text_widget with client-like rendered HTML (no raw tags)."""
    html_source = prepare_html_for_mail_preview(html_source)
    r = HTMLMailRenderer(text_widget)
    try:
        r.feed(html_source)
        r.close()
    except Exception:
        text_widget.insert(tk.END, clean_html(html_source), ("mail_base",))


def message_has_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get_filename():
            return True
        disp = part.get("Content-Disposition", "") or ""
        if "attachment" in disp.lower():
            return True
    return False


def extract_body_and_attachments(msg):
    """Returns (plain_text_for_search, attachments, html_raw_or_none).

    When an HTML part exists it is preferred for display (like typical mail clients); plain_text_for_search
    is still a normalized plain string for filtering (HTML cleaned to text, plus plain part if both exist).
    """
    attachments = []
    plain_raw = None
    html_raw = None
    for part in msg.walk():
        fn = part.get_filename()
        ctype = part.get_content_type()
        if fn:
            attachments.append(part)
        elif ctype == "text/plain" and plain_raw is None:
            plain_raw = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="replace")
        elif ctype == "text/html" and html_raw is None:
            html_raw = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="replace")

    plain_norm = normalize_plain_body(html.unescape(plain_raw)) if plain_raw else ""

    if html_raw:
        plain_for_search = clean_html(html_raw)
        if plain_norm:
            plain_for_search = (plain_for_search + " " + plain_norm).strip()
        return plain_for_search, attachments, html_raw
    return plain_norm, attachments, None


def search_blob_for_message(msg):
    """Lowercase string spanning sender, subject, body, and attachment filenames for search."""
    try:
        body, attachments, _html = extract_body_and_attachments(msg)
        from_hdr = msg.get("from", "") or ""
        _, addr = parseaddr(from_hdr)
        subj = str(msg.get("subject", ""))
        names = []
        for part in attachments:
            fn = part.get_filename()
            if fn:
                names.append(fn)
        attach_txt = " ".join(names)
        return " ".join([from_hdr, addr, subj, body, attach_txt]).lower()
    except Exception:
        return ""


def load_email_row(path, root_path):
    """Parse one .eml file into the same dict shape used by the table, or None on failure."""
    try:
        with open(path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        _, from_email = parseaddr(msg.get("from", ""))

        folder_display = folder_display_for_path(path, root_path)

        date_str = msg.get("date")
        dt_obj = None
        if date_str:
            try:
                dt_obj = parsedate_to_datetime(date_str)
            except Exception:
                pass

        display_date = dt_obj.strftime("%Y-%m-%d %H:%M") if dt_obj else "Unknown"
        has_att = message_has_attachments(msg)

        return {
            "folder": folder_display,
            "date": display_date,
            "from": from_email or "Unknown",
            "subject": str(msg.get("subject", "(No Subject)")),
            "path": path,
            "sort_key": dt_obj.timestamp() if dt_obj else 0,
            "has_attachments": has_att,
            "attach_label": "Yes" if has_att else "No",
            "search_blob": search_blob_for_message(msg),
        }
    except Exception:
        return None


class EMLViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("EML Archive Explorer")
        self.root.geometry("1400x850")

        self.path_map = {}
        self.current_attachments = []
        self.queue = queue.Queue()
        self.emails = []

        self.sort_column = "Date"
        self.sort_reverse = True

        self.filter_vars = {c: tk.StringVar() for c in COLUMNS}

        self.paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.tree_frame = ttk.Frame(self.paned)
        self.paned.add(self.tree_frame, weight=1)

        filter_outer = ttk.LabelFrame(self.tree_frame, text="Filters (substring match)", padding=6)
        filter_outer.pack(fill=tk.X, padx=4, pady=4)

        for i, col in enumerate(COLUMNS):
            ttk.Label(filter_outer, text=col + ":").grid(row=0, column=i * 2, sticky=tk.W, padx=(4, 2))
            ent = ttk.Entry(filter_outer, textvariable=self.filter_vars[col], width=18)
            ent.grid(row=0, column=i * 2 + 1, padx=(0, 8), pady=2, sticky=tk.EW)
            ent.bind("<KeyRelease>", self._on_filter_change)

        for c in range(len(COLUMNS) * 2):
            filter_outer.columnconfigure(c, weight=1 if c % 2 == 1 else 0)

        search_row = ttk.Frame(self.tree_frame)
        search_row.pack(fill=tk.X, padx=4, pady=(0, 4))
        ttk.Label(search_row, text="Search (subject, sender, body, attachment names):").pack(side=tk.LEFT, padx=(4, 8))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_row, textvariable=self.search_var, width=50)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        search_entry.bind("<KeyRelease>", self._on_filter_change)

        self.tree = ttk.Treeview(self.tree_frame, columns=COLUMNS, show="headings")

        for col in COLUMNS:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            w = 100 if col == "Attachments" else 140
            if col == "Subject":
                w = 260
            self.tree.column(col, width=w)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<<TreeviewSelect>>", self.on_file_select)

        self.content_frame = ttk.Frame(self.paned)
        self.paned.add(self.content_frame, weight=2)

        self.display_text = tk.Text(
            self.content_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 11),
            padx=15,
            pady=15,
            bg="#ffffff",
            fg="#222222",
        )
        self.display_text.pack(fill=tk.BOTH, expand=True)

        self.attach_frame = ttk.Frame(self.content_frame, relief=tk.FLAT, padding=5)
        self.attach_frame.pack(fill=tk.X)

        self.attach_info_label = ttk.Label(self.attach_frame, text="No attachments", foreground="#777777")
        self.attach_info_label.pack(side=tk.LEFT, padx=5)

        self.save_btn = ttk.Button(self.attach_frame, text="Save All", command=self.save_attachments, state="disabled")
        self.save_btn.pack(side=tk.RIGHT, padx=5)

        configure_mail_preview_tags(self.display_text)

        self._update_heading_labels()
        self.start_scan()
        self.process_queue()

    def _on_filter_change(self, _event=None):
        self.refresh_table()

    def _arrow_for_column(self, col):
        if col != self.sort_column:
            return ""
        return " \u25bc" if self.sort_reverse else " \u25b2"

    def _update_heading_labels(self):
        labels = {
            "Folder": "Folder",
            "Date": "Date",
            "From": "From",
            "Subject": "Subject",
            "Attachments": "Attachments",
        }
        for col in COLUMNS:
            self.tree.heading(col, text=labels[col] + self._arrow_for_column(col))

    def sort_by_column(self, col):
        if col == self.sort_column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = col == "Date"
        self._update_heading_labels()
        self.refresh_table()

    def refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.path_map.clear()

        filters = {c: self.filter_vars[c].get() for c in COLUMNS}
        rows = filtered_sorted_emails(
            self.emails,
            self.sort_column,
            self.sort_reverse,
            filters,
            global_search=self.search_var.get(),
        )

        for d in rows:
            vals = (
                d["folder"],
                d["date"],
                d["from"],
                d["subject"],
                d["attach_label"],
            )
            node = self.tree.insert("", "end", values=vals)
            self.path_map[node] = d["path"]

    def start_scan(self):
        def scan_task():
            root_path = os.getcwd()
            files_to_process = []

            for walk_root, _, files in os.walk(root_path):
                for f in files:
                    if f.lower().endswith(".eml"):
                        files_to_process.append(os.path.join(walk_root, f))

            total = max(len(files_to_process), 1)
            for i, path in enumerate(files_to_process):
                item_data = load_email_row(path, root_path)
                if item_data is None:
                    continue
                self.queue.put(("item", item_data))
                if i % 10 == 0:
                    self.queue.put(("progress", int((i + 1) / total * 100)))

            self.queue.put(("status", "Finished loading."))

        threading.Thread(target=scan_task, daemon=True).start()

    def process_queue(self):
        got_item = False
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == "item":
                    self.emails.append(msg[1])
                    got_item = True
                elif msg[0] == "progress":
                    pass
        except queue.Empty:
            pass
        if got_item:
            self.refresh_table()
        self.root.after(50, self.process_queue)

    def on_file_select(self, event):
        selection = self.tree.selection()
        if not selection:
            return
        file_path = self.path_map.get(selection[0])

        try:
            with open(file_path, "rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)

            plain, self.current_attachments, html_raw = extract_body_and_attachments(msg)

            self.display_text.config(state="normal")
            self.display_text.delete("1.0", tk.END)

            header = f"SENDER:  {msg.get('from')}\nSUBJECT: {msg.get('subject')}\nDATE:    {msg.get('date')}\n"
            header += "=" * 60 + "\n\n"

            self.display_text.insert(tk.END, header, ("mail_base",))
            if html_raw:
                render_html_mail_body(self.display_text, html_raw)
            else:
                self.display_text.insert(tk.END, plain, ("mail_base",))

            self.display_text.config(state="disabled")

            if self.current_attachments:
                names = [p.get_filename() for p in self.current_attachments]
                self.attach_info_label.config(text=f"Files: {', '.join(names)}", foreground="#999999")
                self.save_btn.config(state="normal")
            else:
                self.attach_info_label.config(text="No attachments", foreground="#cccccc")
                self.save_btn.config(state="disabled")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_attachments(self):
        target = filedialog.askdirectory()
        if not target:
            return
        for part in self.current_attachments:
            with open(os.path.join(target, part.get_filename()), "wb") as f:
                f.write(part.get_payload(decode=True))
        messagebox.showinfo("Success", "Attachments saved.")


if __name__ == "__main__":
    root = tk.Tk()
    app = EMLViewer(root)
    root.mainloop()
