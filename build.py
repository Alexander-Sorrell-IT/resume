#!/usr/bin/env python3
"""
build.py - Authentic & Faithful Resume Builder
Faithfully renders the exact content, structure, and text of the 3 resumes
with embedded Starreckon QR code, exact PDF/DOCX downloads, and responsive layout.
"""

import os
import shutil
import re
import html
import subprocess
import docx

ROOT = os.path.dirname(os.path.abspath(__file__))
RESUMES_DIR = os.path.join(ROOT, "resumes")
SITE_DIR = os.path.join(ROOT, "_site")
ASSETS_DIR = os.path.join(SITE_DIR, "assets")
DOWNLOADS_DIR = os.path.join(SITE_DIR, "downloads")

RESUME_CONFIGS = [
    {
        "id": "public",
        "title": "All Public Repos",
        "subtitle": "Consolidated Public Surface (Both Accounts)",
        "badge": "38 Repos · 170k LOC · 3,342 Tests",
        "docx_src": "Alexander Sorrell Resume — Public.docx",
        "pdf_name": "Alexander Sorrell Resume — Public.pdf",
        "docx_name": "Alexander Sorrell Resume — Public.docx",
        "default": True
    },
    {
        "id": "alexander-sorrell-it",
        "title": "Alexander-Sorrell-IT",
        "subtitle": "Primary Architecture & Verification",
        "badge": "33 Repos · 156k LOC · 3,168 Tests",
        "docx_src": "Alexander Sorrell Resume — Public (Alexander-Sorrell-IT).docx",
        "pdf_name": "Alexander Sorrell Resume — Public (Alexander-Sorrell-IT).pdf",
        "docx_name": "Alexander Sorrell Resume — Public (Alexander-Sorrell-IT).docx",
        "default": False
    },
    {
        "id": "matrixbuilderops",
        "title": "matrixbuilderops",
        "subtitle": "Agent Fleets & Operational Machinery",
        "badge": "5 Repos · 13.8k LOC · 174 Tests",
        "docx_src": "Alexander Sorrell Resume — Public (matrixbuilderops).docx",
        "pdf_name": "Alexander Sorrell Resume — Public (matrixbuilderops).pdf",
        "docx_name": "Alexander Sorrell Resume — Public (matrixbuilderops).docx",
        "default": False
    }
]

def clean(text):
    if not text:
        return ""
    return text.replace("\u00a0", " ").strip()

def is_border(line):
    s = line.strip()
    if not s:
        return True
    return set(s).issubset({"=", "-", "─", "│", " "})

def linkify(text):
    # Linkify github.com, linkedin.com, emails, and npx
    escaped = html.escape(text)
    # Linkify URLs
    escaped = re.sub(
        r"(https?://[^\s]+)",
        r'<a href="\1" target="_blank" rel="noopener">\1</a>',
        escaped
    )
    # Linkify github.com/xxx/yyy
    escaped = re.sub(
        r'(?<!")(github\.com/[a-zA-Z0-9_\-\./]+)',
        r'<a href="https://\1" target="_blank" rel="noopener">\1</a>',
        escaped
    )
    # Linkify linkedin.com/in/xxx
    escaped = re.sub(
        r'(?<!")(linkedin\.com/in/[a-zA-Z0-9_\-]+)',
        r'<a href="https://\1" target="_blank" rel="noopener">\1</a>',
        escaped
    )
    # Linkify email
    escaped = re.sub(
        r'([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)',
        r'<a href="mailto:\1">\1</a>',
        escaped
    )
    return escaped

def parse_docx_resume(file_path):
    doc = docx.Document(file_path)
    
    # Skills from Table 1
    skills = []
    if len(doc.tables) > 1:
        for row in doc.tables[1].rows:
            cat = clean(row.cells[0].text)
            details = clean(row.cells[1].text) if len(row.cells) > 1 else ""
            if cat and cat.lower() != "category":
                skills.append({"category": cat, "details": details})
                
    paras = [clean(p.text) for p in doc.paragraphs if clean(p.text)]
    return skills, paras

def render_pipe_table(pipe_lines, title=""):
    rows = []
    for l in pipe_lines:
        if "│" in l and not is_border(l):
            cells = [clean(c) for c in l.split("│")]
            rows.append(cells)
            
    if not rows:
        return ""
        
    headers = rows[0]
    data_rows = rows[1:]
    
    out = ["<div class='resume-table-wrapper'>"]
    if title:
        out.append(f"<h4 class='table-title'>{html.escape(title)}</h4>")
    out.append("<div class='table-scroll'><table class='resume-data-table'>")
    out.append("<thead><tr>")
    for h in headers:
        out.append(f"<th>{html.escape(h)}</th>")
    out.append("</tr></thead><tbody>")
    
    for r in data_rows:
        is_highlight = any(kw in "".join(r).upper() for kw in ["TOTAL", "ALL ECOSYSTEM", "AVERAGES"])
        row_cls = " class='highlight-row'" if is_highlight else ""
        out.append(f"<tr{row_cls}>")
        for idx, cell in enumerate(r):
            if idx == 0 and cell.lower() in ["pip", "npx"]:
                badge_cls = "badge-pip" if cell.lower() == "pip" else "badge-npx"
                out.append(f"<td><span class='pkg-pill {badge_cls}'>{html.escape(cell)}</span></td>")
            else:
                out.append(f"<td>{html.escape(cell)}</td>")
        out.append("</tr>")
        
    out.append("</tbody></table></div></div>")
    return "\n".join(out)

def render_resume_body(skills, paras):
    """
    Renders paragraphs faithfully following the exact sequence of the document.
    """
    html_out = []
    i = 0
    n = len(paras)
    
    while i < n:
        p = paras[i]
        
        # Skip pure ascii border lines
        if is_border(p):
            i += 1
            continue
            
        # Section Heading
        if p.isupper() and len(p) < 80 and not ("│" in p) and not ("•" in p):
            html_out.append(f"<h3 class='resume-heading'>{html.escape(p)}</h3>")
            
            # If this is Skills section, inject skills table right after
            if "TECHNICAL SKILLS" in p and skills:
                html_out.append("<div class='skills-table-wrap'><table class='skills-table'><tbody>")
                for sk in skills:
                    html_out.append(f"""
                    <tr>
                        <td class='skills-cat'>{html.escape(sk['category'])}</td>
                        <td class='skills-val'>{linkify(sk['details'])}</td>
                    </tr>
                    """)
                html_out.append("</tbody></table></div>")
            i += 1
            continue
            
        # Detect ASCII pipe table block
        if "│" in p:
            table_lines = [p]
            i += 1
            while i < n and ("│" in paras[i] or is_border(paras[i])):
                if not is_border(paras[i]):
                    table_lines.append(paras[i])
                i += 1
            html_out.append(render_pipe_table(table_lines))
            continue
            
        # Detect multi-line Project Entry (with lines, commits, Code:, etc.)
        if "\n" in p and (" — " in p or "Code:" in p or "lines" in p):
            sublines = [clean(s) for s in p.split("\n") if clean(s)]
            html_out.append("<div class='project-entry'>")
            first_line = sublines[0]
            if " — " in first_line and not first_line.startswith("•"):
                p_parts = first_line.split(" — ", 1)
                html_out.append(f"<p class='proj-headline'><strong class='proj-name'>{html.escape(p_parts[0])}</strong> — <span class='proj-summary'>{linkify(p_parts[1])}</span></p>")
            else:
                html_out.append(f"<p class='proj-headline'>{linkify(first_line)}</p>")
                
            for sl in sublines[1:]:
                if sl.startswith("Code:"):
                    html_out.append(f"<p class='proj-code-line'><strong>Code:</strong> {linkify(sl.replace('Code:', '').strip())}</p>")
                elif sl.startswith("Run:") or "Built on:" in sl:
                    html_out.append(f"<p class='proj-run-line'><strong>Run / Built on:</strong> {linkify(sl)}</p>")
                else:
                    html_out.append(f"<p class='proj-metrics-line'>{linkify(sl)}</p>")
                    
            html_out.append("</div>")
            i += 1
            continue

        # Bullet item
        if p.startswith("•") or p.startswith("-"):
            clean_b = re.sub(r"^[•\-\*]\s*", "", p)
            match = re.match(r"^([^.:]{4,55}[.:])\s*(.*)", clean_b)
            if match:
                head, tail = match.groups()
                html_out.append(f"<div class='resume-bullet'><span class='bullet-dot'>•</span><div class='bullet-content'><strong>{html.escape(head)}</strong> {linkify(tail)}</div></div>")
            else:
                html_out.append(f"<div class='resume-bullet'><span class='bullet-dot'>•</span><div class='bullet-content'>{linkify(clean_b)}</div></div>")
            i += 1
            continue
            
        # Experience Header (Role — Company · Location · Dates)
        if " — " in p and ("Present" in p or "202" in p or "Jan" in p or "Apr" in p):
            html_out.append(f"<h4 class='exp-role-title'>{html.escape(p)}</h4>")
            i += 1
            continue
            
        # CLI command
        if p.startswith("curl ") or p.startswith("npx "):
            html_out.append(f"<pre class='code-block'><code>{html.escape(p)}</code></pre>")
            i += 1
            continue
            
        # Standard paragraph
        html_out.append(f"<p class='resume-p'>{linkify(p)}</p>")
        i += 1
        
    return "\n".join(html_out)

def build_all():
    os.makedirs(ASSETS_DIR, exist_ok=True)
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)
    
    # 1. Convert docx to exact 1:1 PDFs with LibreOffice
    for cfg in RESUME_CONFIGS:
        src_docx = os.path.join(ROOT, cfg["docx_src"])
        if not os.path.exists(src_docx):
            src_docx = os.path.join("/home/phantomcore", cfg["docx_src"])
            
        # Copy original docx to downloads
        dest_docx = os.path.join(DOWNLOADS_DIR, cfg["docx_name"])
        shutil.copy2(src_docx, dest_docx)
        
        # Convert to PDF
        pdf_dest = os.path.join(DOWNLOADS_DIR, cfg["pdf_name"])
        if not os.path.exists(pdf_dest):
            cmd = f'libreoffice --headless --convert-to pdf "{src_docx}" --outdir "{DOWNLOADS_DIR}"'
            subprocess.run(cmd, shell=True, check=True)
            
    # 2. Copy assets (CSS, CNAME, Headshot, QR Code, Logo)
    for f in ["style.css", "CNAME"]:
        p = os.path.join(ROOT, f)
        if os.path.exists(p):
            shutil.copy2(p, os.path.join(SITE_DIR, f))
            
    headshot = os.path.join(ROOT, "alexander-sorrell-headshot.jpg")
    if os.path.exists(headshot):
        shutil.copy2(headshot, os.path.join(ASSETS_DIR, "alexander-sorrell.jpg"))
        
    qr_png = os.path.join(ROOT, "assets/starreckon-qr.png")
    if os.path.exists(qr_png):
        shutil.copy2(qr_png, os.path.join(ASSETS_DIR, "starreckon-qr.png"))
        
    fav = os.path.join(ROOT, "assets/favicon.svg")
    if os.path.exists(fav):
        shutil.copy2(fav, os.path.join(ASSETS_DIR, "favicon.svg"))
        shutil.copy2(fav, os.path.join(ASSETS_DIR, "logo.svg"))

    # 3. Parse and render resumes
    tab_buttons = []
    tab_panes = []
    
    for cfg in RESUME_CONFIGS:
        src_docx = os.path.join(ROOT, cfg["docx_src"])
        skills, paras = parse_docx_resume(src_docx)
        
        rendered_body = render_resume_body(skills, paras)
        active_cls = "active" if cfg["default"] else ""
        
        tab_buttons.append(f"""
        <button class="resume-tab-btn {active_cls}" data-tab="{cfg['id']}">
            <span class="btn-name">{cfg['title']}</span>
            <span class="btn-stat">{cfg['badge']}</span>
        </button>
        """)
        
        tab_panes.append(f"""
        <article id="{cfg['id']}" class="resume-sheet {active_cls}">
            <!-- Resume Header Table (Matching Word Document) -->
            <header class="resume-doc-header">
                <div class="header-left-col">
                    <h1 class="header-name">Alexander Sorrell</h1>
                    <p class="header-role">AI Principal / Builder / Breaker — Verification &amp; Evaluation</p>
                    <p class="header-contact">
                        (817) 996-6123 · <a href="mailto:alexander.sorrell.it@gmail.com">alexander.sorrell.it@gmail.com</a><br>
                        <a href="https://www.linkedin.com/in/alexander-sorrell-it/" target="_blank" rel="noopener">linkedin.com/in/alexander-sorrell-it</a> · 
                        <a href="https://github.com/Alexander-Sorrell-IT" target="_blank" rel="noopener">github.com/Alexander-Sorrell-IT</a><br>
                        Colleyville, TX (Dallas–Fort Worth) — open to remote, onsite, or relocation
                    </p>
                </div>
                
                <div class="header-right-col">
                    <a class="qr-link-box" href="https://alexander-sorrell-it.github.io/starreckon/#s=30.1&amp;g=MASTERWORK&amp;a=The_Systems_Architect&amp;v=7.0%2C6.6%2C6.1%2C5.1%2C5.3&amp;ss=14264&amp;h=1290&amp;d=75&amp;k=49&amp;tok=109.4B&amp;n=Alexander+Sorrell&amp;gh=Alexander-Sorrell-IT&amp;li=alexander-sorrell-it&amp;em=alexander.sorrell.it%40gmail.com&amp;web=github.com%2FAlexander-Sorrell-IT" target="_blank" rel="noopener" title="Scan or click to view verified Starreckon audit card">
                        <img class="resume-qr-img" src="/assets/starreckon-qr.png" alt="Starreckon Proof of Work QR Code" width="115" height="115">
                    </a>
                    <div class="qr-caption">
                        <a href="https://github.com/Alexander-Sorrell-IT/starreckon" target="_blank" rel="noopener">github.com/Alexander-Sorrell-IT/starreckon</a><br>
                        <code>npx starreckon</code><br>
                        <strong class="audit-verdict">VERIFIED AUDIT · 30.1 MASTERWORK</strong>
                    </div>
                </div>
            </header>

            <!-- Download Action Bar for This Resume -->
            <div class="download-bar no-print">
                <div class="download-label">
                    <span>Active Version: <strong>{cfg['title']}</strong></span>
                    <span class="download-badge">{cfg['badge']}</span>
                </div>
                <div class="download-buttons">
                    <a class="dl-btn dl-pdf" href="/downloads/{cfg['pdf_name']}" download>
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
                        Download PDF
                    </a>
                    <a class="dl-btn dl-docx" href="/downloads/{cfg['docx_name']}" download>
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                        Download DOCX
                    </a>
                    <button class="dl-btn dl-print" onclick="window.print()">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
                        Print
                    </button>
                </div>
            </div>

            <!-- Resume Content -->
            <div class="resume-sheet-content">
                {rendered_body}
            </div>
        </article>
        """)

    index_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0">
    <title>Alexander Sorrell — Resume</title>
    <meta name="description" content="Alexander Sorrell — AI Principal / Builder / Breaker — Verification &amp; Evaluation.">
    <link rel="canonical" href="https://www.alexander-sorrell-it-resume.dev/">
    <meta property="og:title" content="Alexander Sorrell — Resume">
    <meta property="og:description" content="Models propose. A deterministic function owns the verdict. Every decision ships with a receipt a stranger can re-check.">
    <meta property="og:url" content="https://www.alexander-sorrell-it-resume.dev/">
    <meta property="og:type" content="website">
    <link rel="icon" type="image/svg+xml" href="/assets/favicon.svg">
    <link rel="stylesheet" href="/style.css">
</head>
<body>

    <!-- Top Navigation Bar -->
    <header class="top-nav-bar no-print">
        <div class="site-container nav-inner">
            <a class="nav-brand" href="/">
                <img src="/assets/favicon.svg" alt="Alexander Sorrell" width="24" height="24">
                <span>Alexander<strong>Sorrell</strong></span>
            </a>
            <div class="nav-domain-label">
                <span>www.alexander-sorrell-it-resume.dev</span>
            </div>
            <nav class="nav-links">
                <a href="https://www.linkedin.com/in/alexander-sorrell-it/" target="_blank" rel="noopener">LinkedIn</a>
                <a href="https://github.com/Alexander-Sorrell-IT" target="_blank" rel="noopener">GitHub</a>
                <a href="mailto:alexander.sorrell.it@gmail.com">Contact</a>
            </nav>
        </div>
    </header>

    <main class="site-container main-wrapper">
        <!-- Resume Tabs Selector -->
        <nav class="tabs-header-nav no-print">
            <div class="tabs-label">Select Tailored Resume Version:</div>
            <div class="tabs-buttons-row">
                {''.join(tab_buttons)}
            </div>
        </nav>

        <!-- Resume Documents (One per tab) -->
        <section class="resume-sheets-container">
            {''.join(tab_panes)}
        </section>
    </main>

    <footer class="site-footer no-print">
        <div class="site-container footer-inner">
            <p>Alexander Sorrell · Sovereign Local Verification &amp; Deterministic Evaluation</p>
            <p><a href="https://github.com/Alexander-Sorrell-IT" target="_blank" rel="noopener">GitHub Profile</a> · <a href="https://www.linkedin.com/in/alexander-sorrell-it/" target="_blank" rel="noopener">LinkedIn</a> · <a href="https://github.com/Alexander-Sorrell-IT/starreckon" target="_blank" rel="noopener">Starreckon Audit Card</a></p>
        </div>
    </footer>

    <script>
        const tabBtns = document.querySelectorAll('.resume-tab-btn');
        const sheets = document.querySelectorAll('.resume-sheet');

        function selectTab(tabId) {{
            tabBtns.forEach(btn => {{
                if (btn.dataset.tab === tabId) {{
                    btn.classList.add('active');
                }} else {{
                    btn.classList.remove('active');
                }}
            }});

            sheets.forEach(sheet => {{
                if (sheet.id === tabId) {{
                    sheet.classList.add('active');
                }} else {{
                    sheet.classList.remove('active');
                }}
            }});

            if (history.replaceState) {{
                history.replaceState(null, null, '#' + tabId);
            }}
        }}

        tabBtns.forEach(btn => {{
            btn.addEventListener('click', () => {{
                selectTab(btn.dataset.tab);
            }});
        }});

        const initialHash = window.location.hash.replace('#', '');
        if (initialHash && document.getElementById(initialHash)) {{
            selectTab(initialHash);
        }}
    </script>
</body>
</html>"""

    with open(os.path.join(SITE_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)
        
    print(f"[build] Rendered authentic resumes into {SITE_DIR}")

if __name__ == "__main__":
    build_all()
