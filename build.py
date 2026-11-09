#!/usr/bin/env python3
"""
build.py - Executive Resume Site Builder
Parses .docx resumes into clean, human-readable, executive web documents.
Completely eliminates horizontal overflow, ASCII border art, and cramping.
Fully responsive on mobile, tablet, and desktop.
"""

import os
import shutil
import re
import html
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
        "file": "Alexander_Sorrell_Resume_Public_All.docx",
        "orig_file": "Alexander Sorrell Resume — Public.docx",
        "default": True
    },
    {
        "id": "alexander-sorrell-it",
        "title": "Alexander-Sorrell-IT",
        "subtitle": "Primary Architecture & Verification",
        "badge": "33 Repos · 156k LOC · 3,168 Tests",
        "file": "Alexander_Sorrell_Resume_Alexander_Sorrell_IT.docx",
        "orig_file": "Alexander Sorrell Resume — Public (Alexander-Sorrell-IT).docx",
        "default": False
    },
    {
        "id": "matrixbuilderops",
        "title": "matrixbuilderops",
        "subtitle": "Agent Fleets & Operational Machinery",
        "badge": "5 Repos · 13.8k LOC · 174 Tests",
        "file": "Alexander_Sorrell_Resume_matrixbuilderops.docx",
        "orig_file": "Alexander Sorrell Resume — Public (matrixbuilderops).docx",
        "default": False
    }
]

def clean_text(text):
    if not text:
        return ""
    # Strip non-breaking spaces and trim
    return text.replace("\u00a0", " ").strip()

def is_border_line(line):
    s = line.strip()
    if not s:
        return True
    chars = set(s)
    # If the line consists only of =, -, ─, │, spaces
    if chars.issubset({"=", "-", "─", "│", " "}):
        return True
    return False

def parse_docx(path):
    doc = docx.Document(path)
    
    # Skills from Table 1
    skills = []
    if len(doc.tables) > 1:
        for row in doc.tables[1].rows:
            cat = clean_text(row.cells[0].text)
            details = clean_text(row.cells[1].text) if len(row.cells) > 1 else ""
            if cat and cat != "Category":
                skills.append({"category": cat, "details": details})
                
    raw_paras = [clean_text(p.text) for p in doc.paragraphs if clean_text(p.text)]
    
    # Filter out pure ASCII dividers (=== and ---)
    paras = [p for p in raw_paras if not is_border_line(p)]
    
    # Group into structured sections
    sections = {}
    current_sec = "SUMMARY"
    sections[current_sec] = []
    
    headings = [
        "PROFESSIONAL SUMMARY",
        "PROOF OF WORK",
        "ARTIFICIAL INTELLIGENCE & TECHNICAL SKILLS",
        "PROFESSIONAL EXPERIENCE",
        "BY THE NUMBERS",
        "PROJECTS",
        "LIVE TERMINAL TELEMETRY",
        "PROGRAM DIRECTORY",
        "QR CODE TARGET ENGINE",
        "PUBLIC INDEPENDENT VERIFICATION ENDPOINTS"
    ]
    
    for p in paras:
        matched = False
        upper_p = p.upper()
        for h in headings:
            if h in upper_p and len(p) < 90 and not ("│" in p):
                current_sec = h
                if current_sec not in sections:
                    sections[current_sec] = []
                matched = True
                break
        if not matched:
            sections[current_sec].append(p)
            
    return {
        "skills": skills,
        "sections": sections
    }

def format_bullet_html(line):
    clean = re.sub(r"^[•\-\*]\s*", "", line).strip()
    # If starts with bold intro: "Title. description..." or "Title: description"
    match = re.match(r"^([^.:]{4,50}[.:])\s*(.*)", clean)
    if match:
        lead, rest = match.groups()
        return f"<li><strong class='bullet-lead'>{html.escape(lead)}</strong> <span>{html.escape(rest)}</span></li>"
    return f"<li><span>{html.escape(clean)}</span></li>"

def render_pipe_table_html(lines, title=""):
    """
    Parses ASCII pipe tables ('Type │ Package │ ...') into a responsive HTML table.
    """
    rows = []
    for l in lines:
        if "│" in l:
            cells = [clean_text(c) for c in l.split("│")]
            rows.append(cells)
            
    if not rows:
        return ""
    
    headers = rows[0]
    data_rows = rows[1:]
    
    html_out = ["<div class='table-card'>"]
    if title:
        html_out.append(f"<h4 class='table-card-title'>{html.escape(title)}</h4>")
    html_out.append("<div class='table-responsive'><table class='resume-table'>")
    html_out.append("<thead><tr>")
    for h in headers:
        html_out.append(f"<th>{html.escape(h)}</th>")
    html_out.append("</tr></thead><tbody>")
    
    for r in data_rows:
        # Check if row is a summary or subheader
        is_subtotal = any(kw in "".join(r).upper() for kw in ["TOTAL", "ALL ECOSYSTEM", "AVERAGES"])
        row_cls = "row-subtotal" if is_subtotal else ""
        html_out.append(f"<tr class='{row_cls}'>")
        for idx, cell in enumerate(r):
            # Badge for pip/npx
            if idx == 0 and cell.lower() in ["pip", "npx"]:
                badge_cls = "badge-pip" if cell.lower() == "pip" else "badge-npx"
                html_out.append(f"<td><span class='pkg-badge {badge_cls}'>{html.escape(cell)}</span></td>")
            else:
                html_out.append(f"<td>{html.escape(cell)}</td>")
        html_out.append("</tr>")
        
    html_out.append("</tbody></table></div></div>")
    return "\n".join(html_out)

def render_projects_section(lines):
    html_out = ["<div class='resume-section projects-section'>", "<h3>Key Systems &amp; Verified Repositories</h3>"]
    
    intro_lines = []
    current_proj = None
    proj_lines = []
    
    for l in lines:
        # Subheaders like "AI EVALUATION, VERIFICATION & RED-TEAMING"
        if l.isupper() and len(l) < 70 and not ("—" in l):
            if current_proj:
                html_out.append(render_single_project(current_proj, proj_lines))
                current_proj = None
                proj_lines = []
            html_out.append(f"<h4 class='proj-group-title'>{html.escape(l)}</h4>")
            continue
            
        # Detect project headline: "name — summary"
        if " — " in l and not l.startswith("•") and len(l.split(" — ")[0]) < 40:
            if current_proj:
                html_out.append(render_single_project(current_proj, proj_lines))
            current_proj = l
            proj_lines = []
        else:
            if current_proj:
                proj_lines.append(l)
            else:
                intro_lines.append(l)
                
    if current_proj:
        html_out.append(render_single_project(current_proj, proj_lines))
        
    for il in intro_lines:
        html_out.insert(2, f"<p class='section-intro'>{html.escape(il)}</p>")
        
    html_out.append("</div>")
    return "\n".join(html_out)

def render_single_project(header_line, lines):
    parts = header_line.split(" — ", 1)
    name = parts[0].strip()
    tagline = parts[1].strip() if len(parts) > 1 else ""
    
    metrics = []
    code_info = []
    run_info = []
    other = []
    
    for l in lines:
        if any(w in l for w in ["lines", "commits", "tests passing", "downloads"]):
            metrics.append(l)
        elif l.startswith("Code:"):
            code_info.append(l.replace("Code:", "").strip())
        elif l.startswith("Run:") or "Built on:" in l:
            run_info.append(l)
        else:
            other.append(l)
            
    out = [f"<div class='project-card'>"]
    out.append(f"<div class='project-header'>")
    out.append(f"<h4 class='project-name'>{html.escape(name)}</h4>")
    if tagline:
        out.append(f"<p class='project-tagline'>{html.escape(tagline)}</p>")
    out.append("</div>")
    
    if metrics:
        out.append("<div class='project-metrics-pills'>")
        for m in metrics:
            pills = [p.strip() for p in m.split(" · ") if p.strip()]
            for p in pills:
                if "github.com" in p:
                    out.append(f"<a class='metric-pill link' href='https://{p}' target='_blank' rel='noopener'>GitHub ↗</a>")
                elif "downloads" in p:
                    out.append(f"<span class='metric-pill dl'>{html.escape(p)}</span>")
                elif "tests passing" in p:
                    out.append(f"<span class='metric-pill test'>✔ {html.escape(p)}</span>")
                else:
                    out.append(f"<span class='metric-pill'>{html.escape(p)}</span>")
        out.append("</div>")
        
    for o in other:
        out.append(f"<p class='project-desc'>{html.escape(o)}</p>")
        
    if code_info:
        for c in code_info:
            out.append(f"<p class='project-meta'><strong>Architecture:</strong> {html.escape(c)}</p>")
            
    if run_info:
        for r in run_info:
            out.append(f"<p class='project-meta tech-stack'><strong>Tech Stack &amp; CLI:</strong> {html.escape(r)}</p>")
            
    out.append("</div>")
    return "\n".join(out)

def render_section_content(sec_name, lines):
    if not lines:
        return ""
        
    # Check if lines contain pipe table
    pipe_lines = [l for l in lines if "│" in l]
    non_pipe_lines = [l for l in lines if not ("│" in l)]
    
    html_parts = []
    
    if non_pipe_lines:
        # Check bullets
        if any(l.startswith("•") or l.startswith("-") for l in non_pipe_lines):
            html_parts.append("<ul class='resume-bullet-list'>")
            for l in non_pipe_lines:
                if l.startswith("•") or l.startswith("-"):
                    html_parts.append(format_bullet_html(l))
                else:
                    html_parts.append(f"<p class='resume-para'>{html.escape(l)}</p>")
            html_parts.append("</ul>")
        else:
            for l in non_pipe_lines:
                if " — " in l and ("Present" in l or "202" in l or "Jan" in l or "Apr" in l):
                    html_parts.append(f"<div class='exp-block'><h4 class='exp-header'>{html.escape(l)}</h4>")
                elif l.startswith("curl "):
                    html_parts.append(f"<div class='cli-command-box'><code>{html.escape(l)}</code></div>")
                else:
                    html_parts.append(f"<p class='resume-para'>{html.escape(l)}</p>")
            if "exp-block" in "".join(html_parts) and not html_parts[-1].endswith("</div>"):
                html_parts.append("</div>")
                
    if pipe_lines:
        html_parts.append(render_pipe_table_html(pipe_lines))
        
    return "\n".join(html_parts)

def build_site():
    os.makedirs(ASSETS_DIR, exist_ok=True)
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)
    
    # Copy assets
    css_src = os.path.join(ROOT, "style.css")
    if os.path.exists(css_src):
        shutil.copy2(css_src, os.path.join(SITE_DIR, "style.css"))
    cname_src = os.path.join(ROOT, "CNAME")
    if os.path.exists(cname_src):
        shutil.copy2(cname_src, os.path.join(SITE_DIR, "CNAME"))
    headshot_src = os.path.join(ROOT, "alexander-sorrell-headshot.jpg")
    if os.path.exists(headshot_src):
        shutil.copy2(headshot_src, os.path.join(ASSETS_DIR, "alexander-sorrell.jpg"))
    bastion_logo = "/media/phantomcore/AI_DRIVE/websites/bastion-fold/src/assets/logo.svg"
    if os.path.exists(bastion_logo):
        shutil.copy2(bastion_logo, os.path.join(ASSETS_DIR, "logo.svg"))

    rendered_tabs = []
    tab_nav = []
    
    for cfg in RESUME_CONFIGS:
        docx_path = os.path.join(RESUMES_DIR, cfg["file"])
        if not os.path.exists(docx_path):
            docx_path = os.path.join(ROOT, cfg["orig_file"])
            
        shutil.copy2(docx_path, os.path.join(DOWNLOADS_DIR, cfg["file"]))
        parsed = parse_docx(docx_path)
        
        resume_html = []
        
        # 1. Summary
        sum_lines = parsed["sections"].get("SUMMARY", []) + parsed["sections"].get("PROFESSIONAL SUMMARY", [])
        if sum_lines:
            resume_html.append("<section class='resume-section summary-section'>")
            resume_html.append("<h3 class='section-title'>Professional Summary</h3>")
            for sl in sum_lines:
                resume_html.append(f"<p class='lede-para'>{html.escape(sl)}</p>")
            resume_html.append("</section>")
            
        # 2. Proof of Work
        pow_lines = parsed["sections"].get("PROOF OF WORK", [])
        if pow_lines:
            resume_html.append("<section class='resume-section pow-section'>")
            resume_html.append("<h3 class='section-title'>Proof of Work &amp; Verification</h3>")
            resume_html.append(render_section_content("PROOF OF WORK", pow_lines))
            resume_html.append("</section>")
            
        # 3. Technical Skills
        if parsed["skills"]:
            resume_html.append("<section class='resume-section skills-section'>")
            resume_html.append("<h3 class='section-title'>Artificial Intelligence &amp; Technical Skills</h3>")
            resume_html.append("<div class='skills-grid'>")
            for sk in parsed["skills"]:
                resume_html.append(f"""
                <div class='skill-item'>
                    <span class='skill-label'>{html.escape(sk['category'])}</span>
                    <p class='skill-desc'>{html.escape(sk['details'])}</p>
                </div>
                """)
            resume_html.append("</div></section>")
            
        # 4. Professional Experience
        exp_lines = parsed["sections"].get("PROFESSIONAL EXPERIENCE", [])
        if exp_lines:
            resume_html.append("<section class='resume-section exp-section'>")
            resume_html.append("<h3 class='section-title'>Professional Experience</h3>")
            resume_html.append(render_section_content("PROFESSIONAL EXPERIENCE", exp_lines))
            resume_html.append("</section>")
            
        # 5. By The Numbers
        btn_key = next((k for k in parsed["sections"] if "BY THE NUMBERS" in k), None)
        if btn_key and parsed["sections"][btn_key]:
            resume_html.append("<section class='resume-section metrics-section'>")
            resume_html.append("<h3 class='section-title'>Measured Repository Metrics</h3>")
            resume_html.append(render_section_content(btn_key, parsed["sections"][btn_key]))
            resume_html.append("</section>")
            
        # 6. Projects & Repositories
        proj_key = next((k for k in parsed["sections"] if "PROJECTS" in k), None)
        if proj_key and parsed["sections"][proj_key]:
            resume_html.append(render_projects_section(parsed["sections"][proj_key]))
            
        # 7. Live Package Telemetry
        tele_key = next((k for k in parsed["sections"] if "LIVE TERMINAL TELEMETRY" in k or "PROGRAM DIRECTORY" in k), None)
        if tele_key and parsed["sections"][tele_key]:
            resume_html.append("<section class='resume-section telemetry-section'>")
            resume_html.append("<h3 class='section-title'>Package Telemetry &amp; Registry Verification</h3>")
            resume_html.append(render_section_content(tele_key, parsed["sections"][tele_key]))
            resume_html.append("</section>")

        active_class = "active" if cfg["default"] else ""
        
        tab_nav.append(f"""
        <button class="tab-button {active_class}" data-target="{cfg['id']}">
            <span class="tab-btn-title">{cfg['title']}</span>
            <span class="tab-btn-badge">{cfg['badge']}</span>
        </button>
        """)
        
        rendered_tabs.append(f"""
        <article id="{cfg['id']}" class="resume-tab-pane {active_class}">
            <div class="resume-tab-header">
                <div>
                    <span class="dossier-pill">{cfg['badge']}</span>
                    <h2 class="dossier-title">{cfg['title']}</h2>
                    <p class="dossier-subtitle">{cfg['subtitle']}</p>
                </div>
                <div class="dossier-actions no-print">
                    <a class="action-btn primary" href="/downloads/{cfg['file']}" download>
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                        Download DOCX
                    </a>
                    <button class="action-btn secondary" onclick="window.print()">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
                        Print / PDF
                    </button>
                </div>
            </div>
            
            <div class="resume-body-content">
                {''.join(resume_html)}
            </div>
        </article>
        """)

    index_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0">
    <title>Alexander Sorrell — Interactive &amp; Verifiable Resume</title>
    <meta name="description" content="AI Principal / Builder / Breaker — Verification &amp; Evaluation. Interactive resume across 38 public repositories with verifiable telemetry and reproducible benchmarks.">
    <link rel="canonical" href="https://www.alexander-sorrell-it-resume.dev/">
    <meta property="og:title" content="Alexander Sorrell — Resume">
    <meta property="og:description" content="Models propose. A deterministic function owns the verdict. Every decision ships with a receipt a stranger can re-check.">
    <meta property="og:url" content="https://www.alexander-sorrell-it-resume.dev/">
    <meta property="og:type" content="website">
    <link rel="icon" type="image/svg+xml" href="/assets/logo.svg">
    <link rel="stylesheet" href="/style.css">
</head>
<body>

    <!-- Header bar -->
    <header class="site-header no-print">
        <div class="page-wrap header-inner">
            <a class="brand-link" href="https://bastionfold.com" target="_blank" rel="noopener">
                <img src="/assets/logo.svg" alt="Bastion Fold" width="26" height="26">
                <span>Bastion<strong>Fold</strong></span>
            </a>
            <div class="header-domain-pill">
                <span>www.alexander-sorrell-it-resume.dev</span>
            </div>
            <nav class="header-nav">
                <a href="https://bastionfold.com/get-to-know-us/" target="_blank" rel="noopener">About Me</a>
                <a href="https://www.linkedin.com/in/alexander-sorrell-it/" target="_blank" rel="noopener">LinkedIn</a>
                <a href="https://github.com/Alexander-Sorrell-IT" target="_blank" rel="noopener">GitHub</a>
            </nav>
        </div>
    </header>

    <main class="page-wrap main-content">
        <!-- Profile Executive Card -->
        <section class="profile-hero">
            <div class="profile-photo-container">
                <img class="profile-avatar" src="/assets/alexander-sorrell.jpg" alt="Alexander Sorrell" width="150" height="200">
            </div>
            <div class="profile-details">
                <div class="profile-badges">
                    <span class="badge-status">AVAILABLE FOR ROLES</span>
                    <span class="badge-audit">★ 30.1 MASTERWORK · SYSTEMS ARCHITECT</span>
                </div>
                <h1 class="profile-title">Alexander Sorrell</h1>
                <p class="profile-subtitle">AI Principal / Builder / Breaker — Verification &amp; Evaluation</p>
                <blockquote class="profile-quote">
                    <strong>Central Thesis:</strong> Models propose. A deterministic function owns the verdict. Every decision ships with a receipt a stranger can re-check — a rule ID, a cryptographic hash chain, or the OS kernel itself.
                </blockquote>
                <div class="profile-contact-bar no-print">
                    <a class="contact-btn" href="https://www.linkedin.com/in/alexander-sorrell-it/" target="_blank" rel="noopener">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M19 0h-14c-2.761 0-5 2.239-5 5v14c0 2.761 2.239 5 5 5h14c2.762 0 5-2.239 5-5v-14c0-2.761-2.238-5-5-5zm-11 19h-3v-11h3v11zm-1.5-12.268c-.966 0-1.75-.79-1.75-1.764s.784-1.764 1.75-1.764 1.75.79 1.75 1.764-.783 1.764-1.75 1.764zm13.5 12.268h-3v-5.604c0-3.368-4-3.113-4 0v5.604h-3v-11h3v1.765c1.396-2.586 7-2.777 7 2.476v6.759z"/></svg>
                        LinkedIn
                    </a>
                    <a class="contact-btn" href="https://github.com/Alexander-Sorrell-IT" target="_blank" rel="noopener">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M12 0c-6.626 0-12 5.373-12 12 0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23.957-.266 1.983-.399 3.003-.404 1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.921.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576 4.765-1.589 8.199-6.086 8.199-11.386 0-6.627-5.373-12-12-12z"/></svg>
                        Alexander-Sorrell-IT
                    </a>
                    <a class="contact-btn" href="https://github.com/matrixbuilderops" target="_blank" rel="noopener">
                        matrixbuilderops
                    </a>
                    <a class="contact-btn" href="mailto:alexander.sorrell.it@gmail.com">
                        ✉ alexander.sorrell.it@gmail.com
                    </a>
                    <span class="contact-text">✆ (817) 996-6123</span>
                    <span class="contact-text">📍 Colleyville, TX</span>
                </div>
            </div>
        </section>

        <!-- Tailored Dossiers Tabs -->
        <section class="tab-bar-section no-print">
            <div class="tab-bar-title">Select Resume Focus:</div>
            <div class="tab-bar-grid">
                {''.join(tab_nav)}
            </div>
        </section>

        <!-- Resume Panes -->
        <section class="resume-panes-wrapper">
            {''.join(rendered_tabs)}
        </section>
    </main>

    <footer class="site-footer no-print">
        <div class="page-wrap footer-inner">
            <p>Bastion Fold LLC · SignalCore Inc · Sovereign local verification.</p>
            <p><a href="https://bastionfold.com" target="_blank">Bastion Fold Main Website</a> · <a href="https://github.com/Alexander-Sorrell-IT/starreckon" target="_blank">Starreckon Audit Card</a></p>
        </div>
    </footer>

    <script>
        const buttons = document.querySelectorAll('.tab-button');
        const panes = document.querySelectorAll('.resume-tab-pane');

        function activateTab(targetId) {{
            buttons.forEach(b => {{
                if (b.dataset.target === targetId) {{
                    b.classList.add('active');
                }} else {{
                    b.classList.remove('active');
                }}
            }});

            panes.forEach(p => {{
                if (p.id === targetId) {{
                    p.classList.add('active');
                }} else {{
                    p.classList.remove('active');
                }}
            }});

            if (history.replaceState) {{
                history.replaceState(null, null, '#' + targetId);
            }}
        }}

        buttons.forEach(b => {{
            b.addEventListener('click', () => {{
                activateTab(b.dataset.target);
            }});
        }});

        const hash = window.location.hash.replace('#', '');
        if (hash && document.getElementById(hash)) {{
            activateTab(hash);
        }}
    </script>
</body>
</html>"""

    with open(os.path.join(SITE_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)
        
    print(f"[build] Built executive site with {len(RESUME_CONFIGS)} tabs into {SITE_DIR}")

if __name__ == "__main__":
    build_site()
