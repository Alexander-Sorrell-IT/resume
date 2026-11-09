#!/usr/bin/env python3
"""
build.py - Standalone Resume Site Builder for Cloudflare Pages
Reads all .docx resumes in resumes/, converts them to structured web content,
and outputs a complete, responsive, tabbed static site into _site/.
"""

import os
import shutil
import json
import re
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
    },
    {
        "id": "all-repos",
        "title": "All Repositories (Master)",
        "subtitle": "Complete Verifiable Dossier (Public + Private)",
        "badge": "119 Repos · 786k LOC · 6,150 Tests",
        "file": "Alexander_Sorrell_Resume_All_Repos.docx",
        "orig_file": "Alexander Sorrell Resume — All Repos.docx",
        "default": False
    }
]

def parse_docx(path):
    doc = docx.Document(path)
    
    # Skills from Table 1
    skills = []
    if len(doc.tables) > 1:
        for row in doc.tables[1].rows:
            cat = row.cells[0].text.strip()
            details = row.cells[1].text.strip() if len(row.cells) > 1 else ""
            if cat:
                skills.append({"category": cat, "details": details})
                
    # Header from Table 0
    header_text = ""
    if len(doc.tables) > 0:
        cells = [c.text.strip() for r in doc.tables[0].rows for c in r.cells]
        header_text = " \n ".join(filter(None, cells))
        
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    # Section split
    sections = {}
    current_sec = "SUMMARY"
    sections[current_sec] = []
    
    heading_keywords = [
        "PROFESSIONAL SUMMARY",
        "PROOF OF WORK",
        "ARTIFICIAL INTELLIGENCE & TECHNICAL SKILLS",
        "PROFESSIONAL EXPERIENCE",
        "BY THE NUMBERS",
        "PROJECTS",
        "PROGRAM DIRECTORY",
        "QR CODE TARGET ENGINE",
        "PUBLIC INDEPENDENT VERIFICATION ENDPOINTS"
    ]
    
    for p in paragraphs:
        matched = False
        for hk in heading_keywords:
            if hk in p.upper() and len(p) < 80:
                current_sec = hk
                if current_sec not in sections:
                    sections[current_sec] = []
                matched = True
                break
        if not matched:
            sections[current_sec].append(p)
            
    return {
        "header_raw": header_text,
        "skills": skills,
        "sections": sections
    }

def format_bullet(line):
    # Strip bullet prefix
    clean = re.sub(r"^[•\-\*]\s*", "", line)
    # Highlight bold parts if "• Bold sentence. Remaining"
    parts = clean.split(". ", 1)
    if len(parts) == 2 and len(parts[0]) < 60:
        return f"<li><strong>{parts[0]}.</strong> {parts[1]}</li>"
    return f"<li>{clean}</li>"

def render_section_html(title, lines):
    if not lines:
        return ""
    
    html = [f"<div class='resume-section'>", f"<h3>{title}</h3>"]
    
    # If bullets
    if any(l.startswith("•") or l.startswith("-") for l in lines):
        html.append("<ul class='bullet-list'>")
        for l in lines:
            if l.startswith("•") or l.startswith("-"):
                html.append(format_bullet(l))
            else:
                html.append(f"<p>{l}</p>")
        html.append("</ul>")
    else:
        for l in lines:
            # Check if role header: "Title — Company · Location · Dates"
            if " — " in l and (" · " in l or "–" in l):
                html.append(f"<div class='experience-item'><h4 class='exp-title'>{l}</h4>")
            elif l.startswith("="):
                continue
            elif "│" in l or "─" in l:
                # ascii table line
                html.append(f"<pre class='ascii-table'>{l}</pre>")
            else:
                html.append(f"<p>{l}</p>")
                
    html.append("</div>")
    return "\n".join(html)

def build_site():
    os.makedirs(ASSETS_DIR, exist_ok=True)
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)
    
    # Copy style.css
    css_src = os.path.join(ROOT, "style.css")
    if os.path.exists(css_src):
        shutil.copy2(css_src, os.path.join(SITE_DIR, "style.css"))
    
    # Copy headshot and logo if present
    headshot_src = os.path.join(ROOT, "alexander-sorrell-headshot.jpg")
    if os.path.exists(headshot_src):
        shutil.copy2(headshot_src, os.path.join(ASSETS_DIR, "alexander-sorrell.jpg"))
        
    bastion_logo = "/media/phantomcore/AI_DRIVE/websites/bastion-fold/src/assets/logo.svg"
    if os.path.exists(bastion_logo):
        shutil.copy2(bastion_logo, os.path.join(ASSETS_DIR, "logo.svg"))

    # Parse each resume
    rendered_tabs = []
    tab_nav = []
    
    for cfg in RESUME_CONFIGS:
        docx_path = os.path.join(RESUMES_DIR, cfg["file"])
        if not os.path.exists(docx_path):
            docx_path = os.path.join(ROOT, cfg["orig_file"])
            
        shutil.copy2(docx_path, os.path.join(DOWNLOADS_DIR, cfg["file"]))
        
        parsed = parse_docx(docx_path)
        
        # Build HTML for this resume
        resume_html = []
        
        # Summary
        summary_lines = parsed["sections"].get("PROFESSIONAL SUMMARY", [])
        if summary_lines:
            resume_html.append("<div class='resume-section summary'>")
            resume_html.append("<h3>Professional Summary</h3>")
            for sl in summary_lines:
                resume_html.append(f"<p class='lede-text'>{sl}</p>")
            resume_html.append("</div>")
            
        # Proof of Work
        pow_lines = parsed["sections"].get("PROOF OF WORK", [])
        if pow_lines:
            resume_html.append(render_section_html("Proof of Work & Verification", pow_lines))
            
        # Skills (from Table)
        if parsed["skills"]:
            resume_html.append("<div class='resume-section skills-section'>")
            resume_html.append("<h3>Artificial Intelligence & Technical Skills</h3>")
            resume_html.append("<div class='skills-grid'>")
            for sk in parsed["skills"]:
                resume_html.append(f"<div class='skill-card'><span class='skill-cat'>{sk['category']}</span><p class='skill-det'>{sk['details']}</p></div>")
            resume_html.append("</div></div>")
            
        # Professional Experience
        exp_lines = parsed["sections"].get("PROFESSIONAL EXPERIENCE", [])
        if exp_lines:
            resume_html.append(render_section_html("Professional Experience", exp_lines))
            
        # By the Numbers
        btn_key = next((k for k in parsed["sections"] if "BY THE NUMBERS" in k), None)
        if btn_key:
            resume_html.append(render_section_html("Measured Repository Metrics", parsed["sections"][btn_key]))
            
        # Projects / Repos
        proj_key = next((k for k in parsed["sections"] if "PROJECTS" in k), None)
        if proj_key:
            resume_html.append(render_section_html("Key Systems & Repositories", parsed["sections"][proj_key]))
            
        # Program Directory / Verification
        prog_key = next((k for k in parsed["sections"] if "PROGRAM DIRECTORY" in k or "VERIFICATION" in k), None)
        if prog_key:
            resume_html.append(render_section_html("Package Registry & Reproducible Verification", parsed["sections"][prog_key]))

        active_class = "active" if cfg["default"] else ""
        
        tab_nav.append(f"""
        <button class="tab-btn {active_class}" data-target="{cfg['id']}">
            <span class="tab-title">{cfg['title']}</span>
            <span class="tab-badge">{cfg['badge']}</span>
        </button>
        """)
        
        rendered_tabs.append(f"""
        <article id="{cfg['id']}" class="resume-tab-pane {active_class}">
            <div class="resume-pane-header">
                <div>
                    <span class="pane-badge">{cfg['badge']}</span>
                    <h2 class="pane-title">{cfg['title']}</h2>
                    <p class="pane-subtitle">{cfg['subtitle']}</p>
                </div>
                <div class="pane-actions no-print">
                    <a class="btn-download" href="/downloads/{cfg['file']}" download>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                        Download Word (.docx)
                    </a>
                    <button class="btn-print" onclick="window.print()">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
                        Print / PDF
                    </button>
                </div>
            </div>
            
            <div class="resume-body">
                {''.join(resume_html)}
            </div>
        </article>
        """)

    index_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Alexander Sorrell — Interactive & Verifiable Resume</title>
    <meta name="description" content="AI Principal / Builder / Breaker — Verification & Evaluation. Interactive resume with verifiable telemetry, reproducible benchmarks, and ATS-clean export.">
    <meta property="og:title" content="Alexander Sorrell — Resume">
    <meta property="og:description" content="Models propose. A deterministic function owns the verdict. Every decision ships with a receipt a stranger can re-check.">
    <meta property="og:type" content="website">
    <link rel="icon" type="image/svg+xml" href="/assets/logo.svg">
    <link rel="stylesheet" href="/style.css">
</head>
<body>

    <!-- Header bar -->
    <header class="site-header no-print">
        <div class="container header-wrap">
            <a class="brand" href="https://bastionfold.com" target="_blank" rel="noopener">
                <img src="/assets/logo.svg" alt="Bastion Fold Logo" width="28" height="28">
                <span>Bastion<strong>Fold</strong></span>
            </a>
            <div class="header-links">
                <a href="https://www.linkedin.com/in/alexander-sorrell-it/" target="_blank" rel="noopener">LinkedIn</a>
                <a href="https://github.com/Alexander-Sorrell-IT" target="_blank" rel="noopener">GitHub</a>
                <a href="https://github.com/BastionFold/resume" target="_blank" rel="noopener">Resume Repo</a>
            </div>
        </div>
    </header>

    <!-- Main Container -->
    <main class="container">
        <!-- Profile Banner -->
        <section class="profile-card">
            <div class="profile-photo-wrap">
                <img class="profile-photo" src="/assets/alexander-sorrell.jpg" alt="Alexander Sorrell" width="160" height="213">
            </div>
            <div class="profile-info">
                <div class="profile-meta-top">
                    <span class="status-pill">AVAILABLE FOR ROLES</span>
                    <span class="audit-pill">★ 30.1 MASTERWORK · SYSTEMS ARCHITECT</span>
                </div>
                <h1 class="profile-name">Alexander Sorrell</h1>
                <p class="profile-headline">AI Principal / Builder / Breaker — Verification &amp; Evaluation</p>
                <p class="profile-thesis">
                    <strong>Central Thesis:</strong> Models propose. A deterministic function owns the verdict. Every decision ships with a receipt a stranger can re-check — a rule ID, a cryptographic hash chain, or the OS kernel itself.
                </p>
                <div class="contact-links no-print">
                    <a class="contact-item" href="https://www.linkedin.com/in/alexander-sorrell-it/" target="_blank" rel="noopener">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><path d="M19 0h-14c-2.761 0-5 2.239-5 5v14c0 2.761 2.239 5 5 5h14c2.762 0 5-2.239 5-5v-14c0-2.761-2.238-5-5-5zm-11 19h-3v-11h3v11zm-1.5-12.268c-.966 0-1.75-.79-1.75-1.764s.784-1.764 1.75-1.764 1.75.79 1.75 1.764-.783 1.764-1.75 1.764zm13.5 12.268h-3v-5.604c0-3.368-4-3.113-4 0v5.604h-3v-11h3v1.765c1.396-2.586 7-2.777 7 2.476v6.759z"/></svg>
                        LinkedIn
                    </a>
                    <a class="contact-item" href="https://github.com/Alexander-Sorrell-IT" target="_blank" rel="noopener">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><path d="M12 0c-6.626 0-12 5.373-12 12 0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23.957-.266 1.983-.399 3.003-.404 1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.921.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576 4.765-1.589 8.199-6.086 8.199-11.386 0-6.627-5.373-12-12-12z"/></svg>
                        Alexander-Sorrell-IT
                    </a>
                    <a class="contact-item" href="https://github.com/matrixbuilderops" target="_blank" rel="noopener">
                        matrixbuilderops
                    </a>
                    <a class="contact-item" href="mailto:alexander.sorrell.it@gmail.com">
                        ✉ alexander.sorrell.it@gmail.com
                    </a>
                    <span class="contact-item">✆ (817) 996-6123</span>
                    <span class="contact-item">📍 Colleyville, TX (Dallas–Fort Worth)</span>
                </div>
            </div>
        </section>

        <!-- Tabs Navigation -->
        <section class="tab-nav-section no-print">
            <div class="tab-nav-label">Select Tailored Dossier:</div>
            <div class="tabs-container">
                {''.join(tab_nav)}
            </div>
        </section>

        <!-- Tab Content Panes -->
        <section class="resume-panes-section">
            {''.join(rendered_tabs)}
        </section>
    </main>

    <footer class="site-footer no-print">
        <div class="container footer-wrap">
            <p>Bastion Fold LLC · SignalCore Inc · Sovereign local verification.</p>
            <p><a href="https://bastionfold.com" target="_blank">Bastion Fold Main Website</a> · <a href="https://github.com/Alexander-Sorrell-IT/starreckon" target="_blank">Starreckon Card</a></p>
        </div>
    </footer>

    <script>
        // Tab switching logic
        const buttons = document.querySelectorAll('.tab-btn');
        const panes = document.querySelectorAll('.resume-tab-pane');

        function switchTab(targetId) {{
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
                switchTab(b.dataset.target);
            }});
        }});

        // Handle direct URL hash (e.g. #matrixbuilderops)
        const hash = window.location.hash.replace('#', '');
        if (hash && document.getElementById(hash)) {{
            switchTab(hash);
        }}
    </script>
</body>
</html>"""

    with open(os.path.join(SITE_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)
        
    print(f"[build] Generated site with {len(RESUME_CONFIGS)} tabs into {SITE_DIR}")

if __name__ == "__main__":
    build_site()
