#!/usr/bin/env python3
"""Write live package numbers into the resumes, scoped to each resume's account.

package_telemetry.py is the only source of numbers. It is run once on the full
package list for the JSON, and once per scope with a filtered copy of its
config for the rendered appendix (metrics table, averages, program directory),
so the averages are the tool's own math, not a reimplementation.

Scope comes from the filename: "(matrixbuilderops)" and "(Alexander-Sorrell-IT)"
resumes carry only that account's packages; every other resume carries all.
In a scoped resume the other account's packages, its starreckon QR section and
its registry's verification commands are removed, the header QR is swapped for
the account's own, and every count, total and date is recomputed for the scope.

Only run text is rewritten and only whole paragraphs are added or removed;
run formatting is never touched. A figure that cannot be tied to a live number,
or a package name outside the resume's scope, is an error, not a skip.

    python3 refresh_telemetry.py --tool <package-stats dir> *.docx
"""

import argparse
import copy
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import urllib.request

import docx
from docx.text.paragraph import Paragraph

IT = "Alexander-Sorrell-IT"
MBO = "matrixbuilderops"

# Which account's repository publishes each package. A package the tool reports
# that is missing here stops the run, so a new package cannot slip into the
# wrong account's resume.
PACKAGE_ACCOUNT = {
    "cli-enforcement": IT, "cli-wikia": IT, "cli-collective": IT, "cli-fleet": IT,
    "starforge-cli": IT, "starreckon": IT,
    "sitemap-mcp": MBO, "filelens-mcp": MBO, "verifier-mcp": MBO,
    "sorrell-engine-mcp": MBO, "package-telemetry": MBO,
}
# Names that identify one account's work without being a package.
# starforge: former name of Alexander-Sorrell-IT/starreckon (GitHub redirects it).
OTHER_NAMES = {"starforge": IT}
QR_PACKAGE = "starreckon"
# Header QR for a scope whose resume must not carry the starreckon audit card.
SCOPE_QR = {MBO: ("github.com/matrixbuilderops",
                  os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "matrixbuilderops-qr.png"))}
REGISTRY_SCOPE = {"PyPI and npm": "PyPI & npm live registries", "PyPI": "PyPI live registry",
                  "npm": "npm live registry"}
COUNT_WORDS = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven"}

ANSI = re.compile(r"\x1b\[[0-9;]*m")
EMPTY_REGISTRY = re.compile(r"\u2502\s+0 pkgs \u2502")
DOWNLOADS_AFTER_NAME = re.compile(r"^(\s*)([\d,]+)( downloads\b)")
NPM_HEADLINE = re.compile(r"([\d,]+)( npm downloads \((?:npx )?([a-z0-9-]+)\))")
SPLIT_HEADLINE = re.compile(r"([\d,]+)( npm downloads \()$")  # package name is in the next run
ON_PYPI = re.compile(r"(on PyPI \u2014 )([\d,]+)( downloads)")  # figure for the entry's own package
ANY_DOWNLOADS = re.compile(r"([\d,]+) (?:npm |PyPI )?downloads\b")
PROJECT = re.compile(r"^[A-Za-z0-9][\w.-]* \u2014 ")
LIFETIME = re.compile(r"^[\d,]+( lifetime downloads)")
COUNTS = re.compile(r"across \d+ published packages and \d+ releases on (?:PyPI and npm|PyPI|npm)")
SHIPPED = re.compile(r"Shipped \d+ packages")
SUBTOTAL = re.compile(r"^(PyPI|npm) \([\d,]+\)(:?)$")
K_FIGURE = re.compile(r"^ \(\d+\.\dk\)([.,] ?)$")
LIVE_DATE = re.compile(r"\(live, \d{4}-\d{2}-\d{2}\)")
EXECUTED = re.compile(r"Executed: \d{4}-\d{2}-\d{2}")
REG_SCOPE = re.compile(r"Scope: (?:PyPI & npm live registries|PyPI live registry|npm live registry)")
WITHOUT_MIRRORS = re.compile(r"the \w+ PyPI packages? totals? [\d,]+")

IPV4_SHIM = (
    "import socket,sys,runpy\n"
    "o=socket.getaddrinfo\n"
    "socket.getaddrinfo=lambda h,p,f=0,t=0,pr=0,fl=0:o(h,p,socket.AF_INET,t,pr,fl)\n"
    "sys.argv=sys.argv[1:]\n"
    "runpy.run_path(sys.argv[0],run_name='__main__')\n"
)


def ipv4_only():
    """Resolve names over IPv4 only, in this process.

    Hosts with a blackholed IPv6 route make urllib, which has no happy-eyeballs
    fallback, wait out a full timeout on every request. Where IPv6 works this
    costs nothing.
    """
    original = socket.getaddrinfo
    socket.getaddrinfo = lambda h, p, f=0, t=0, pr=0, fl=0: original(h, p, socket.AF_INET, t, pr, fl)


def fmt(n):
    return f"{n:,}"


def scope_of(path):
    name = os.path.basename(path)
    if f"({MBO})" in name:
        return MBO
    if f"({IT})" in name:
        return IT
    return None


def run_tool(tool_dir, packages, *flags):
    """Run package_telemetry.py from a copy whose config lists only `packages`."""
    with tempfile.TemporaryDirectory() as tmp:
        work = os.path.join(tmp, "package-stats")
        shutil.copytree(tool_dir, work, symlinks=True)
        config_path = os.path.join(work, "packages_config.json")
        with open(config_path, encoding="utf-8") as handle:
            config = json.load(handle)
        for registry in ("pypi", "npm"):
            config[registry] = [p for p in config[registry] if p in packages]
        with open(config_path, "w", encoding="utf-8") as handle:
            json.dump(config, handle, indent=2)
        result = subprocess.run(
            [sys.executable, "-c", IPV4_SHIM, "package_telemetry.py", "--no-save", *flags],
            cwd=work, check=True, capture_output=True, text=True, timeout=600,
        )
    return result.stdout


def appendix(tool_dir, packages):
    """The tool's --detailed rendering, minus colour and its local Config: path."""
    lines = [ANSI.sub("", line) for line in run_tool(tool_dir, packages, "--detailed").splitlines()]
    lines = [line for line in lines if not line.startswith("Config:")]
    lines = [line for line in lines if not EMPTY_REGISTRY.search(line)]  # e.g. PyPI row for an npm-only scope
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    width = len(lines[3])
    for i, line in enumerate(lines):
        if len(line) > width and "│" in line:  # a package name overflowed its column
            lines[i] = re.sub(r"(│ \S+) │", r"\1│", line, count=1)
    return lines


def without_mirrors(packages):
    total = 0
    for name in packages:
        url = f"https://pypistats.org/api/packages/{name}/overall"
        request = urllib.request.Request(url, headers={"User-Agent": "refresh_telemetry"})
        with urllib.request.urlopen(request, timeout=30) as response:
            rows = json.load(response)["data"]
        total += sum(r["downloads"] for r in rows if r["category"] == "without_mirrors")
    return total


class Scope:
    """The packages, totals and rendered appendix one resume should show."""

    def __init__(self, account, data, tool_dir):
        self.account = account
        everything = data["pypi"] + data["npm"]
        unknown = [p["package"] for p in everything if p["package"] not in PACKAGE_ACCOUNT]
        if unknown:
            raise SystemExit(f"no account recorded for package(s): {', '.join(unknown)}")
        keep = lambda p: account is None or PACKAGE_ACCOUNT[p["package"]] == account  # noqa: E731
        self.pypi = [p for p in data["pypi"] if keep(p)]
        self.npm = [p for p in data["npm"] if keep(p)]
        self.packages = {p["package"]: p for p in self.pypi + self.npm}
        self.excluded = sorted(n for n, a in {**PACKAGE_ACCOUNT, **OTHER_NAMES}.items()
                               if account is not None and a != account)
        self.lifetime = sum(p["downloads_lifetime"] for p in self.packages.values())
        self.releases = sum(p["total_releases"] for p in self.packages.values())
        self.date = data["timestamp"][:10]
        self.registries = " and ".join(r for r, group in (("PyPI", self.pypi), ("npm", self.npm)) if group)
        self.appendix = appendix(tool_dir, set(self.packages))
        self.pypi_without_mirrors = without_mirrors([p["package"] for p in self.pypi]) if self.pypi else 0

    def subtotal(self, registry):
        group = self.pypi if registry == "PyPI" else self.npm
        return sum(p["downloads_lifetime"] for p in group)


def delete(paragraph):
    paragraph._p.getparent().remove(paragraph._p)


def find(document, predicate):
    return [p for p in document.paragraphs if predicate(p.text)]


def rewrite_runs(document, scope):
    edits = 0
    pkgs = scope.packages
    for paragraph in document.paragraphs:
        runs = paragraph.runs
        project = paragraph.text.split(" \u2014 ", 1)[0].strip().lower() if PROJECT.match(paragraph.text) else None
        for i, run in enumerate(runs):
            name = run.text.strip()
            if name in pkgs and i + 1 < len(runs):
                nxt = runs[i + 1]
                m = DOWNLOADS_AFTER_NAME.match(nxt.text)
                if m:
                    nxt.text = m.group(1) + fmt(pkgs[name]["downloads_lifetime"]) + nxt.text[m.end(2):]
                    edits += 1

            def npm_headline(m):
                pkg = pkgs.get(m.group(3))
                if pkg is None:
                    return m.group(0)  # another account's package: verify() reports it
                return fmt(pkg["downloads_lifetime"]) + m.group(2)

            text = NPM_HEADLINE.sub(npm_headline, run.text)
            split = SPLIT_HEADLINE.search(text)
            if split and i + 1 < len(runs):
                follower = re.match(r"(?:npx )?([a-z0-9-]+)", runs[i + 1].text)
                if follower and follower.group(1) in pkgs:
                    figure = fmt(pkgs[follower.group(1)]["downloads_lifetime"])
                    text = text[:split.start()] + figure + split.group(2)
            if ON_PYPI.search(text):
                if project not in pkgs:
                    raise SystemExit(f"'on PyPI' figure in entry {project!r} has no matching package")
                text = ON_PYPI.sub(lambda m: m.group(1) + fmt(pkgs[project]["downloads_lifetime"]) + m.group(3), text)
            text = LIFETIME.sub(lambda m: fmt(scope.lifetime) + m.group(1), text)
            text = COUNTS.sub(f"across {len(pkgs)} published packages and {scope.releases} releases "
                              f"on {scope.registries}", text)
            text = SHIPPED.sub(f"Shipped {len(pkgs)} packages", text)
            text = LIVE_DATE.sub(f"(live, {scope.date})", text)
            text = EXECUTED.sub(f"Executed: {scope.date}", text)
            text = REG_SCOPE.sub("Scope: " + REGISTRY_SCOPE[scope.registries], text)
            if WITHOUT_MIRRORS.search(text):
                n = len(scope.pypi)
                text = WITHOUT_MIRRORS.sub(
                    f"the {COUNT_WORDS[n]} PyPI package{'s' if n > 1 else ''} "
                    f"total{'' if n > 1 else 's'} {fmt(scope.pypi_without_mirrors)}", text)
            if text != run.text:
                run.text = text
                edits += 1
    return edits


def rewrite_breakdown(document, scope):
    """The '• PyPI (N): a (x.xk), b (y.yk).' and '• npm (...)' bullets."""
    edits = 0
    for registry, group in (("PyPI", scope.pypi), ("npm", scope.npm)):
        bullets = find(document, lambda t, r=registry: t.startswith(f"• {r} ("))
        if not group and not bullets:  # removed by a previous refresh
            continue
        if len(bullets) != 1:
            raise SystemExit(f"expected one '{registry}' breakdown bullet, found {len(bullets)}")
        bullet = bullets[0]
        if not group:
            delete(bullet)
            edits += 1
            continue
        runs = bullet.runs
        label = next(r for r in runs if SUBTOTAL.match(r.text.strip()))
        colon = SUBTOTAL.match(label.text.strip()).group(2)
        label.text = f"{registry} ({fmt(scope.subtotal(registry))}){colon}"
        slots = [(runs[i], runs[i + 1]) for i in range(len(runs) - 1)
                 if runs[i].text.strip() in PACKAGE_ACCOUNT and K_FIGURE.match(runs[i + 1].text)]
        ranked = sorted(group, key=lambda p: p["downloads_lifetime"], reverse=True)
        if len(ranked) > len(slots):
            raise SystemExit(f"{registry} bullet has {len(slots)} slots for {len(ranked)} packages")
        for (name_run, figure_run), pkg in zip(slots, ranked):
            name_run.text = pkg["package"]
            figure_run.text = f" ({pkg['downloads_lifetime'] / 1000:.1f}k), "
        for name_run, figure_run in slots[len(ranked):]:
            for run in (name_run, figure_run):
                run._r.getparent().remove(run._r)
        slots[len(ranked) - 1][1].text = slots[len(ranked) - 1][1].text[:-2] + "."
        edits += 1
    return edits


def rewrite_appendix(document, scope):
    """Replace the metrics/averages/directory block with the scope's rendering."""
    paragraphs = document.paragraphs
    starts = [i for i, p in enumerate(paragraphs) if "PACKAGE METRICS & TELEMETRY" in p.text]
    ends = [i for i, p in enumerate(paragraphs) if p.text.startswith("Registry Ledger")]
    if len(starts) != 1 or len(ends) != 1:
        raise SystemExit("expected exactly one telemetry appendix")
    block = paragraphs[starts[0] - 1:ends[0]]
    for paragraph in block:
        if len(paragraph.runs) > 1:
            raise SystemExit(f"appendix line is not a single run: {paragraph.text[:40]!r}")
    template = copy.deepcopy(block[1]._p)  # the title line; every block line shares its formatting
    anchor = block[0]._p.getprevious()
    for paragraph in block:
        delete(paragraph)
    for line in scope.appendix:
        new = copy.deepcopy(template)
        anchor.addnext(new)
        anchor = new
        paragraph = Paragraph(new, document._body)
        paragraph.runs[0].text = line
    return len(scope.appendix)


def remove_section(document, heading_prefix, stop_prefix):
    """Delete the paragraphs from a heading up to (not including) the next one.

    Already gone is the goal state (a previous refresh removed it); verify()
    checks the result either way.
    """
    paragraphs = document.paragraphs
    starts = [i for i, p in enumerate(paragraphs) if p.text.strip().startswith(heading_prefix)]
    if not starts:
        return 0
    if len(starts) != 1:
        raise SystemExit(f"expected one '{heading_prefix}' section, found {len(starts)}")
    i = starts[0]
    removed = 0
    while i < len(paragraphs) and not paragraphs[i].text.strip().startswith(stop_prefix):
        delete(paragraphs[i])
        removed += 1
        i += 1
    return removed


def rewrite_verification(document, scope):
    """Keep only the scope's registries and packages in the curl examples."""
    edits = 0
    if not scope.pypi:
        edits += remove_section(document, "# 1. PyPI Verification", "# 2. npm Verification")
        for paragraph in find(document, lambda t: t.startswith("# 2. npm Verification")):
            paragraph.runs[0].text = paragraph.runs[0].text.replace("# 2.", "# 1.", 1)
            edits += 1
    ranked = [p["package"] for p in sorted(scope.npm, key=lambda p: p["downloads_lifetime"], reverse=True)]
    curl = [p for p in document.paragraphs if p.text.startswith("curl -s https://api.npmjs.org/")]
    used = {m for p in curl for m in re.findall(r"/([a-z0-9-]+) \| jq", p.text) if m in scope.packages}
    spare = [p for p in ranked if p not in used]
    for paragraph in curl:
        run = paragraph.runs[0]
        m = re.search(r"/([a-z0-9-]+) \| jq", run.text)
        if m and m.group(1) not in scope.packages:
            if not spare:
                delete(paragraph)
            else:
                run.text = run.text.replace(f"/{m.group(1)} | jq", f"/{spare.pop(0)} | jq")
            edits += 1
    return edits


def verify(document, path, scope):
    cells = [c.text for t in document.tables for r in t.rows for c in r.cells]
    text = "\n".join([p.text for p in document.paragraphs] + cells)
    for name in scope.excluded:
        if re.search(r"(?<![\w-])" + re.escape(name) + r"(?![\w-])", text):
            line = next(t for t in text.splitlines() if re.search(r"(?<![\w-])" + re.escape(name) + r"(?![\w-])", t))
            raise SystemExit(f"{path}: names {name}, which is not this resume's package: {line[:90]!r}")
    if scope.account == MBO and QR_PACKAGE in text:
        raise SystemExit(f"{path}: still mentions {QR_PACKAGE}")
    live = {fmt(p["downloads_lifetime"]) for p in scope.packages.values()}
    for paragraph in document.paragraphs:
        runs = paragraph.runs
        for i, run in enumerate(runs):
            m = DOWNLOADS_AFTER_NAME.match(run.text)
            if m and m.group(2) not in live:
                raise SystemExit(f"{path}: stale '{m.group(2)} downloads' after {runs[i - 1].text.strip()!r}")
        t = paragraph.text
        for figure in ANY_DOWNLOADS.findall(t):
            if figure not in live and figure != fmt(scope.lifetime):
                raise SystemExit(f"{path}: stale '{figure} downloads' in {t[:70]!r}")
        if "lifetime downloads" in t and f"{fmt(scope.lifetime)} lifetime downloads" not in t:
            raise SystemExit(f"{path}: stale lifetime total in {t[:60]!r}")
        for n in re.findall(r"(\d+) (?:published )?packages\b", t):
            if int(n) != len(scope.packages):
                raise SystemExit(f"{path}: says {n} packages, scope has {len(scope.packages)}: {t[:60]!r}")


def drop_foreign_projects(document, scope):
    """In a scoped resume, drop project entries and bullets about only the other account's work.

    The PyPI / npm breakdown bullets are rewritten by rewrite_breakdown instead.
    """
    dropped = 0
    for paragraph in document.paragraphs:
        text = paragraph.text
        bullet = text.startswith("\u2022 ") and not text.startswith(("\u2022 PyPI (", "\u2022 npm ("))
        if not (PROJECT.match(text) or bullet):
            continue
        named = set(re.findall(r"[a-z0-9]+(?:-[a-z0-9]+)+|[a-z0-9]+", paragraph.text.lower()))
        if named & set(scope.excluded) and not named & set(scope.packages):
            delete(paragraph)
            dropped += 1
    return dropped


def rescope_header_qr(document, scope):
    """Swap the header QR and its caption for the scope's own, when it has one."""
    if scope.account not in SCOPE_QR:
        return 0
    caption, image = SCOPE_QR[scope.account]
    qr_cells = [c for t in document.tables for r in t.rows for c in r.cells if c._tc.xpath(".//pic:pic")]
    if any(caption in c.text and QR_PACKAGE not in c.text for c in qr_cells):
        return 0  # swapped by a previous refresh
    cells = [c for c in qr_cells if QR_PACKAGE in c.text]
    if len(cells) != 1:
        raise SystemExit(f"expected one header QR cell naming {QR_PACKAGE}, found {len(cells)}")
    cell = cells[0]
    blips = cell._tc.xpath(".//a:blip/@r:embed")
    if len(blips) != 1:
        raise SystemExit(f"expected one image in the header QR cell, found {len(blips)}")
    with open(image, "rb") as handle:
        document.part.related_parts[blips[0]]._blob = handle.read()
    for paragraph in cell.paragraphs:
        if QR_PACKAGE in paragraph.text:  # "github.com/.../starreckon" <br> "npx starreckon"
            first, *rest = paragraph.runs
            first.text = caption
            for run in rest:
                run._r.getparent().remove(run._r)
        elif paragraph.text.strip() and not paragraph._p.xpath(".//pic:pic"):  # the audit score line
            for run in paragraph.runs:
                run.text = ""
    return 1


def refresh(path, scope):
    document = docx.Document(path)
    edits = drop_foreign_projects(document, scope) if scope.account else 0
    if scope.account and QR_PACKAGE not in scope.packages:
        remove_section(document, "QR CODE TARGET ENGINE", "PUBLIC INDEPENDENT VERIFICATION")
    edits += rescope_header_qr(document, scope)
    edits += rewrite_breakdown(document, scope)
    edits += rewrite_appendix(document, scope)
    edits += rewrite_verification(document, scope)
    edits += rewrite_runs(document, scope)
    verify(document, path, scope)
    document.save(path)
    label = scope.account or "all accounts"
    print(f"[telemetry] {os.path.basename(path)}: {label}, {len(scope.packages)} packages, "
          f"{fmt(scope.lifetime)} lifetime as of {scope.date}, {edits} edits")


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--tool", required=True, help="folder holding package_telemetry.py")
    parser.add_argument("docx", nargs="+")
    args = parser.parse_args()
    tool_dir = os.path.expanduser(args.tool)
    ipv4_only()

    data = json.loads(run_tool(tool_dir, set(PACKAGE_ACCOUNT), "--json"))
    if data.get("persist_blockers"):
        raise SystemExit(f"telemetry not trustworthy: {data['persist_blockers']}")
    stale = [p["package"] for p in data["pypi"] + data["npm"] if p.get("downloads_stale")]
    if stale:
        raise SystemExit(f"telemetry returned cached numbers for: {', '.join(stale)}")

    scopes = {}
    for path in args.docx:
        account = scope_of(path)
        if account not in scopes:
            scopes[account] = Scope(account, data, tool_dir)
        refresh(path, scopes[account])


if __name__ == "__main__":
    main()
