#!/usr/bin/env python3
"""Publish freshly generated resumes to every site.

build_master_resume.py (repo-dossier) writes four resumes. Three are public
and belong on the sites; the fourth, "All Repos", covers private repositories
and must never reach a website. This script moves the public three into the
resume repository, rebuilds the static site, syncs the Bastion Fold copy, and
refuses to publish anything that names an account outside OWNED_ACCOUNTS.

    python3 publish_resumes.py --from ~/Desktop/job-hunt/resumes/current
    python3 publish_resumes.py --from <dir> --no-push     # stage only
"""

import argparse
import os
import re
import shutil
import subprocess
import sys

import docx

ROOT = os.path.dirname(os.path.abspath(__file__))
BASTION = os.path.normpath(os.path.join(ROOT, "..", "bastion-fold"))
BASTION_DOWNLOADS = os.path.join(BASTION, "src", "assets", "downloads")
DOCS_DIR = os.path.join(ROOT, "docs")
SITE_DIR = os.path.join(ROOT, "_site")

OWNED_ACCOUNTS = {"Alexander-Sorrell-IT", "matrixbuilderops"}

# Published on every site.
PUBLIC_RESUMES = [
    "Alexander Sorrell Resume — Public",
    "Alexander Sorrell Resume — Public (Alexander-Sorrell-IT)",
    "Alexander Sorrell Resume — Public (matrixbuilderops)",
]

# Covers private repositories. Stays on disk, never published, never committed.
PRIVATE_RESUME = "Alexander Sorrell Resume — All Repos"

# Legacy underscore filenames the Bastion Fold site still carries.
BASTION_ALIASES = {
    "Alexander Sorrell Resume — Public": "Alexander_Sorrell_Resume_Public_All.docx",
    "Alexander Sorrell Resume — Public (Alexander-Sorrell-IT)":
        "Alexander_Sorrell_Resume_Alexander_Sorrell_IT.docx",
    "Alexander Sorrell Resume — Public (matrixbuilderops)":
        "Alexander_Sorrell_Resume_matrixbuilderops.docx",
}

OWNER_RX = re.compile(r"github\.com/([A-Za-z0-9_.-]+)")


def docx_text(path):
    """Every paragraph and table cell of a .docx, as one string."""
    document = docx.Document(path)
    parts = [p.text for p in document.paragraphs]
    parts += [c.text for t in document.tables for r in t.rows for c in r.cells]
    return "\n".join(parts)


def audit(path):
    """Return the GitHub owners named in a resume that are not ours."""
    owners = set(OWNER_RX.findall(docx_text(path)))
    return sorted(owners - OWNED_ACCOUNTS)


def run(cmd, cwd):
    subprocess.run(cmd, cwd=cwd, check=True)


def stage(source_dir):
    """Copy the three public resumes in, after checking each one."""
    staged = []
    for name in PUBLIC_RESUMES:
        for ext in (".docx", ".pdf"):
            src = os.path.join(source_dir, name + ext)
            if not os.path.exists(src):
                raise SystemExit(f"missing generated file: {src}")

        foreign = audit(os.path.join(source_dir, name + ".docx"))
        if foreign:
            raise SystemExit(
                f"refusing to publish {name}.docx: names accounts that are not "
                f"yours: {', '.join(foreign)}. Fix the generator's "
                f"_OUTSIDE_CONTRIBUTIONS step, not the .docx."
            )

        for ext in (".docx", ".pdf"):
            shutil.copy2(os.path.join(source_dir, name + ext),
                         os.path.join(ROOT, name + ext))
        staged.append(name)
        print(f"[stage] {name}")

    private = os.path.join(source_dir, PRIVATE_RESUME + ".docx")
    if os.path.exists(private):
        print(f"[skip ] {PRIVATE_RESUME} — private, never published")
    return staged


def sync_bastion():
    if not os.path.isdir(BASTION_DOWNLOADS):
        print(f"[warn ] no Bastion Fold checkout at {BASTION}; skipping")
        return False
    for name in PUBLIC_RESUMES:
        for ext in (".docx", ".pdf"):
            shutil.copy2(os.path.join(ROOT, name + ext),
                         os.path.join(BASTION_DOWNLOADS, name + ext))
        shutil.copy2(os.path.join(ROOT, name + ".docx"),
                     os.path.join(BASTION_DOWNLOADS, BASTION_ALIASES[name]))
    print("[sync ] Bastion Fold downloads")
    return True


def verify_published():
    """Nothing served anywhere may name an outside account or the private file."""
    for base in (SITE_DIR, DOCS_DIR, BASTION_DOWNLOADS):
        if not os.path.isdir(base):
            continue
        for folder, _, files in os.walk(base):
            for name in files:
                if PRIVATE_RESUME in name:
                    raise SystemExit(
                        f"private resume reached a published folder: "
                        f"{os.path.join(folder, name)}"
                    )
                if name.endswith(".docx"):
                    foreign = audit(os.path.join(folder, name))
                    if foreign:
                        raise SystemExit(
                            f"{os.path.join(folder, name)} names outside "
                            f"accounts: {', '.join(foreign)}"
                        )
    print("[check] no outside accounts, no private resume in any published folder")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--from", dest="source", required=True,
                        help="folder build_master_resume.py wrote the resumes to")
    parser.add_argument("--no-push", action="store_true",
                        help="stage and build, but do not commit or push")
    args = parser.parse_args()

    source_dir = os.path.expanduser(args.source)
    if not os.path.isdir(source_dir):
        raise SystemExit(f"no such folder: {source_dir}")

    stage(source_dir)

    run([sys.executable, "build.py"], ROOT)
    run(["rsync", "-a", "--delete", SITE_DIR + "/", DOCS_DIR + "/"], ROOT)
    open(os.path.join(DOCS_DIR, ".nojekyll"), "a").close()
    print("[build] site rebuilt and docs/ synced")

    have_bastion = sync_bastion()
    if have_bastion:
        run(["npm", "run", "build"], BASTION)

    verify_published()

    if args.no_push:
        print("[done ] staged only (--no-push)")
        return

    message = "Publish regenerated resumes\n\nGenerated by repo-dossier; published by publish_resumes.py."
    for repo, remotes in ((ROOT, ["origin", "personal"]), (BASTION, ["origin"])):
        subprocess.run(["git", "add", "-A"], cwd=repo, check=True)
        if subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=repo).returncode == 0:
            print(f"[push ] {os.path.basename(repo)}: nothing changed")
            continue
        run(["git", "commit", "-q", "-m", message], repo)
        for remote in remotes:
            run(["git", "push", remote, "main"], repo)
        print(f"[push ] {os.path.basename(repo)}: pushed to {', '.join(remotes)}")


if __name__ == "__main__":
    main()
