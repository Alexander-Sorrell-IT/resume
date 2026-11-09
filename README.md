# Alexander Sorrell — Resume Repository & Cloudflare Pages Site

Private repository under `BastionFold` housing the canonical resume `.docx` sources and automated static site builder for **Cloudflare Pages**.

## Resume Profiles & Tabs

1. **All Public (`Alexander_Sorrell_Resume_Public_All.docx`)**
   - **38 public repos · 170k lines · 3,342 tests**
   - Consolidated public view across both accounts.
2. **Alexander-Sorrell-IT (`Alexander_Sorrell_Resume_Alexander_Sorrell_IT.docx`)**
   - **33 repos · 156k lines · 3,168 tests**
   - Focus on core verification architectures, Starreckon, and OmniCLI.
3. **matrixbuilderops (`Alexander_Sorrell_Resume_matrixbuilderops.docx`)**
   - **5 repos · 13.8k lines · 174 tests**
   - Focus on autonomous agent fleets, Endless Watch, and runtime systems.
4. **All Repos Master Dossier (`Alexander_Sorrell_Resume_All_Repos.docx`)**
   - **119 repos · 786k lines · 6,150 tests**
   - Full masterwork dossier (public + private).

## How Dynamic Updates Work

When you update or add any `.docx` file in `resumes/` and push to `main`:
1. The GitHub Action runs `python build.py`.
2. Re-parses all `.docx` files into clean, ATS-compliant HTML with tabbed navigation.
3. Deploys live to Cloudflare Pages.
4. Triggers an automatic rebuild of the main Bastion Fold company site via `repository_dispatch`.

## Local Preview

```bash
cd /media/phantomcore/AI_DRIVE/websites/resume
python3 build.py
python3 -m http.server 8085 --directory _site
```
Open [http://localhost:8085](http://localhost:8085) to view.
