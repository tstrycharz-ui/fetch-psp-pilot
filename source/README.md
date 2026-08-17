# Source

This folder holds the Markdown knowledge base that the generated site (everything else in this repo — `psp/`, `best-friends-standard.html`, etc.) is built from. The HTML is a one-way build output of these files; edits should go here, not into the generated HTML directly.

- `company-defaults.md` — Best Friends company-wide policy standard
- `brand-knowledge/paw-sweet-paw-brand.md` — Paw Sweet Paw brand-level facts shared by both locations
- `location-knowledge/246-paw-sweet-paw-jamboree.md`, `247-paw-sweet-paw-spectrum.md` — per-location knowledge, organized around the six-tab dashboard schema (Center Information, Boarding, Daycare, Training, Grooming, Policies & Procedures)
- `open-questions.md` — running list of facts that are missing, ambiguous, or deferred, to be resolved with the Center Manager / District Manager
- `generate_site.py` — the generator; re-run after any Markdown edit to rebuild the site (`python3 generate_site.py` from this folder, with the sibling knowledge files present, then copy the output to the repo root)

Full working project (screenshots, PDFs, other brands, in-progress work) lives outside this repo; this folder is the PSP pilot's published subset.
