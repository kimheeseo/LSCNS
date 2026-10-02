# CPO Supply Chain

The CPO Supply Chain application is maintained in this repository under `cpo-supply-chain/`.

- Application: [GitHub Pages](https://kimheeseo.github.io/LSCNS/cpo-supply-chain/)
- Current release: **v2.10.0**
- Main entry: `index.html` loads the versioned `app.v2.10.0.payload`.

## Version log

### version: 2.10.0 — 2026-10-02

- Added a first-visit PC/mobile screen choice and saved the selected layout for later visits.
- Added country labels to company names in the supplier explorer.

For each future user-requested application improvement that changes the release version, append a dated `### version: x.y.z — YYYY-MM-DD` entry here with the concrete changes. Keep the version in the launcher and payload filename aligned with the latest release.

## Newly added companies

Record each newly added supplier as one row in [`list/added_companies.csv`](list/added_companies.csv). Include its addition date, name, country, CPO component/category, website, source, release version, and notes. This file is an additions log; do not re-list existing suppliers unless correcting their record.
