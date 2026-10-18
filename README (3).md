# A Locally Run, Rule-Based Pipeline for Structuring ICU Records

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.23461019.svg)](https://doi.org/10.5281/zenodo.23461019)
![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Python](https://img.shields.io/badge/python-3.12%2B-blue.svg)
![Tests](https://img.shields.io/badge/tests-passing-brightgreen.svg)
![Status: research preview](https://img.shields.io/badge/status-research%20preview-orange.svg)

*A fully local and reproducible extraction pipeline, validated field by field against manual review.*

This is a local, rule-based pipeline that converts unstructured Turkish intensive care
discharge summary PDFs into a structured, analysis-ready dataset. It does not use a large
language model, neural network, or external service. Extraction is fully deterministic, and
every output value can be traced to the rule that produced it.

The pipeline is the first component of a broader effort to replace slow, labor-intensive and error-prone manual chart review
with automated extraction that is auditable, reproducible,
privacy-preserving, and validated field-by-field against manual review in a non-English
clinical setting.

## Design rationale

Automating clinical data collection increasingly means applying large language models, but in
a safety-critical, privacy-sensitive workflow those carry concrete costs: fabricated or
dropped values, non-reproducible output, and, for hosted models, transmission of sensitive
records to third parties. A deterministic, rule-based design prioritizes three properties that
matter more in this context over broad generalization:

- **Auditability:** each extracted value maps back to a specific rule.
- **Reproducibility:** identical input yields identical output every time.
- **Privacy:** processing is entirely local, so no data leaves the institution.

Rule-based extraction is not universally preferable, because generative models handle heterogeneous
free text in ways rules cannot. For a defined extraction task over reasonably stable
report formats, on data that must remain on-site, it is a sound fit.

## Validation

Validated field-by-field against manual abstraction of the same PDFs in 20 patients, with
every disagreement adjudicated against the source document (6,200 fields):

| Metric | Result |
|---|---|
| Raw agreement (before adjudication) | 97.3% (95% CI 96.9–97.7) |
| Adjudicated field-level accuracy | **99.1%** (95% CI 98.9–99.3); worst case 97.7% |
| Domains with zero pipeline errors | 5 of 7 (errors only in free-text microbiology and one antimicrobial course) |
| Intra-rater agreement of the manual reference | 99.7% (3,279/3,288 field pairs) |
| Median processing time | 8.8 s/patient vs. 26 min manual (≈180× faster) |

Method, how to reproduce, per-domain adjudication, and limitations are in
**[`docs/VALIDATION.md`](docs/VALIDATION.md)**. The engineering and design rationale (why
rule-based, the architecture, and the Turkish-specific handling) is in
**[`DESIGN.md`](DESIGN.md)**.

## What it extracts

Seven structured domains, drawn from four routine report types:

| Report type | Domains extracted |
|---|---|
| Discharge summary (E1) | Demographics · clinical severity scores · antimicrobial courses · vasopressor courses · blood gas |
| Daily ICU form (E2) | Clinical severity scores (APACHE II, SOFA, GCS) · sepsis / septic-shock flags |
| Laboratory report (LAB) | Laboratory values |
| Microbiology report (M) | Culture results · antibiogram / susceptibility |

<details>
<summary><b>Field-level detail (click to expand)</b></summary>

<br>

- **Demographics**: age, sex, admission/discharge dates, length of stay, disposition, ICD codes
- **Clinical severity**: APACHE II, SOFA, GCS (eye/verbal/motor), sepsis and septic-shock flags
- **Laboratory**: routine chemistry and hematology values
- **Blood gas**: pH, pCO₂, pO₂, HCO₃, base excess, SpO₂, electrolytes
- **Antimicrobials**: agent, dose, route, frequency, start/stop, duration, with ATC / WHO AWaRe / DDD lookups
- **Vasopressors**: agent, start/stop, duration
- **Microbiology**: specimen, organism, culture result, antibiotic, susceptibility

</details>

## Getting started

Python 3.12 or later (the study runs used Python 3.14.1). Pin a single interpreter version for reproducibility.

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Extract structured data from a folder of PDFs:

```bash
python icu_pipeline.py --klasor path/to/pdf_folder --cikti output_name
```

Run the field-by-field concordance analysis (manual vs. automatic) and the intra-rater check:

```bash
python validate.py --manuel manual_extraction.xlsx --otomatik automatic_extraction.xlsx --cikti report.xlsx
python analysis/intra_rater.py first_abstraction.xlsx second_abstraction.xlsx
```

`icu_pipeline_h2.py` is the version adapted to the second center's report format (see
`docs/VALIDATION.md`). `tools/amr_mapper.py` builds the drug → ATC → AWaRe → DDD reference table
from the public TİTCK product list and the WHO AWaRe 2023 list (not redistributed here), and
`tools/pdf_rename.py` renames raw report exports to the `{id}_{E1|E2|LAB|M1..M9}.pdf` convention.

## Try it on synthetic data

The `examples/` folder contains three fully synthetic patients, invented throughout and containing no
real data, so the pipeline can be run end-to-end immediately:

```bash
python icu_pipeline.py --klasor examples --cikti demo_output
python -m pytest -q          # run the unit tests
```

The synthetic PDFs are generated by `examples/generate_synthetic.py`, which can be inspected
or edited to produce different cases.

## Roadmap

Actively developed. The pipeline was built to support an ongoing, large-scale multicenter
study; this repository covers the first site. Planned work:

- **Multicenter external validation**: at a second site with a largely free-text format, the
  unmodified pipeline covered 7.1% of reference fields and an adapted version 33.5% (clinical
  scores and laboratory values near-complete; blood gas and medication data not transferable;
  see `docs/VALIDATION.md`). Independent validation at further sites is planned.
- **More robust free-text microbiology parsing**: the source of the few residual errors.
- **Linked references** from results tables to the corresponding manuscript sections (a
  future revision).

## Privacy

This repository contains only source code and synthetic examples; it does not contain
patient-level data, identifiers, or institution names. The pipeline is designed to run locally
within the institution so that records never leave it. Real PDFs and any extracted datasets are
excluded from version control via `.gitignore`. Health data in the source study was processed
under institutional ethics approval and applicable data-protection law.

## Citation & license

See [`CITATION.cff`](CITATION.cff). Released under the MIT License.
