# Validation — methods and results

A technical, reproducible record of how the pipeline was checked against manual review,
and the resulting numbers. This document is deliberately limited to method and data;
interpretation of what these numbers mean belongs in the associated manuscript, not here.

The numbers below are from the current version of the pipeline. The accuracy validation
reported here has been completed for a single center (the first participating center of an
ongoing multicenter study); multicenter validation is in progress and is labeled as such
throughout. This repository contains no patient-level data, identifiers, or institution
names; every number is an aggregate count from a field-by-field comparison, and the source
discharge-summary PDFs remain in the local environment.

## How the comparison was done

- **Rules frozen before validation.** Parsing rules were finalized on patients outside the
  validation sample, so none of the scored data influenced rule development.
- **Blinded manual abstraction.** A clinical pharmacist abstracted every clinically relevant
  field for 20 randomly selected patients (5% of the 400-patient cohort) directly from the
  source PDFs, without reference to pipeline output.
- **Source-document adjudication.** Each disagreement was checked against the source PDF to
  assign the correct value, rather than assuming either side correct by default.
- **Statistics.** Sample size followed a single-proportion calculation (Cochran; 98%
  expected agreement, ±1% precision → ~750 fields required; 20 patients exceed this).
  Confidence intervals use the Wilson method.

## How to reproduce

```bash
# 1. Extract structured data from a folder of discharge-summary PDFs
python icu_pipeline.py --klasor path/to/pdf_folder --cikti automatic_output

# 2. Compare the automatic output against a manual abstraction, field by field
python validate.py --manuel manual_extraction --otomatik automatic_output --cikti concordance_report
```

The concordance tool emits the per-field match/mismatch counts and the discrepancy list
used below. Runtime is measured with `time.perf_counter` and logged per patient. Because
extraction is deterministic, re-running the same inputs reproduces the same numbers.

## Results

Paired fields: 8,428. After excluding fields empty in both outputs, **6,200** were compared.

| | Value |
|---|---|
| Raw concordance (pre-adjudication) | 97.3% (6,034/6,200; 95% CI 96.9–97.7) |
| Discrepancies | 166 (101 alignment artifacts + 65 genuine differences) |
| Adjudicated field-level accuracy | 99.7% (6,180/6,200; 95% CI 99.5–99.8) |

Per-domain adjudication of the 65 genuine differences:

| Domain | Genuine differences | Pipeline errors |
|---|---:|---:|
| Demographics | 5 | 0 |
| Clinical severity scores | 0 | 0 |
| Laboratory | 4 | 0 |
| Blood gas | 3 | 0 |
| Antimicrobials | 6 | 0 |
| Vasopressors | 2 | 0 |
| Microbiology | 45 | 20 |

Microbiology, adjudicated: 99.3% (95% CI 99.0–99.6). The 20 counted errors arise from one
pattern — antibiogram text placed in the culture-result field rather than split into
organism / antibiotic / susceptibility columns — plus one timestamp misread as a result.

<details>
<summary><b>Discrepancy accounting</b></summary>

<br>

| Category | Count |
|---|---:|
| Total discrepancies | 166 |
| — Alignment artifacts (not genuine differences) | 101 |
| — Genuine differences | 65 |
| Genuine differences adjudicated to the manual side | 25 |
| Genuine differences that are rounding/definition (no error either way) | 20 |
| Genuine differences counted as pipeline errors (all microbiology) | 20 |

</details>

### Comparison artifacts (adjusted before scoring)

Three classes of non-substantive difference were normalized before computing accuracy, to
avoid measuring the comparison method rather than the pipeline:

- **Turkish case-folding.** `str.casefold()` mishandles the dotted/dotless *i*, folding one
  value into two strings by case; this produced the bulk of the alignment artifacts. The
  fix is covered by unit tests in `tests/`.
- **Route labels.** Canonical `P` (parenteral) vs. the manual `IV` — treated as equivalent.
- **Numeric formatting.** A few laboratory differences are rounding only.

## Runtime

| | Value |
|---|---|
| Manual abstraction | median 26 min/patient |
| Pipeline | ~6 s/patient; full validation set in ~2 min |
| Reduction | ≈258× (standard hardware, no GPU) |

## External-center transportability (in progress)

The pipeline was run unmodified on a second center with a different document structure, as a
coverage check only (no manual reference exists there yet). Document-independent fields,
including culture results, were extracted; laboratory and blood-gas values were not, because
that center embeds them in free text; structured antibiogram tables are absent there, so
susceptibility was not present to extract. A small adaptation to that center's format
recovered date-anchored laboratory values. Formal multicenter accuracy validation is part of
the ongoing parent study.

## Limitations

- Single-center accuracy validation; multicenter accuracy not yet established.
- Reference standard is manual abstraction (single-reviewer; no inter-rater statistic).
- Residual pipeline error is confined to free-text microbiology parsing.
- Administration route is inferred from the drug's standard formulation, not parsed from the
  prescription (rare intramuscular use not distinguished; none in this cohort).
- Validation sample is 20 patients; patient-level clustering is handled in the parent analysis.

---

*All numbers come from a field-by-field comparison of pipeline output against an independent
manual review of the same discharge-summary PDFs, with source-PDF adjudication of each genuine
discrepancy. No patient-level data, identifiers, or institution names are stored in this
repository.*
