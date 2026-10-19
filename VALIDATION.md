# Validation: methods and results

A technical, reproducible record of how the pipeline was checked against manual abstraction,
and the resulting numbers. It is limited to method and data; interpretation belongs in the
associated manuscript. This repository contains no patient-level data, identifiers or
institution names; every number below is an aggregate count.

## Design

- **Rules frozen before validation.** Parsing rules were developed on patients outside the
  validation sample and were not changed on the basis of validation results.
- **Sample.** 20 randomly selected patients (5% of a 400-patient single-center cohort).
  Sample size followed a single-proportion calculation (Cochran; 98% expected agreement,
  ±1% precision → ~753 fields required).
- **Reference standard.** A clinical pharmacist abstracted all clinically relevant fields from
  the source PDFs without viewing pipeline output. The original, uncorrected abstraction is the
  reference. The same investigator developed the rules and adjudicated discrepancies, so the
  abstraction is not blinded to the extraction logic.
- **Excluded from scoring.** Derived fields (age, length of stay, the day number of laboratory
  and clinical rows, active ingredient, ATC, AWaRe, DDD), tool metadata (`Kaynak`) and MIC values.
  Start and end day numbers of antimicrobial and vasopressor courses were compared.
- **Adjudication.** Every discrepancy was checked against the source PDF and classified as a
  pipeline error, a manual-abstraction error, an equally valid day selection within the
  ±2-day window, a naming artifact (the same course recorded under a differently labeled agent
  name) or a rounding difference. Only pipeline errors count against the pipeline.
- **Statistics.** Wilson 95% CIs and patient-level cluster bootstrap (2,000 resamples,
  seed 2025). Fields empty in both outputs are reported separately and excluded from the
  denominator.

## Normalization before comparison

Applied identically to both outputs by `validate.py`:

- Unicode NFC and **Turkish-aware letter folding** (I, İ, ı, i → one letter). Python's default
  `casefold()` turns `İ` (U+0130) into `i` + U+0307, which produced **478 spurious mismatches**
  in microbiology text before this fix (microbiology agreement 81.8% → 98.1%). Covered by `tests/test_normalize.py`.
- Decimal comma → point; dates → ISO; route label `P` (parenteral) ≡ `IV`.
- **Duplicate-key merge.** Multi-page microbiology reports repeat the antibiogram table without
  the date header; merging duplicates (a populated value never overwrites an empty one) removed
  36 spurious date mismatches.
- **Key re-matching.** When the pipeline failed to extract a key field (organism or antibiotic),
  the unmatched record was re-matched within the same patient and report, and the failure was
  counted once as a pipeline error.

## How to reproduce

```bash
python icu_pipeline.py --klasor path/to/pdf_folder --cikti automatic_output.xlsx
python validate.py --manuel manual_extraction.xlsx --otomatik automatic_output.xlsx --cikti report.xlsx
python analysis/intra_rater.py first_abstraction.xlsx second_abstraction.xlsx
```

Extraction is deterministic: repeated runs on the same inputs produced identical output.
In the study's manual file, seven laboratory rows carried the pseudonymous patient ID instead of
the local record number; they were mapped to the same patients before comparison.

## Results: accuracy

Paired fields 8,428; empty in both 2,228; **evaluated 6,200**.

| Metric | Value |
|---|---|
| Raw agreement (before adjudication) | 97.3% (6,034/6,200; Wilson 96.9–97.7; bootstrap 95.7–98.4) |
| Adjudicated pipeline accuracy | **99.1%** (6,147/6,200; Wilson 98.9–99.3; bootstrap 98.8–99.4) |
| Worst case (manual errors and naming artifacts counted against the pipeline) | 97.7% (6,055/6,200) |
| Patients with ≥1 pipeline error | 13/20 |
| Patients with ≥1 manual-abstraction error | 12/20 |

| Domain | Fields | Raw agreement % (n) | 95% CI | Pipeline errors | Manual errors | Other* | Adjudicated % |
|---|---:|---:|---|---:|---:|---:|---:|
| Demographics | 262 | 98.1 (257) | 95.6–99.2 | 0 | 5 | 0 | 100 |
| Clinical severity scores | 316 | 94.9 (300) | 91.9–96.9 | 0 | 0 | 16 | 100 |
| Laboratory values | 685 | 99.4 (681) | 98.5–99.8 | 0 | 1 | 3 | 100 |
| Arterial blood gas | 1,606 | 99.3 (1,595) | 98.8–99.6 | 0 | 9 | 2 | 100 |
| Antimicrobial courses | 308 | 82.1 (253) | 77.5–86.0 | 7 | 34 | 14 | 97.7 |
| Vasopressor courses | 84 | 78.6 (66) | 68.7–86.0 | 0 | 18 | 0 | 100 |
| Microbiology | 2,939 | 98.1 (2,882) | 97.5–98.5 | 46 | 11 | 0 | 98.4 |
| **Total** | **6,200** | **97.3 (6,034)** | 96.9–97.7 | **53** | **78** | **35** | **99.1** |

*Other: equally valid day selections (16), naming artifacts (14) and rounding differences (5);
none counted as errors.

**Pipeline errors (53)** arose from three mechanisms: Gram-stain or antibiogram free text
written into the culture-result field (27 fields); organism or antibiotic name not parsed from
free-text microbiology reports (19 fields); one antimicrobial course not captured (7 fields).

## Results: intra-rater reliability of the reference standard

The same investigator re-abstracted the 20 patients from the source PDFs ten weeks later,
without access to the first abstraction or to pipeline output, creating the records (patients,
days, courses) anew under the same sampling rules. Records were paired by their order within each
domain. Microbiology was re-abstracted for one patient only and is excluded.

| Domain | Field pairs | Agreement % (n) | 95% CI |
|---|---:|---:|---|
| Demographics | 282 | 98.9 (279) | 96.9–99.6 |
| Clinical severity scores | 328 | 99.7 (327) | 98.3–99.9 |
| Laboratory values | 698 | 99.6 (695) | 98.7–99.9 |
| Arterial blood gas | 1,598 | 99.9 (1,596) | 99.5–100.0 |
| Antimicrobial courses | 287 | 100 (287) | 98.7–100.0 |
| Vasopressor courses | 95 | 100 (95) | 96.1–100.0 |
| **Total** | **3,288** | **99.7 (3,279)** | 99.5–99.9 |

## Results: runtime

| | Value |
|---|---|
| Manual abstraction (20 patients) | median 26 min/patient (IQR 22–33; range 13–72); 608 min total |
| Pipeline (same 20 patients, five repeated runs; per-patient median across runs) | median **8.8 s/patient** (IQR 5.1–16.6) |
| Reduction | ≈180-fold |
| Full cohort (400 patients) | median 7.1 s/patient (IQR 4.1–10.6; range 0.1–56.1); 57.5 min total |

Timing covers PDF text extraction and parsing per patient (`time.perf_counter`), not writing
the spreadsheet. Environment: Windows 10, 15.8 GB RAM, no GPU, Python 3.14.1, no network access.

## Results: second-center transportability

20 patients from an independent ICU with a different record system and largely free-text
discharge summaries; manual reference built with the same field definitions. The pipeline was
run unmodified (`icu_pipeline.py`) and after a targeted adaptation (`icu_pipeline_h2.py`; H2 mode
activates when the input path contains `hastane2`). The adaptation was developed and evaluated on
the same 20 patients.

Coverage = share of manually populated reference fields that the pipeline also populated.
Precision = share of pipeline-populated fields that matched the reference exactly.

| Domain | Reference fields | Coverage, unmodified % | Coverage, adapted % | Precision, adapted % |
|---|---:|---:|---:|---:|
| Demographics | 262 | 53.1 | 60.7 | 65.4 |
| Clinical severity scores | 309 | 0 | 100 | 100 |
| Laboratory values | 595 | 0 | 99.8 | 99.3 |
| Microbiology | 358 | 30.2 | 30.2 | 45.2 |
| Arterial blood gas | 1,606 | 0 | 0 | — |
| Antimicrobial courses | 287 | 0 | 0 | — |
| Vasopressor courses | 76 | 0 | 0 | — |
| **Total** | **3,493** | **7.1** | **33.5** | **86.3** |

## Limitations

- Single-center accuracy validation; rule development, reference abstraction and adjudication
  by one investigator. A worst-case accuracy is reported to bound this.
- Fields empty in both outputs were not verified against the source.
- 20-patient sample; limited precision at the patient level.
- Derived fields (ATC, AWaRe, DDD) were not evaluated.
- Second-center adaptation was developed and tested on the same patients.
- Rule-based parsing must be re-validated when report templates change.
