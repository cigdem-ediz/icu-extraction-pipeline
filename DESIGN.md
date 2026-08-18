# Design notes

Why this pipeline is built the way it is. This document focuses on the engineering and methods
behind the architecture and the decisions made during its construction. It does not interpret
the study's results, which belong in the associated manuscript.

## Why rule-based, not an LLM

The task involves extracting information from unstructured clinical documents that follow
highly regular conventions. These documents include ICU discharge summaries, daily forms,
laboratory reports, and microbiology reports. Each document follows a stable institutional
layout. A deterministic rule-based approach was deliberately chosen over a generative model
for three engineering reasons:

- **Auditability.** Every extracted value can be traced to the specific rule that produced
  it. When an output looks wrong, the cause is inspectable, not hidden in model weights.
- **Reproducibility.** The same input always yields the same output. There is no sampling
  temperature, no model-version drift, nothing to pin beyond the code itself.
- **Privacy.** Everything runs locally in a closed loop. No documents are sent to an external
  service, which is a hard requirement for this data, not a preference.

This is not a claim that rules beat learned models in general; they do not generalize across
heterogeneous free text the way generative models can. It is a claim that for a bounded
extraction task over stable formats, on data that must not leave the institution, the
deterministic route is the better engineering fit.

## Architecture

The pipeline maps four routine report types onto seven structured domains:

| Report type | Domains |
|---|---|
| Discharge summary (E1) | Demographics, severity scores, antimicrobial courses, vasopressor courses, blood gas |
| Daily ICU form (E2) | Severity scores (APACHE II, SOFA, GCS), sepsis / septic-shock flags |
| Laboratory report (LAB) | Laboratory values |
| Microbiology report (M) | Culture results, antibiogram / susceptibility |

Each report type has its own parser. Files are grouped by patient and protocol number,
parsed independently, and merged into one structured record. The parsing is text-based
(`pdfplumber`) and uses domain-specific rules tied to the layout of each report type. For
instance, daily narrative blocks are split on the `date + N.GÜN` marker. Blood-gas panels are
read from a fixed token sequence. Microbiological results are located by section headers.

## Turkish-specific engineering

Working in Turkish surfaced problems that do not exist in English pipelines and that shaped
the design:

- **Case folding.** The `str.casefold()` function in Python does not handle the Turkish
  dotted/dotless *i* correctly. Therefore, the same term can fold to two different strings
  depending on the case. This issue is most significant in the comparison layer, where a naive
  field comparison flags a large number of "differences" that are merely encoding artifacts. A
  Turkish-aware normalization step resolves this issue, and unit tests ensure that it cannot
  silently regress. The broader lesson — and a reason the comparison tool is versioned
  alongside the pipeline — is that, in a morphologically rich non-English language, the
  measurement instrument itself must be validated before it can measure anything.
- **Morphology and abbreviations.** Drug names, routes, and clinical terms appear in
  clinician shorthand and inflected forms; the rules normalize these to canonical values
  (e.g., ATC / WHO AWaRe / DDD lookups for antimicrobials) rather than matching surface text.

## Methodological choices

Two decisions in how the pipeline is evaluated are worth stating, because they are design
choices rather than afterthoughts:

- **Adjudicate against the source, not against the human.** When the automatic and manual
  values disagree, the comparison does not assume the manual value is correct; each
  disagreement is resolved by returning to the source PDF. This keeps the evaluation honest
  in both directions.
- **Version the comparison tool with the pipeline.** `validate.py` is part of the repository,
  not a throwaway script, so the exact concordance procedure is reproducible and inspectable.

## Reproducibility and privacy by construction

- Dependencies are pinned; a single Python version is specified.
- A de-identification step can pseudonymize direct identifiers on output, writing the
  pseudonym↔protocol key to a separate file that is never committed (see `.gitignore`).
- Real inputs and outputs are excluded from version control; only code and fully synthetic
  examples are published.

## Known engineering limitations

- Rules encode one center's conventions and require adaptation for other document layouts.
- Free-text microbiology is the least structured input and the hardest to parse reliably; it
  is the main target for future hardening.
- Route is inferred from the drug's standard formulation rather than parsed per prescription.
