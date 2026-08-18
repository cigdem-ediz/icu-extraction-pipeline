# Synthetic examples

These files are **entirely fictional**. Every patient name, protocol number, national ID,
date, laboratory value, and microbiology result is invented, and no real institution is
named. Each page carries a visible banner:

> SYNTHETIC EXAMPLE — NOT REAL PATIENT DATA

They exist so anyone can run the pipeline end-to-end without any real patient data.

## The three example patients

| Patient | Presentation | Course | Culture |
|---|---|---|---|
| 99001 | Pneumosepsis with acute kidney injury | 8-day ICU stay | *Klebsiella pneumoniae* |
| 99002 | Urosepsis with septic shock | 5-day ICU stay | *Escherichia coli* |
| 99003 | Aspiration pneumonia, respiratory failure | 4-day ICU stay | No growth |

Each patient has four report types — discharge summary (`E1`), daily ICU form (`E2`),
laboratory report (`LAB`), and microbiology (`M1`) — with multi-day clinical narratives of
realistic length.

## Run the demo

From the repository root:

```bash
python icu_pipeline.py --klasor examples --cikti demo_output
```

The pipeline processes all three patients and extracts demographics, clinical severity scores,
laboratory values, blood gas, antimicrobial courses, culture results, and antibiograms.

## Regenerate or edit

The PDFs are produced by `generate_synthetic.py`. Inspect it to see exactly what goes in, or
change the values and regenerate:

```bash
cd examples
python generate_synthetic.py
```

Nothing in this folder is real, and nothing derived from real records is included.
