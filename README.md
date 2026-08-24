# CDA2-to-OMOP Microbiology ETL

This repository contains a Python-based Extract, Transform, Load (ETL) pipeline for converting anonymized microbiology laboratory reports encoded in HL7 Clinical Document Architecture release 2 (CDA2) into a structured representation based on the Observational Medical Outcomes Partnership Common Data Model version 5.4 (OMOP CDM v5.4).

The project focuses on the secondary use of microbiology laboratory data, with particular attention to culture results, microorganism identification, antimicrobial susceptibility testing, and Minimum Inhibitory Concentration (MIC) values.

## Repository contents

```text
.
├── etl_cda2_omop.py
├── vocab_lookup.py
├── example_cda2.xml
├── example_output.xlsx
└── README.md
```

- `etl_cda2_omop.py`: main ETL script.
- `vocab_lookup.py`: semantic vocabulary lookup module using OMOPHub.
- `example_cda2.xml`: anonymized CDA2 microbiology report.
- `example_output.xlsx`: example Excel output containing the generated OMOP tables.
