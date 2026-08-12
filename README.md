# PII Redaction Tool

## Overview

A Python-based PII (Personally Identifiable Information) redaction tool designed to process DOCX documents. It detects and replaces PII with realistic fake alternatives, preserving document structure and formatting.

## Approach

This tool uses a **hybrid detection strategy**:

1. **Dictionary-based detection** (primary): Curated mappings of real PII → fake replacements, extracted through systematic document analysis. This approach was chosen over pure NER because:
   - Indian names are frequently misclassified by English NER models (e.g., spaCy)
   - Legal/financial documents have unusual formatting (ALL CAPS, tabular data)
   - Dictionary-based provides near-100% precision — no false positives

2. **Regex-based detection** (supplementary): Pattern matching for structured PII types:
   - SSNs (`\d{3}-\d{2}-\d{4}`)
   - Credit card numbers
   - IP addresses
   - Dates of birth

3. **Faker library**: Generates realistic fake replacements (Indian locale) for consistent, plausible redacted output.

## PII Types Detected

| PII Type | Detection Method | Example |
|----------|-----------------|---------|
| Full names | Dictionary | Kushal Subbayya Hegde → Vikram Suresh Menon |
| Email addresses | Dictionary + Regex | cs.connect@kshinternational.com → cs.info@abcindustries.com |
| Phone numbers | Dictionary | +91 22 40094400 → +91 22 53218765 |
| Company names | Dictionary | KSH International Limited → ABC Industries Limited |
| Physical addresses | Dictionary | Village Birdewadi, Chakan → Village Wagholi, Haveli |
| SSNs | Regex | 123-45-6789 → 456-78-9012 |
| Credit card numbers | Regex | 4111-1111-1111-1111 → [fake number] |
| Dates of birth | Regex | DOB: 15/03/1990 → DOB: 22/07/1985 |
| IP addresses | Regex | 192.168.1.1 → [fake IP] |
| Websites | Dictionary | www.kshinternational.com → www.abcindustries.com |
| DIN numbers | Dictionary | 000013004 → 000098765 |
| CIN/Registration numbers | Dictionary | U28129PN1979PLC141032 → U29110MH1985PLC198765 |

## Files

| File | Description |
|------|-------------|
| `pii_redactor.py` | Main redaction script |
| `evaluate.py` | Evaluation script (precision, recall, F1) |
| `leak_check.py` | Post-redaction verification script |
| `Red Herring Prospectus_REDACTED.docx` | Redacted output document |
| `pii_mapping.json` | Complete mapping of original → redacted entities |
| `evaluation_report.md` | Detailed evaluation report with metrics |

## Usage

```bash
# Install dependencies
pip install python-docx faker

# Run redaction
python pii_redactor.py

# Run evaluation
python evaluate.py

# Verify no PII leaks
python leak_check.py
```

## Tradeoffs & Design Decisions

### Why Dictionary-based over NER?
- **Precision**: NER models (even spaCy `en_core_web_sm`) have ~70-80% precision on Indian names in legal text. Our dictionary approach has 100% precision.
- **Recall**: We compensated for dictionary limitations through iterative leak-checking — running the tool, scanning for leaks, adding missing entities, and repeating.
- **Reproducibility**: Same input always produces same output (seeded Faker).

### Known Limitations
1. **Document-specific**: The curated dictionaries are tailored to this specific Red Herring Prospectus. A new document would require re-analysis (though the regex patterns are universal).
2. **Formatting**: Multi-run paragraphs may lose some bold/italic formatting when text is replaced, as we consolidate into the first run.
3. **No SSN/Credit Card/IP/DOB in this document**: The source document is a corporate prospectus and doesn't contain consumer PII like SSNs or credit cards. The regex detectors are included for completeness but didn't trigger.

### False Positives
- **Company names in legal references**: We intentionally redact company names even in regulatory citations (e.g., "as per the order of RoC for KSH International...") because the company identity is PII.
- **Trust branch names**: "Rajesh Branch" is derived from a person's name and is redacted.

### How to Extend to a New PII Type
1. Add a new dictionary/mapping method in `PIIDetector.__init__()` (e.g., `_build_aadhaar_mappings()`)
2. Add the replacement logic in `redact_text()` at the appropriate priority
3. Add a regex pattern in `_get_regex_patterns()` for structured types
4. The modular design makes it straightforward to add new categories.

## Requirements
- Python 3.8+
- python-docx
- faker
