# VIRGINIA Public Cyber Intelligence Ontology

Command: **`/glass virginia intel`**

This XUNIA/ZYRA extension converts public cybersecurity publications and open-source project metadata into provenance-bearing defensive guidance. It operates at the VIRGINIA policy boundary between source ingestion and any consequential action.

## Authority and source boundary

- Authoritative upstream: [`nsacyber/nsacyber.github.io`](https://github.com/nsacyber/nsacyber.github.io)
- XUNIA integration repository: [`sonoxo/nsacyber.github.ioZYRA`](https://github.com/sonoxo/nsacyber.github.ioZYRA)
- Permitted source class: public, open-source material only
- Classification claim: unclassified/public only
- NSA, United States Government, or vendor affiliation/endorsement claim: none

The original [`LICENSE.md`](../LICENSE.md), [`INTENT.md`](../INTENT.md), and [`DISCLAIMER.md`](../DISCLAIMER.md) remain controlling source notices.

## Ontology

```text
PUBLIC SOURCE
  -> PUBLICATION / SOFTWARE PROJECT / ADVISORY
  -> VULNERABILITY CLASS
  -> MITIGATION
  -> SECURITY CONTROL
  -> EVIDENCE
  -> ASSESSMENT
  -> VIRGINIA POLICY DECISION
  -> SECRET-FREE AUDIT EVENT
```

## Pipeline

```text
PUBLIC INGEST
  -> PROVENANCE VERIFY
  -> PUBLIC CONTENT CLASSIFY
  -> DEFENSIVE GUIDANCE EXTRACT
  -> SECURITY CONTROL MAP
  -> CONFIDENCE ASSESS
  -> HUMAN REVIEW
  -> SECRET-FREE AUDIT
```

## Decision policy

Allowed actions include reading public sources, indexing public metadata, defensive summarization, control mapping, comparison, simulation, and secret-free reporting.

External claims, repository changes, control deployment, and remediation automation require human review.

Claims of classified access or agency authority, ingestion of nonpublic material, credential storage, exploit execution, malware deployment, real-system targeting, access-control bypass, and automated offensive action are blocked.

## Validate

```bash
python3 scripts/validate_virginia_ontology.py
```

Machine contract: [`public-cyber-intel-ontology.json`](public-cyber-intel-ontology.json)
