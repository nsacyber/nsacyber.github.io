#!/usr/bin/env python3
"""Validate the ZYRA/VIRGINIA public defensive-cyber ontology."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
ONTOLOGY_PATH = ROOT / "virginia" / "public-cyber-intel-ontology.json"


def require(condition: bool, code: str, errors: list[str]) -> None:
    if not condition:
        errors.append(code)


def load_ontology() -> dict[str, Any]:
    return json.loads(ONTOLOGY_PATH.read_text(encoding="utf-8"))


def validate(ontology: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    source = ontology.get("sourceBoundary", {})
    policy = ontology.get("actionPolicy", {})
    controls = ontology.get("controls", {})
    pipeline = ontology.get("pipeline", [])

    require(
        ontology.get("id") == "ZYRA-VIRGINIA-PUBLIC-CYBER-INTEL-ONTOLOGY",
        "ONTOLOGY_ID_INVALID",
        errors,
    )
    require(
        ontology.get("command") == "/glass virginia intel",
        "COMMAND_INVALID",
        errors,
    )
    require(
        source.get("authoritativeUpstream") == "nsacyber/nsacyber.github.io",
        "UPSTREAM_INVALID",
        errors,
    )
    require(
        source.get("sourceClass") == "PUBLIC_OPEN_SOURCE",
        "SOURCE_MUST_BE_PUBLIC",
        errors,
    )
    require(
        source.get("agencyAffiliationClaim") is False,
        "AGENCY_AFFILIATION_MUST_NOT_BE_CLAIMED",
        errors,
    )
    require(
        source.get("governmentEndorsementClaim") is False,
        "ENDORSEMENT_MUST_NOT_BE_CLAIMED",
        errors,
    )

    orders = [stage.get("order") for stage in pipeline]
    require(
        orders == list(range(1, len(pipeline) + 1)),
        "PIPELINE_ORDER_INVALID",
        errors,
    )
    require(
        any(stage.get("component") == "VIRGINIA_POLICY" for stage in pipeline),
        "VIRGINIA_POLICY_STAGE_REQUIRED",
        errors,
    )
    require(
        any(stage.get("action") == "VERIFY_PROVENANCE" for stage in pipeline),
        "PROVENANCE_STAGE_REQUIRED",
        errors,
    )
    require(
        any(stage.get("action") == "REQUIRE_HUMAN_REVIEW" for stage in pipeline),
        "HUMAN_REVIEW_STAGE_REQUIRED",
        errors,
    )

    required_blocks = {
        "CLAIM_CLASSIFIED_ACCESS",
        "CLAIM_AGENCY_AUTHORITY",
        "INGEST_NONPUBLIC_MATERIAL",
        "STORE_CREDENTIAL",
        "EXECUTE_EXPLOIT",
        "DEPLOY_MALWARE",
        "TARGET_REAL_SYSTEM",
        "BYPASS_ACCESS_CONTROL",
        "AUTOMATE_OFFENSIVE_ACTION",
    }
    require(
        required_blocks.issubset(set(policy.get("block", []))),
        "BLOCK_POLICY_INCOMPLETE",
        errors,
    )

    require(controls.get("publicSourcesOnly") is True, "PUBLIC_ONLY_REQUIRED", errors)
    require(controls.get("provenanceRequired") is True, "PROVENANCE_REQUIRED", errors)
    require(controls.get("confidenceRequired") is True, "CONFIDENCE_REQUIRED", errors)
    require(
        controls.get("humanReviewForDeployment") is True,
        "DEPLOYMENT_REVIEW_REQUIRED",
        errors,
    )
    require(
        controls.get("classifiedAccessClaim") is False,
        "CLASSIFIED_CLAIM_BLOCK_REQUIRED",
        errors,
    )
    require(
        controls.get("agencyAuthorityClaim") is False,
        "AGENCY_AUTHORITY_BLOCK_REQUIRED",
        errors,
    )
    require(
        controls.get("offensiveExecution") is False,
        "OFFENSIVE_EXECUTION_BLOCK_REQUIRED",
        errors,
    )
    require(
        controls.get("realWorldTargeting") is False,
        "TARGETING_BLOCK_REQUIRED",
        errors,
    )
    require(
        controls.get("preserveOriginalDisclaimer") is True,
        "ORIGINAL_DISCLAIMER_REQUIRED",
        errors,
    )
    return errors


if __name__ == "__main__":
    ontology = load_ontology()
    validation_errors = validate(ontology)
    if validation_errors:
        print(json.dumps({"valid": False, "errors": validation_errors}, indent=2))
        raise SystemExit(1)

    print(
        json.dumps(
            {
                "valid": True,
                "ontology": ontology["id"],
                "pipelineStages": len(ontology["pipeline"]),
                "policy": "VIRGINIA",
                "sourceClass": "PUBLIC_OPEN_SOURCE",
            },
            indent=2,
        )
    )
