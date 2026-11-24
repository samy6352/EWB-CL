"""Source problems reach the Savings page as readable sentences, never as a Python dict."""
# cloudlens-test: requires-offline
from __future__ import annotations

from app import savings


def test_source_issue_is_flattened_to_text():
    issue = {"orphaned_nics_pricing": "12 resource(s) have no complete currency-qualified billing cost.",
             "unassociated_public_ips_pricing": "4 resource(s) have no complete currency-qualified billing cost."}
    text = savings._issue_text(issue)
    assert "{" not in text and "'" not in text
    assert text.count("billing cost.") == 2
    assert savings._issue_text(["a", "b"]) == "a; b"
    assert savings._issue_text("plain") == "plain"
