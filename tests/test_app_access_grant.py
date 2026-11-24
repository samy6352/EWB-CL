"""The administrator can grant the app identity its roles, and the list shows who has access."""
from __future__ import annotations

import asyncio

import pytest

from app import bootstrap, cost

pytestmark = pytest.mark.offline

OK_SUB = "11111111-1111-4111-8111-111111111111"
NEW_SUB = "22222222-2222-4222-8222-222222222222"
DENIED_SUB = "33333333-3333-4333-8333-333333333333"
PRINCIPAL = "44444444-4444-4444-8444-444444444444"


def test_grant_is_idempotent_per_role_and_reports_denials(monkeypatch):
    monkeypatch.setenv("CLOUDLENS_PRINCIPAL_ID", PRINCIPAL)
    monkeypatch.setattr(bootstrap, "require_writable", lambda: None)
    calls: list[tuple[str, str, dict]] = []

    async def request(method, url, **kwargs):
        calls.append((method, url, kwargs["json_body"]))
        if DENIED_SUB in url:
            raise cost.CostError("403 AuthorizationFailed")
        if OK_SUB in url:
            raise cost.CostError("409 RoleAssignmentExists")
        return {}

    monkeypatch.setattr(cost.azure, "request", request)
    results = {r["id"]: r for r in asyncio.run(bootstrap.grant_app_access([NEW_SUB, OK_SUB, DENIED_SUB, NEW_SUB]))}

    assert set(results) == {NEW_SUB, OK_SUB, DENIED_SUB}
    assert results[NEW_SUB]["ok"] and set(results[NEW_SUB]["roles"].values()) == {"granted"}
    assert results[OK_SUB]["ok"] and set(results[OK_SUB]["roles"].values()) == {"already granted"}
    assert not results[DENIED_SUB]["ok"] and "Owner" in results[DENIED_SUB]["detail"]
    new_calls = [c for c in calls if NEW_SUB in c[1]]
    assert len(new_calls) == len(bootstrap.ROLE_IDS)
    assert all(c[0] == "PUT" and c[2]["properties"]["principalId"] == PRINCIPAL for c in new_calls)
    assert len({c[1] for c in new_calls}) == len(bootstrap.ROLE_IDS)


def test_grant_needs_a_known_app_identity(monkeypatch):
    monkeypatch.delenv("CLOUDLENS_PRINCIPAL_ID", raising=False)
    monkeypatch.setattr(bootstrap, "require_writable", lambda: None)
    with pytest.raises(ValueError, match="CLOUDLENS_PRINCIPAL_ID"):
        asyncio.run(bootstrap.grant_app_access([NEW_SUB]))


def test_overview_unions_app_and_user_subscriptions(monkeypatch):
    monkeypatch.setenv("CLOUDLENS_PRINCIPAL_ID", PRINCIPAL)

    async def listing():
        return {"subscriptions": [{"id": OK_SUB.upper(), "name": "App only"}], "count": 1}

    async def request(method, url, **kwargs):
        if NEW_SUB in url:
            raise cost.CostError("403")
        return {"state": "Enabled"}

    monkeypatch.setattr(cost, "list_subscriptions", listing)
    monkeypatch.setattr(cost.azure, "request", request)
    mine = [{"id": NEW_SUB, "name": "Mine"}, {"id": DENIED_SUB, "name": "Also mine"}]
    out = asyncio.run(bootstrap.app_access_overview(mine))
    rows = {r["id"]: r for r in out["subscriptions"]}

    assert rows[OK_SUB]["app_access"] and not rows[OK_SUB]["user_access"]
    assert not rows[NEW_SUB]["app_access"] and rows[NEW_SUB]["user_access"]
    assert rows[DENIED_SUB]["app_access"] and rows[DENIED_SUB]["user_access"]
    assert out["can_grant"] and out["subscriptions"][0]["id"] == NEW_SUB
