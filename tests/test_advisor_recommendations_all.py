import asyncio

from app import waste

SUB = "11111111-1111-1111-1111-111111111111"


def _run(category: str):
    calls = []

    async def fake_request(method, url, params=None, **kwargs):
        calls.append(params)
        return {"value": [{
            "id": f"/subscriptions/{SUB}/providers/Microsoft.Advisor/recommendations/r1",
            "properties": {
                "category": "Security", "impact": "High",
                "shortDescription": {"problem": "Enable MFA", "solution": "Enable it"},
                "resourceMetadata": {"resourceId": f"/subscriptions/{SUB}/resourceGroups/rg/providers/x/y/name"},
            },
        }]}

    original = waste.azure.request
    waste.azure.request = fake_request
    try:
        result = asyncio.run(waste.advisor_recommendations([SUB], category=category))
    finally:
        waste.azure.request = original
    return calls, result


def test_all_categories_sends_no_category_filter_and_keeps_pillar():
    calls, result = _run("All")
    assert "$filter" not in calls[0]
    row = result["recommendations"][0]
    assert row["category"] == "Security" and row["resourceGroup"] == "rg"


def test_single_category_still_filters():
    calls, _ = _run("Cost")
    assert calls[0]["$filter"] == "Category eq 'Cost'"
