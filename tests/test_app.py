import io
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app import app, calculate


def sample_payload():
    return {"meta": {"class_code": "FFAR 248", "class_cap": 1000},
            "selected": [{"category": "assessments", "label": "Workshop", "points": 3, "quantity": 2},
                         {"category": "instructor-access", "label": "Removes Complexity", "points": -1, "quantity": 1}],
            "custom": {"groups": {"description": "Extra cohorts", "points": 2, "quantity": 3}},
            "gamification_points": 0.75,
            "gamification_quantity": 2,
            "context": "Context", "notes": "Notes"}


def test_calculation():
    result = calculate(sample_payload())
    assert result["cap_points"] == 2
    assert result["total"] == 14.5
    assert result["breakdown"]["assessments"] == 6
    assert result["breakdown"]["groups"] == 6
    assert result["breakdown"]["manual-intervention"] == 1.5
    workshop = next(row for row in result["rows"] if row["item"] == "Workshop")
    assert workshop == {"category": "Assessments", "item": "Workshop", "quantity": 2,
                        "unit_points": 3.0, "points": 6.0}


def test_routes_and_exports():
    client = app.test_client()
    assert client.get("/").status_code == 200
    for kind, signature in [("pdf", b"%PDF"), ("xlsx", b"PK")]:
        response = client.post(f"/export/{kind}", data={"payload": json.dumps(sample_payload())})
        assert response.status_code == 200
        assert response.data.startswith(signature)


def test_revised_specification_is_present():
    html = app.test_client().get("/").get_data(as_text=True)
    for expected in ["Simple Upload", "+0.25", "Forum Marking Guide", "Assignment Marking Guide",
                     "Gamification or interactivity", "Has Supergroups", "Requires ACSD Overrides",
                     "Normal Communication", "Changes This Semester"]:
        assert expected in html
    assert "addQuantityInputs" in html
    assert "syncFromQuantity" in html
    assert "syncFromCheckbox" in html
    assert "align-self: start" in html
    assert "custom-entry-row" in html
    assert "grid-template-columns: auto minmax(0, 1fr) auto auto" in html
    for removed in ["Forum Grading", "Gradebook Setup", "New for Semester?", ">Negligible<"]:
        assert removed not in html


def test_quantity_defaults_to_zero():
    payload = {"meta": {"class_cap": 0},
               "selected": [{"category": "assessments", "label": "Workshop", "points": 3}],
               "custom": {"groups": {"description": "No quantity", "points": 5}},
               "gamification_points": 2}
    result = calculate(payload)
    assert result["total"] == 0
    assert all(row["quantity"] == 0 and row["points"] == 0 for row in result["rows"])
