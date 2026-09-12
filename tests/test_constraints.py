"""constraints：约束档的加载与默认值。"""
import json
import os

from src.constraints import Constraints, load_constraints


def test_load_constraints_none_when_no_path():
    assert load_constraints(None) is None
    assert load_constraints("") is None


def test_load_constraints_roundtrip(tmp_path):
    path = os.path.join(tmp_path, "c.json")
    data = {
        "locked_theme": "business",
        "require_data_source": True,
        "hard_fail_on_layout": True,
        "required_kinds": ["title", "thanks"],
        "max_kind_counts": {"chart": 3},
        "min_slides": 5,
        "max_slides": 15,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f)

    c = load_constraints(path)
    assert isinstance(c, Constraints)
    assert c.locked_theme == "business"
    assert c.require_data_source is True
    assert c.hard_fail_on_layout is True
    assert c.required_kinds == ["title", "thanks"]
    assert c.max_kind_counts == {"chart": 3}
    assert c.min_slides == 5
    assert c.max_slides == 15


def test_load_constraints_defaults_for_missing_keys(tmp_path):
    path = os.path.join(tmp_path, "c.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({}, f)

    c = load_constraints(path)
    assert c.locked_theme is None
    assert c.require_data_source is False
    assert c.hard_fail_on_layout is False
    assert c.required_kinds == []
    assert c.max_kind_counts == {}
    assert c.min_slides is None
    assert c.max_slides is None
