from __future__ import annotations

import base64
import io
import json
import struct
from hashlib import sha256
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

from wcab.adapters import formulafence
from wcab.build import CASE_IDS, build_all
from wcab.manifest import case_rows, manifest_text
from wcab.validate import validate_all, validate_case

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _tree_hashes(root: Path) -> dict[str, str]:
    return {
        path.relative_to(root).as_posix(): sha256(path.read_bytes()).hexdigest()
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }


def _auto_filter_criteria_details() -> dict[str, object]:
    surface = {
        "present": True,
        "worksheet_auto_filter_count": 1,
        "table_auto_filter_count": 0,
        "filter_column_count": 1,
        "filter_criterion_count": 1,
        "sort_state_count": 0,
        "sort_condition_count": 0,
        "default_hidden_sheet_count": 0,
        "default_zero_height_sheet_count": 0,
        "default_zero_width_sheet_count": 0,
        "hidden_row_count": 0,
        "zero_height_row_count": 0,
        "outlined_row_count": 0,
        "collapsed_row_count": 0,
        "hidden_column_count": 0,
        "zero_width_column_count": 0,
        "outlined_column_count": 0,
        "collapsed_column_count": 0,
        "visible_row_override_count": 0,
        "unrecognized_control_count": 0,
    }
    return {
        "before": surface,
        "after": surface,
        "filter_visibility_definition_material_changed": True,
    }


def _pivot_cache_refresh_details() -> dict[str, object]:
    before = {
        "background_query": False,
        "cache_id": 1,
        "connection_id": None,
        "opaque_metadata": {"count": 0, "present": False},
        "refresh_enabled": True,
        "refresh_on_load": False,
        "save_data": True,
        "source_type": "worksheet",
        "upgrade_on_refresh": False,
    }
    return {"before": [before], "after": [{**before, "refresh_on_load": True}]}


def _chart_series_reference_details() -> dict[str, object]:
    profile = {
        "present": True,
        "chart_host_sheet_count": 1,
        "chart_drawing_part_count": 1,
        "chart_reference_count": 1,
        "chart_part_count": 1,
        "chart_ex_reference_count": 0,
        "chart_ex_part_count": 0,
        "chart_ex_series_count": 0,
        "chart_ex_title_count": 0,
        "chart_ex_data_reference_count": 0,
        "chart_user_shape_part_count": 0,
        "chart_user_shape_count": 0,
        "chart_type_count": 1,
        "series_count": 1,
        "title_count": 1,
        "data_reference_count": 3,
        "numeric_data_reference_count": 2,
        "string_data_reference_count": 1,
        "literal_data_point_count": 0,
        "cached_data_point_count": 0,
        "pivot_source_count": 0,
        "external_data_reference_count": 0,
        "user_shape_reference_count": 0,
        "related_relationship_count": 1,
        "external_relationship_count": 0,
        "internal_related_part_count": 0,
        "fingerprinted_related_part_count": 0,
        "uninspected_related_part_count": 0,
        "unrecognized_part_count": 0,
    }
    return {
        "before": profile,
        "after": profile,
        "chart_definition_material_changed": True,
    }


def _pivot_data_field_aggregation_details() -> dict[str, object]:
    profile = {
        "present": True,
        "pivot_table_sheet_count": 1,
        "pivot_table_part_count": 1,
        "pivot_cache_definition_part_count": 1,
        "pivot_cache_records_part_count": 1,
        "pivot_cache_binding_count": 1,
        "layout_location_count": 1,
        "pivot_field_count": 2,
        "row_field_count": 1,
        "column_field_count": 0,
        "page_field_count": 0,
        "data_field_count": 1,
        "filter_count": 0,
        "row_item_count": 0,
        "column_item_count": 0,
        "cache_field_count": 2,
        "shared_item_count": 6,
        "calculated_item_count": 0,
        "calculated_member_count": 0,
        "cache_record_count": 4,
        "related_relationship_count": 2,
        "external_relationship_count": 0,
        "fingerprinted_cache_record_part_count": 1,
        "uninspected_cache_record_part_count": 0,
        "unrecognized_part_count": 0,
    }
    return {
        "before": profile,
        "after": profile,
        "pivot_table_layout_material_changed": True,
    }


def _pivot_slicer_selection_details() -> dict[str, object]:
    profile = {
        "present": True,
        "slicer_cache_part_count": 1,
        "timeline_cache_part_count": 0,
        "slicer_workbook_binding_count": 1,
        "timeline_workbook_binding_count": 0,
        "slicer_pivot_cache_binding_count": 1,
        "slicer_table_binding_count": 0,
        "timeline_pivot_cache_binding_count": 0,
        "slicer_pivot_table_binding_count": 1,
        "timeline_pivot_table_binding_count": 0,
        "slicer_item_count": 2,
        "selected_slicer_item_count": 1,
        "timeline_state_count": 0,
        "timeline_filter_count": 0,
        "related_relationship_count": 0,
        "external_relationship_count": 0,
        "unrecognized_part_count": 0,
    }
    return {
        "before": profile,
        "after": profile,
        "slicer_filter_state_or_definition_material_changed": True,
    }


def _power_query_m_filter_details() -> dict[str, object]:
    profile = {
        "present": True,
        "mashup_count": 1,
        "parsed_mashup_count": 1,
        "formula_document_count": 1,
        "package_part_count": 3,
        "embedded_content_part_count": 0,
        "metadata_document_count": 1,
        "metadata_item_count": 1,
        "permission_controls": {
            "payload_count": 1,
            "parsed_count": 1,
            "firewall_enabled_count": 1,
            "future_packages_allowed_count": 0,
            "workbook_group_type_count": 0,
            "opaque_metadata": {"present": False, "count": 0},
        },
        "permission_binding_count": 0,
        "opaque_metadata": {"present": False, "count": 0},
    }
    return {
        "before": profile,
        "after": profile,
        "formula_material_changed": True,
    }


def _replace_power_query_m_filter_literal(path: Path, *, before: bytes, after: bytes) -> None:
    """Rewrite one test fixture's embedded M formula without changing its truth."""

    with ZipFile(path) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    mashup_member = "customXml/item1.xml"
    mashup = ElementTree.fromstring(members[mashup_member])
    assert isinstance(mashup.text, str)
    stream = base64.b64decode(mashup.text, validate=True)
    version = struct.unpack_from("<I", stream)[0]
    cursor = 4
    fields: list[bytes] = []
    for _ in range(4):
        size = struct.unpack_from("<I", stream, cursor)[0]
        cursor += 4
        fields.append(stream[cursor : cursor + size])
        cursor += size
    assert cursor == len(stream)
    package_payload = io.BytesIO()
    with ZipFile(io.BytesIO(fields[0])) as package:
        package_members = {
            entry.filename: package.read(entry.filename) for entry in package.infolist()
        }
    formula_member = "Formulas/Section1.m"
    assert package_members[formula_member].count(before) == 1
    package_members[formula_member] = package_members[formula_member].replace(before, after)
    with ZipFile(package_payload, "w", compression=ZIP_DEFLATED) as package:
        for name in sorted(package_members):
            package.writestr(name, package_members[name])
    fields[0] = package_payload.getvalue()
    rewritten_stream = struct.pack("<I", version)
    for field in fields:
        rewritten_stream += struct.pack("<I", len(field)) + field
    mashup.text = base64.b64encode(rewritten_stream).decode("ascii")
    members[mashup_member] = ElementTree.tostring(mashup, encoding="utf-8", xml_declaration=True)
    staging = path.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(path)


def test_committed_fixtures_validate() -> None:
    assert validate_all(PROJECT_ROOT / "fixtures") == {}


def test_build_is_byte_reproducible(tmp_path: Path) -> None:
    first = tmp_path / "first"
    second = tmp_path / "second"
    assert len(build_all(first)) == len(CASE_IDS)
    assert len(build_all(second)) == len(CASE_IDS)
    assert _tree_hashes(first) == _tree_hashes(second)


def test_committed_tree_matches_regeneration(tmp_path: Path) -> None:
    regenerated = tmp_path / "regenerated"
    build_all(regenerated)
    assert _tree_hashes(PROJECT_ROOT / "fixtures") == _tree_hashes(regenerated)


def test_truth_ids_are_complete_and_unique() -> None:
    manifests = sorted((PROJECT_ROOT / "fixtures").rglob("truth.json"))
    identifiers = [json.loads(path.read_text(encoding="utf-8"))["id"] for path in manifests]
    assert tuple(sorted(identifiers)) == tuple(sorted(CASE_IDS))
    assert len(set(identifiers)) == len(identifiers)


def test_committed_manifest_matches_fixture_tree() -> None:
    fixture_root = PROJECT_ROOT / "fixtures"
    rows = case_rows(fixture_root, expected_ids=CASE_IDS)
    assert [row["id"] for row in rows] == sorted(CASE_IDS)
    assert all(row["baseline_files"] and row["candidate_files"] for row in rows)
    assert all("coverage_expectations" in row for row in rows)
    dynamic_row = next(
        row for row in rows if row["id"] == "structural.dynamic_reference_introduced"
    )
    assert dynamic_row["coverage_expectations"] == [
        {
            "kind": "dynamic_reference_static_coverage",
            "sheet": "Summary",
            "cell": "B2",
            "functions": ["INDIRECT"],
        }
    ]
    driver_row = next(
        row for row in rows if row["id"] == "structural.indirect_reference_driver_changed"
    )
    assert driver_row["coverage_expectations"] == [
        {
            "kind": "dynamic_reference_driver_changed",
            "driver": {"sheet": "Inputs", "cell": "E12"},
            "formula": {"sheet": "Summary", "cell": "B2"},
            "functions": ["INDIRECT"],
        }
    ]
    auto_filter_row = next(
        row for row in rows if row["id"] == "operations.auto_filter_criteria_changed"
    )
    assert auto_filter_row["facts"] == [
        {
            "kind": "auto_filter_criteria_changed",
            "sheet": "Report",
            "filter_ref": "A1:B5",
            "filter_column_id": 0,
            "baseline_filter_value": "North",
            "candidate_filter_value": "South",
            "subtotal_cell": "D2",
            "subtotal_formula": "=SUBTOTAL(109,B2:B5)",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Report!$D$2",
        }
    ]
    external_data_row = next(
        row for row in rows if row["id"] == "governance.external_data_refresh_on_open"
    )
    assert external_data_row["facts"] == [
        {
            "kind": "external_data_connection_refresh_on_load_changed",
            "connection_id": 1,
            "baseline_refresh_on_load": False,
            "candidate_refresh_on_load": True,
        }
    ]
    pivot_cache_row = next(
        row for row in rows if row["id"] == "governance.pivot_cache_refresh_on_open"
    )
    assert pivot_cache_row["facts"] == [
        {
            "kind": "pivot_cache_refresh_on_load_changed",
            "cache_id": 1,
            "source_type": "worksheet",
            "source_sheet": "Source",
            "source_ref": "A1:B5",
            "pivot_sheet": "Report",
            "pivot_ref": "A1:B2",
            "pivot_output_cell": "B2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Report!$B$2",
            "baseline_refresh_on_load": False,
            "candidate_refresh_on_load": True,
        }
    ]
    pivot_aggregation_row = next(
        row for row in rows if row["id"] == "structural.pivot_data_field_aggregation_changed"
    )
    assert pivot_aggregation_row["facts"] == [
        {
            "kind": "pivot_data_field_aggregation_changed",
            "cache_id": 1,
            "source_type": "worksheet",
            "source_sheet": "Source",
            "source_ref": "A1:B5",
            "pivot_sheet": "Report",
            "pivot_ref": "A1:B2",
            "pivot_output_cell": "B2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Report!$B$2",
            "data_field_source_index": 1,
            "baseline_subtotal": "sum",
            "candidate_subtotal": "average",
        }
    ]
    pivot_slicer_row = next(
        row for row in rows if row["id"] == "structural.pivot_slicer_selection_changed"
    )
    assert pivot_slicer_row["facts"] == [
        {
            "kind": "pivot_slicer_selection_changed",
            "cache_id": 1,
            "source_type": "worksheet",
            "source_sheet": "Source",
            "source_ref": "A1:B5",
            "pivot_sheet": "Report",
            "pivot_ref": "A1:B2",
            "pivot_output_cell": "B2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Report!$B$2",
            "slicer_name": "WCAB Region slicer",
            "slicer_source_name": "Region",
            "slicer_pivot_table_name": "WCAB Pivot Report",
            "slicer_pivot_tab_id": 2,
            "item_count": 2,
            "baseline_selected_item_index": 0,
            "candidate_selected_item_index": 1,
            "baseline_selected_value": "North",
            "candidate_selected_value": "South",
        }
    ]
    power_query_row = next(
        row for row in rows if row["id"] == "structural.power_query_m_filter_changed"
    )
    assert power_query_row["facts"] == [
        {
            "kind": "power_query_m_filter_changed",
            "data_mashup_part": "customXml/item1.xml",
            "source_sheet": "Source",
            "source_table": "SourceData",
            "source_ref": "A1:B5",
            "query_section": "Section1",
            "query_name": "RegionQuery",
            "filter_column": "Region",
            "baseline_filter_value": "North",
            "candidate_filter_value": "South",
            "fill_enabled": False,
            "firewall_enabled": True,
            "future_packages_allowed": False,
        }
    ]
    external_link_policy_row = next(
        row for row in rows if row["id"] == "governance.external_workbook_link_update_on_open"
    )
    assert external_link_policy_row["facts"] == [
        {
            "kind": "external_workbook_link_update_policy_changed",
            "sheet": "LinkedModel",
            "cell": "B2",
            "formula": "='[WCABSource.xlsx]Inputs'!$B$2",
            "baseline_update_links": "never",
            "candidate_update_links": "always",
        }
    ]
    iterative_calculation_row = next(
        row for row in rows if row["id"] == "governance.iterative_calculation_enabled"
    )
    assert iterative_calculation_row["facts"] == [
        {
            "kind": "iterative_calculation_enabled",
            "sheet": "Model",
            "cell": "B2",
            "formula": "=(B2+Inputs!$B$2)/2",
            "baseline_iterate": False,
            "candidate_iterate": True,
            "iteration_count": 100,
            "iteration_delta": 0.001,
        }
    ]
    precision_as_displayed_row = next(
        row for row in rows if row["id"] == "governance.precision_as_displayed_enabled"
    )
    assert precision_as_displayed_row["facts"] == [
        {
            "kind": "precision_as_displayed_enabled",
            "input_sheet": "Inputs",
            "input_cell": "B2",
            "input_value": 10.005,
            "number_format": "0.00",
            "formula_sheet": "Model",
            "formula_cell": "B2",
            "formula": "=Inputs!$B$2*2",
            "baseline_full_precision": True,
            "candidate_full_precision": False,
        }
    ]
    workbook_date_system_row = next(
        row for row in rows if row["id"] == "governance.workbook_date_system_changed"
    )
    assert workbook_date_system_row["facts"] == [
        {
            "kind": "workbook_date_system_changed",
            "baseline_date_1904": False,
            "candidate_date_1904": True,
            "date_compatibility": True,
            "serial_sheet": "Inputs",
            "serial_cell": "B2",
            "serial_value": 45292,
            "number_format": "yyyy-mm-dd",
            "formula_sheet": "Model",
            "formula_cell": "B2",
            "formula": "=Inputs!$B$2+30",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Model!$B$2",
        }
    ]
    formula_cached_result_row = next(
        row for row in rows if row["id"] == "governance.formula_cached_result_changed"
    )
    assert formula_cached_result_row["facts"] == [
        {
            "kind": "formula_cached_result_changed",
            "sheet": "Model",
            "cell": "B2",
            "formula": "=Inputs!$B$2*2",
            "input_sheet": "Inputs",
            "input_cell": "B2",
            "input_value": 10,
            "result_type": "numeric",
            "baseline_cached_result": 20,
            "candidate_cached_result": 25,
        }
    ]
    chart_series_row = next(
        row for row in rows if row["id"] == "structural.chart_series_reference_changed"
    )
    assert chart_series_row["facts"] == [
        {
            "kind": "chart_series_value_reference_changed",
            "chart_sheet": "Dashboard",
            "chart_anchor": "D2",
            "source_sheet": "Source",
            "series_title_ref": "'Source'!B1",
            "category_ref": "'Source'!$A$2:$A$4",
            "baseline_value_ref": "'Source'!$B$2:$B$4",
            "candidate_value_ref": "'Source'!$C$2:$C$4",
        }
    ]
    array_row = next(row for row in rows if row["id"] == "structural.array_formula_mode_changed")
    assert array_row["facts"] == [
        {
            "kind": "array_formula_mode_changed",
            "sheet": "Model",
            "cell": "B1",
            "formula": "=LEN(Inputs!A1:A3)",
            "baseline_mode": "legacy_cse",
            "candidate_mode": "dynamic",
            "baseline_output_range": "B1:B3",
            "candidate_output_range": "B1:B3",
        }
    ]
    assert (fixture_root / "manifest.jsonl").read_text(encoding="utf-8") == manifest_text(rows)


def test_manifest_is_reproducible_with_fixture_build(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    first = (fixture_root / "manifest.jsonl").read_text(encoding="utf-8")
    build_all(fixture_root)
    assert (fixture_root / "manifest.jsonl").read_text(encoding="utf-8") == first


def test_validator_rejects_a_false_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "finance" / "formula_to_value"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["cell"] = "A1"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_structured_table_scope(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "structured_table_scope_expansion"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_ref"] = "A1:D6"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_dynamic_reference_coverage_expectation(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "dynamic_reference_introduced"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["coverage_expectations"][0]["functions"] = ["OFFSET"]
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_dynamic_reference_driver_expectation(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "indirect_reference_driver_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["coverage_expectations"][0]["formula"]["cell"] = "B3"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_external_data_refresh_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_data_refresh_on_open"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_refresh_on_load"] = False
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_pivot_cache_refresh_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "pivot_cache_refresh_on_open"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_refresh_on_load"] = False
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_pivot_cache_refresh_flag(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "pivot_cache_refresh_on_open" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    definition_member = "xl/pivotCache/pivotCacheDefinition1.xml"
    definition = ElementTree.fromstring(members[definition_member])
    definition.set("refreshOnLoad", "0")
    members[definition_member] = ElementTree.tostring(
        definition, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_pivot_cache_control_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "pivot_cache_refresh_on_open" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    definition_member = "xl/pivotCache/pivotCacheDefinition1.xml"
    definition = ElementTree.fromstring(members[definition_member])
    definition.set("enableRefresh", "0")
    members[definition_member] = ElementTree.tostring(
        definition, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_pivot_cache_source_binding(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "pivot_cache_refresh_on_open" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    definition_member = "xl/pivotCache/pivotCacheDefinition1.xml"
    definition = ElementTree.fromstring(members[definition_member])
    source = definition.find(f".//{{{namespace}}}worksheetSource")
    assert source is not None
    source.set("ref", "A1:B4")
    members[definition_member] = ElementTree.tostring(
        definition, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_pivot_cache_records_part(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "pivot_cache_refresh_on_open" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    records_member = "xl/pivotCache/pivotCacheRecords1.xml"
    records = ElementTree.fromstring(members[records_member])
    second_record = records.findall(f"{{{namespace}}}r")[1]
    second_record.findall(f"{{{namespace}}}x")[1].set("v", "3")
    members[records_member] = ElementTree.tostring(records, encoding="utf-8", xml_declaration=True)
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_pivot_cache_relationship_binding(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "pivot_cache_refresh_on_open" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "xl/pivotTables/_rels/pivotTable1.xml.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationship = next(
        relationship
        for relationship in relationships
        if relationship.get("Id") == "rIdWCABPivotCache"
    )
    relationship.set("Target", "../pivotCache/pivotCacheDefinition2.xml")
    members[relationships_member] = ElementTree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_pivot_cache_refresh_pair_changes_only_its_cache_definition(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "pivot_cache_refresh_on_open"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/pivotCache/pivotCacheDefinition1.xml"]


def test_validator_rejects_a_false_pivot_data_field_aggregation_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "pivot_data_field_aggregation_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_subtotal"] = "sum"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_pivot_data_field_aggregation(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "structural" / "pivot_data_field_aggregation_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    pivot_member = "xl/pivotTables/pivotTable1.xml"
    pivot_table = ElementTree.fromstring(members[pivot_member])
    data_field = pivot_table.find(f".//{{{namespace}}}dataField")
    assert data_field is not None
    data_field.set("subtotal", "sum")
    members[pivot_member] = ElementTree.tostring(
        pivot_table, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_pivot_data_field_relationship(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "structural" / "pivot_data_field_aggregation_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "xl/pivotTables/_rels/pivotTable1.xml.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationship = next(item for item in relationships if item.get("Id") == "rIdWCABPivotCache")
    relationship.set(
        "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    )
    members[relationships_member] = ElementTree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_pivot_data_field_aggregation_pair_changes_only_its_pivot_table(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "pivot_data_field_aggregation_changed"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        assert baseline.testzip() is None
        assert candidate.testzip() is None
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/pivotTables/pivotTable1.xml"]


def test_validator_rejects_a_false_pivot_slicer_selection_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "pivot_slicer_selection_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_selected_item_index"] = 0
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_pivot_slicer_selection(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "pivot_slicer_selection_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    slicer_member = "xl/slicerCaches/slicerCache1.xml"
    slicer_cache = ElementTree.fromstring(members[slicer_member])
    items = slicer_cache.findall(
        f"{{{namespace}}}data/{{{namespace}}}tabular/{{{namespace}}}items/{{{namespace}}}i"
    )
    assert len(items) == 2
    items[0].set("s", "1")
    items[1].set("s", "0")
    members[slicer_member] = ElementTree.tostring(
        slicer_cache, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_pivot_slicer_relationship(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "pivot_slicer_selection_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "xl/_rels/workbook.xml.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationship = next(item for item in relationships if item.get("Id") == "rIdWCABSlicerCache")
    relationship.set(
        "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    )
    members[relationships_member] = ElementTree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_pivot_slicer_selection_pair_changes_only_its_slicer_cache(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "pivot_slicer_selection_changed"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        assert baseline.testzip() is None
        assert candidate.testzip() is None
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/slicerCaches/slicerCache1.xml"]


def test_validator_rejects_a_false_power_query_m_filter_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "power_query_m_filter_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_filter_value"] = "North"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_power_query_m_filter_literal(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "power_query_m_filter_changed" / "candidate.xlsx"
    _replace_power_query_m_filter_literal(candidate, before=b'"South"', after=b'"North"')
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_power_query_root_relationship(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "power_query_m_filter_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "_rels/.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationship = next(item for item in relationships if item.get("Id") == "rIdWCABPowerQuery")
    relationship.set(
        "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    )
    members[relationships_member] = ElementTree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_power_query_m_filter_pair_changes_only_its_data_mashup_part(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "power_query_m_filter_changed"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        assert baseline.testzip() is None
        assert candidate.testzip() is None
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["customXml/item1.xml"]


def test_validator_rejects_a_false_chart_series_reference_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "chart_series_reference_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_value_ref"] = "'Source'!$B$2:$B$4"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_chart_series_value_reference(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "chart_series_reference_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    chart_member = "xl/charts/chart1.xml"
    chart = ElementTree.fromstring(members[chart_member])
    formula = chart.find(
        f".//{{{namespace}}}ser/{{{namespace}}}val/{{{namespace}}}numRef/{{{namespace}}}f"
    )
    assert formula is not None
    formula.text = "'Source'!$B$2:$B$4"
    members[chart_member] = ElementTree.tostring(chart, encoding="utf-8", xml_declaration=True)
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_chart_relationship_binding(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "chart_series_reference_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "xl/drawings/_rels/drawing1.xml.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationship = next(
        relationship for relationship in relationships if relationship.get("Id") == "rId1"
    )
    relationship.set(
        "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    )
    members[relationships_member] = ElementTree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_chart_series_reference_pair_changes_only_its_chart_part(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "chart_series_reference_changed"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        assert baseline.testzip() is None
        assert candidate.testzip() is None
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/charts/chart1.xml"]


def test_validator_rejects_a_false_external_workbook_link_policy_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_workbook_link_update_on_open"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_update_links"] = "never"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_external_workbook_link_policy(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "external_workbook_link_update_on_open" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b'updateLinks="always"', b'updateLinks="never"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_external_workbook_link_policy_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_workbook_link_update_on_open"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/workbook.xml"]


def test_validator_rejects_a_false_iterative_calculation_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "iterative_calculation_enabled"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["iteration_count"] = 99
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_iterative_calculation_control(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "iterative_calculation_enabled" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b'iterate="1"', b'iterate="0"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_iterative_calculation_control_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "iterative_calculation_enabled" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b'forceFullCalc="0"', b'forceFullCalc="1"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_iterative_calculation_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "iterative_calculation_enabled"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/workbook.xml"]


def test_validator_rejects_a_false_precision_as_displayed_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "precision_as_displayed_enabled"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["input_value"] = 10.01
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_precision_as_displayed_control(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "precision_as_displayed_enabled" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b'fullPrecision="0"', b'fullPrecision="1"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_precision_as_displayed_control_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "precision_as_displayed_enabled" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b'forceFullCalc="0"', b'forceFullCalc="1"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_precision_as_displayed_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "precision_as_displayed_enabled"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/workbook.xml"]


def test_validator_rejects_a_false_auto_filter_criteria_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "auto_filter_criteria_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_filter_value"] = "West"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_auto_filter_criterion(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "operations" / "auto_filter_criteria_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet = ElementTree.fromstring(members["xl/worksheets/sheet1.xml"])
    filter_element = next(worksheet.iter(f"{{{namespace}}}filter"))
    filter_element.set("val", "North")
    members["xl/worksheets/sheet1.xml"] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_auto_filter_control_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "operations" / "auto_filter_criteria_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet = ElementTree.fromstring(members["xl/worksheets/sheet1.xml"])
    filter_column = worksheet.find(f".//{{{namespace}}}filterColumn")
    filter_column.set("showButton", "0")
    members["xl/worksheets/sheet1.xml"] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_extra_auto_filter_control(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "operations" / "auto_filter_criteria_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet = ElementTree.fromstring(members["xl/worksheets/sheet1.xml"])
    filter_column = worksheet.find(f".//{{{namespace}}}filterColumn")
    assert filter_column is not None
    ElementTree.SubElement(filter_column, f"{{{namespace}}}top10", {"val": "1"})
    members["xl/worksheets/sheet1.xml"] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".staging.xlsx")
    with ZipFile(staging, "w", ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_auto_filter_subtotal_formula(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "operations" / "auto_filter_criteria_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet = ElementTree.fromstring(members["xl/worksheets/sheet1.xml"])
    subtotal_cell = next(
        cell for cell in worksheet.iter(f"{{{namespace}}}c") if cell.get("r") == "D2"
    )
    subtotal_cell.find(f"{{{namespace}}}f").text = "SUM(B2:B5)"
    members["xl/worksheets/sheet1.xml"] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_auto_filter_criteria_pair_changes_only_its_report_worksheet(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "auto_filter_criteria_changed"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/worksheets/sheet1.xml"]


def test_validator_rejects_a_false_workbook_date_system_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "workbook_date_system_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_date_1904"] = False
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_workbook_date_system_control(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "workbook_date_system_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    workbook = ElementTree.fromstring(members["xl/workbook.xml"])
    workbook.find(f"{{{namespace}}}workbookPr").set("date1904", "0")
    members["xl/workbook.xml"] = ElementTree.tostring(
        workbook, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_a_corrupted_workbook_date_system_serial(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "workbook_date_system_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    cell = next(cell for cell in worksheet.iter(f"{{{namespace}}}c") if cell.get("r") == "B2")
    cell.find(f"{{{namespace}}}v").text = "45293"
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_workbook_date_system_control_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "workbook_date_system_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    workbook = ElementTree.fromstring(members["xl/workbook.xml"])
    workbook.find(f"{{{namespace}}}workbookPr").set("showObjects", "none")
    members["xl/workbook.xml"] = ElementTree.tostring(
        workbook, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_workbook_date_system_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "workbook_date_system_changed"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/workbook.xml"]


def test_validator_rejects_a_false_formula_cached_result_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "formula_cached_result_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_cached_result"] = 24
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_formula_cached_result(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "formula_cached_result_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    worksheet_member = "xl/worksheets/sheet2.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    cell = next(cell for cell in worksheet.iter(f"{{{namespace}}}c") if cell.get("r") == "B2")
    cell.find(f"{{{namespace}}}v").text = "20"
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_formula_cached_result_formula_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "formula_cached_result_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    worksheet_member = "xl/worksheets/sheet2.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    cell = next(cell for cell in worksheet.iter(f"{{{namespace}}}c") if cell.get("r") == "B2")
    cell.find(f"{{{namespace}}}f").text = "Inputs!$B$2*3"
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_formula_cached_result_pair_changes_only_model_worksheet_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "formula_cached_result_changed"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/worksheets/sheet2.xml"]


def test_validator_rejects_a_false_array_formula_mode_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "array_formula_mode_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_mode"] = "legacy_cse"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_missing_dynamic_array_metadata_binding(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "array_formula_mode_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/metadata.xml"] = members["xl/metadata.xml"].replace(
        b'fDynamic="1"', b'fDynamic="0"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_formulafence_adapter_maps_a_pair_fact(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [{"kind": "formula_to_value", "location": "Revenue!C8"}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(fixture_root / "finance" / "formula_to_value")
    assert result["status"] == "matched"
    assert result["matched"] == ["formula_to_value"]


def test_formulafence_adapter_maps_a_structured_table_scope_fact(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [{"kind": "table_definition_changed", "location": None}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "structured_table_scope_expansion"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["structured_table_scope_changed"]


def test_formulafence_adapter_maps_dynamic_fact_and_coverage_expectation(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 2},
            "changes": [
                {
                    "kind": "dynamic_formula_reference_added",
                    "location": "Summary!B2",
                    "details": {"functions": ["INDIRECT"]},
                }
            ],
            "findings": [
                {
                    "rule_id": "FF012",
                    "location": "Summary!B2",
                    "details": {"functions": ["INDIRECT"]},
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "dynamic_reference_introduced"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["dynamic_formula_reference_added"]
    assert len(result["matched_coverage_expectations"]) == 1
    assert result["missed_coverage_expectations"] == []


def test_formulafence_adapter_maps_dynamic_driver_coverage_from_profile(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [{"kind": "value_changed", "location": "Inputs!E12"}],
        }

    def fake_profile(*_args, **_kwargs):
        return {
            "features": {
                "dynamic_reference_cells": [{"location": "Summary!B2", "functions": ["INDIRECT"]}]
            }
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    monkeypatch.setattr(formulafence, "profile", fake_profile)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "indirect_reference_driver_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["value_changed"]
    assert len(result["matched_coverage_expectations"]) == 1
    assert result["missed_coverage_expectations"] == []


def test_formulafence_adapter_maps_external_data_refresh_fact(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_data_connections_changed",
                    "location": None,
                    "details": {
                        "before": [{"id": 1, "refresh_on_load": False}],
                        "after": [{"id": 1, "refresh_on_load": True}],
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_data_refresh_on_open"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["external_data_connection_refresh_on_load_changed"]


def test_formulafence_adapter_requires_the_exact_external_data_refresh_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_data_connections_changed",
                    "location": None,
                    "details": {
                        "before": [{"id": 1, "refresh_on_load": False}],
                        "after": [{"id": 1, "refresh_on_load": False}],
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_data_refresh_on_open"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["external_data_connection_refresh_on_load_changed"]


def test_formulafence_adapter_maps_the_exact_external_workbook_link_policy(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_data_refresh_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "update_links": "never",
                            "allow_refresh_query": False,
                            "refresh_all_connections": False,
                            "save_external_link_values": True,
                        },
                        "after": {
                            "update_links": "always",
                            "allow_refresh_query": False,
                            "refresh_all_connections": False,
                            "save_external_link_values": True,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_workbook_link_update_on_open"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["external_workbook_link_update_policy_changed"]


def test_formulafence_adapter_rejects_an_inexact_external_workbook_link_policy(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_data_refresh_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "update_links": "never",
                            "allow_refresh_query": False,
                            "refresh_all_connections": False,
                            "save_external_link_values": True,
                        },
                        "after": {
                            "update_links": "always",
                            "allow_refresh_query": True,
                            "refresh_all_connections": False,
                            "save_external_link_values": True,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_workbook_link_update_on_open"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["external_workbook_link_update_policy_changed"]


def test_formulafence_adapter_maps_the_exact_iterative_calculation_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "calculation_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "iterate": False,
                            "iterateCount": 100,
                            "iterateDelta": 0.001,
                        },
                        "after": {
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "iterate": True,
                            "iterateCount": 100,
                            "iterateDelta": 0.001,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "iterative_calculation_enabled"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["iterative_calculation_enabled"]


def test_formulafence_adapter_rejects_an_inexact_iterative_calculation_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "calculation_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "iterate": False,
                            "iterateCount": 100,
                            "iterateDelta": 0.001,
                        },
                        "after": {
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "iterate": True,
                            "iterateCount": 100,
                            "iterateDelta": 0.01,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "iterative_calculation_enabled"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["iterative_calculation_enabled"]


def test_formulafence_adapter_maps_the_exact_precision_as_displayed_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "calculation_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "calcCompleted": True,
                            "calcMode": "auto",
                            "calcOnSave": True,
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "fullPrecision": True,
                        },
                        "after": {
                            "calcCompleted": True,
                            "calcMode": "auto",
                            "calcOnSave": True,
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "fullPrecision": False,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "precision_as_displayed_enabled"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["precision_as_displayed_enabled"]


def test_formulafence_adapter_rejects_an_inexact_precision_as_displayed_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "calculation_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "calcCompleted": True,
                            "calcMode": "auto",
                            "calcOnSave": True,
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "fullPrecision": True,
                        },
                        "after": {
                            "calcCompleted": True,
                            "calcMode": "auto",
                            "calcOnSave": False,
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "fullPrecision": False,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "precision_as_displayed_enabled"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["precision_as_displayed_enabled"]


def test_formulafence_adapter_maps_the_exact_workbook_date_system_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = {
            "before": {
                "date_1904": False,
                "date_compatibility": True,
                "date_compatibility_declared": True,
                "unrecognized_control_count": 0,
            },
            "after": {
                "date_1904": True,
                "date_compatibility": True,
                "date_compatibility_declared": True,
                "unrecognized_control_count": 0,
            },
        }
        return {
            "summary": {"change_count": 2},
            "changes": [
                {
                    "kind": "workbook_date_system_changed",
                    "location": None,
                    "details": details,
                },
                {"kind": "value_changed", "location": "Inputs!B2"},
            ],
            "findings": [{"rule_id": "FF117", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "workbook_date_system_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["workbook_date_system_changed"]


def test_formulafence_adapter_requires_the_workbook_date_system_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 2},
            "changes": [
                {
                    "kind": "workbook_date_system_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "date_1904": False,
                            "date_compatibility": True,
                            "date_compatibility_declared": True,
                            "unrecognized_control_count": 0,
                        },
                        "after": {
                            "date_1904": True,
                            "date_compatibility": True,
                            "date_compatibility_declared": True,
                            "unrecognized_control_count": 0,
                        },
                    },
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "workbook_date_system_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["workbook_date_system_changed"]


def test_formulafence_adapter_maps_the_exact_auto_filter_criteria_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _auto_filter_criteria_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "filter_visibility_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF036", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "auto_filter_criteria_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["auto_filter_criteria_changed"]


def test_formulafence_adapter_requires_the_auto_filter_finding(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "filter_visibility_controls_changed",
                    "location": None,
                    "details": _auto_filter_criteria_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "auto_filter_criteria_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["auto_filter_criteria_changed"]


def test_formulafence_adapter_maps_the_exact_pivot_cache_refresh_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _pivot_cache_refresh_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "pivot_cache_refresh_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF023", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "pivot_cache_refresh_on_open"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["pivot_cache_refresh_on_load_changed"]


def test_formulafence_adapter_requires_the_pivot_cache_refresh_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "pivot_cache_refresh_controls_changed",
                    "location": None,
                    "details": _pivot_cache_refresh_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "pivot_cache_refresh_on_open"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["pivot_cache_refresh_on_load_changed"]


def test_formulafence_adapter_maps_the_exact_pivot_data_field_aggregation_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _pivot_data_field_aggregation_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "pivot_table_definitions_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF031", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "pivot_data_field_aggregation_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["pivot_data_field_aggregation_changed"]


def test_formulafence_adapter_rejects_an_inexact_pivot_data_field_aggregation_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _pivot_data_field_aggregation_details()
        details["cache_record_payload_material_changed"] = True
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "pivot_table_definitions_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF031", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "pivot_data_field_aggregation_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["pivot_data_field_aggregation_changed"]


def test_formulafence_adapter_requires_the_pivot_data_field_aggregation_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "pivot_table_definitions_changed",
                    "location": None,
                    "details": _pivot_data_field_aggregation_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "pivot_data_field_aggregation_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["pivot_data_field_aggregation_changed"]


def test_formulafence_adapter_maps_the_exact_pivot_slicer_selection_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _pivot_slicer_selection_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "slicer_timeline_cache_definitions_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF032", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "pivot_slicer_selection_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["pivot_slicer_selection_changed"]


def test_formulafence_adapter_rejects_an_inexact_pivot_slicer_selection_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _pivot_slicer_selection_details()
        details["timeline_filter_state_or_definition_material_changed"] = True
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "slicer_timeline_cache_definitions_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF032", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "pivot_slicer_selection_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["pivot_slicer_selection_changed"]


def test_formulafence_adapter_requires_the_pivot_slicer_selection_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "slicer_timeline_cache_definitions_changed",
                    "location": None,
                    "details": _pivot_slicer_selection_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "pivot_slicer_selection_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["pivot_slicer_selection_changed"]


def test_formulafence_adapter_maps_the_exact_power_query_m_filter_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _power_query_m_filter_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "power_query_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF024", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "power_query_m_filter_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["power_query_m_filter_changed"]


def test_formulafence_adapter_rejects_an_inexact_power_query_m_filter_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _power_query_m_filter_details()
        details["metadata_control_material_changed"] = True
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "power_query_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF024", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "power_query_m_filter_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["power_query_m_filter_changed"]


def test_formulafence_adapter_requires_the_power_query_m_filter_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "power_query_changed",
                    "location": None,
                    "details": _power_query_m_filter_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "power_query_m_filter_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["power_query_m_filter_changed"]


def test_formulafence_adapter_maps_the_exact_chart_series_reference_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _chart_series_reference_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "chart_definitions_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF030", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "chart_series_reference_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["chart_series_value_reference_changed"]


def test_formulafence_adapter_rejects_an_inexact_chart_series_reference_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _chart_series_reference_details()
        details["cached_series_material_changed"] = True
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "chart_definitions_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF030", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "chart_series_reference_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["chart_series_value_reference_changed"]


def test_formulafence_adapter_requires_the_chart_series_reference_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "chart_definitions_changed",
                    "location": None,
                    "details": _chart_series_reference_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "chart_series_reference_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["chart_series_value_reference_changed"]


def test_formulafence_adapter_maps_the_exact_formula_cached_result_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        profile = {
            "present": True,
            "formula_cell_count": 2,
            "cached_result_cell_count": 1,
            "missing_cached_result_cell_count": 1,
            "numeric_cached_result_count": 1,
            "string_cached_result_count": 0,
            "boolean_cached_result_count": 0,
            "error_cached_result_count": 0,
            "unrecognized_cached_result_count": 0,
        }
        details = {
            "before": profile,
            "after": profile,
            "unexplained_cached_result_change_count": 1,
            "cached_result_material_changed": True,
        }
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "formula_cached_result_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF042", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "formula_cached_result_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["formula_cached_result_changed"]


def test_formulafence_adapter_rejects_an_explained_formula_cached_result_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        profile = {
            "present": True,
            "formula_cell_count": 2,
            "cached_result_cell_count": 1,
            "missing_cached_result_cell_count": 1,
            "numeric_cached_result_count": 1,
            "string_cached_result_count": 0,
            "boolean_cached_result_count": 0,
            "error_cached_result_count": 0,
            "unrecognized_cached_result_count": 0,
        }
        details = {
            "before": profile,
            "after": profile,
            "unexplained_cached_result_change_count": 0,
            "cached_result_material_changed": True,
        }
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "formula_cached_result_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF042", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "formula_cached_result_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["formula_cached_result_changed"]


def test_formulafence_adapter_requires_the_formula_cached_result_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        profile = {
            "present": True,
            "formula_cell_count": 2,
            "cached_result_cell_count": 1,
            "missing_cached_result_cell_count": 1,
            "numeric_cached_result_count": 1,
            "string_cached_result_count": 0,
            "boolean_cached_result_count": 0,
            "error_cached_result_count": 0,
            "unrecognized_cached_result_count": 0,
        }
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "formula_cached_result_changed",
                    "location": None,
                    "details": {
                        "before": profile,
                        "after": profile,
                        "unexplained_cached_result_change_count": 1,
                        "cached_result_material_changed": True,
                    },
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "formula_cached_result_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["formula_cached_result_changed"]


def test_formulafence_adapter_maps_the_exact_array_formula_mode_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "array_formula_mode_changed",
                    "location": "Model!B1",
                    "details": {
                        "before": {"mode": "legacy_cse", "output_range": "B1:B3"},
                        "after": {"mode": "dynamic", "output_range": "B1:B3"},
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "array_formula_mode_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["array_formula_mode_changed"]


def test_formulafence_adapter_rejects_an_inexact_array_formula_mode_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "array_formula_mode_changed",
                    "location": "Model!B1",
                    "details": {
                        "before": {"mode": "legacy_cse", "output_range": "B1:B3"},
                        "after": {"mode": "dynamic", "output_range": "B1:B4"},
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "array_formula_mode_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["array_formula_mode_changed"]


def test_formulafence_adapter_maps_a_portfolio_impact(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_portfolio(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "workbooks": [
                {
                    "path": "drivers.xlsx",
                    "changes": [{"kind": "value_changed", "location": "Inputs!B2"}],
                    "findings": [
                        {
                            "rule_id": "FF079",
                            "details": {"sample_impacts": [{"workbook": "model.xlsx"}]},
                        }
                    ],
                }
            ],
        }

    monkeypatch.setattr(formulafence, "portfolio", fake_portfolio)
    result = formulafence.evaluate_diff_case(
        fixture_root / "portfolio" / "external_driver_value_change"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["portfolio_value_changed", "portfolio_external_reference"]
