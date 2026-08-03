from __future__ import annotations

import base64
import io
import json
import struct
from hashlib import sha256
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from wcab.adapters import formulafence
from wcab.build import CASE_IDS, build_all
from wcab.manifest import ManifestError, case_rows, manifest_text
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


def _named_sheet_view_filter_criterion_details() -> dict[str, object]:
    profile = {
        "present": True,
        "worksheet_count": 1,
        "part_count": 1,
        "named_sheet_view_count": 1,
        "named_filter_count": 1,
        "column_filter_count": 1,
        "filter_criterion_count": 1,
        "sort_rule_count": 0,
        "sort_condition_count": 0,
        "unrecognized_named_sheet_view_count": 0,
    }
    return {
        "before": dict(profile),
        "after": dict(profile),
        "named_sheet_view_definition_material_changed": True,
    }


def _xml_map_table_xpath_details() -> dict[str, object]:
    profile = {
        "present": True,
        "xml_map_part_count": 1,
        "xml_schema_count": 1,
        "xml_map_count": 1,
        "xml_map_data_binding_count": 1,
        "xml_map_file_binding_count": 1,
        "xml_map_connection_binding_count": 1,
        "table_xml_binding_part_count": 1,
        "table_xml_binding_count": 1,
        "single_cell_xml_binding_sheet_count": 1,
        "single_cell_xml_binding_part_count": 1,
        "single_cell_xml_binding_count": 1,
        "single_cell_xml_connection_binding_count": 1,
        "unrecognized_xml_mapping_count": 0,
    }
    return {
        "before": dict(profile),
        "after": dict(profile),
        "xml_mapping_bindings_changed": True,
    }


def _office_web_addin_auto_show_details() -> dict[str, object]:
    before = {
        "present": True,
        "declared_taskpane_part_count": 1,
        "taskpane_part_count": 1,
        "web_extension_part_count": 1,
        "unrecognized_part_count": 0,
        "taskpane_count": 1,
        "visible_taskpane_count": 0,
        "locked_taskpane_count": 1,
        "web_extension_reference_count": 1,
        "auto_show_taskpane_count": 0,
        "store_reference_count": 1,
        "alternate_reference_count": 0,
        "binding_count": 0,
        "snapshot_reference_count": 0,
        "related_relationship_count": 1,
        "external_relationship_count": 0,
        "worksheet_binding_sheet_count": 0,
        "worksheet_binding_count": 0,
        "in_content_drawing_part_count": 0,
        "in_content_web_extension_reference_count": 0,
        "in_content_web_extension_part_count": 0,
    }
    return {
        "before": before,
        "after": {**before, "auto_show_taskpane_count": 1},
        "web_extension_definition_material_changed": True,
    }


def _ole_object_auto_load_details() -> dict[str, object]:
    before = {
        "present": True,
        "control_sheet_count": 1,
        "worksheet_control_count": 0,
        "active_x_part_count": 0,
        "active_x_binary_reference_count": 0,
        "form_control_property_part_count": 0,
        "legacy_vml_drawing_part_count": 0,
        "legacy_vml_control_count": 0,
        "legacy_vml_macro_assignment_count": 0,
        "legacy_vml_cell_link_count": 0,
        "legacy_vml_source_range_count": 0,
        "legacy_vml_camera_source_range_count": 0,
        "control_macro_assignment_count": 0,
        "control_cell_link_count": 0,
        "control_source_range_count": 0,
        "form_control_formula_binding_count": 0,
        "ole_object_count": 1,
        "linked_ole_object_count": 0,
        "auto_load_ole_object_count": 0,
        "auto_update_ole_object_count": 0,
        "related_relationship_count": 1,
        "external_relationship_count": 0,
        "internal_related_part_count": 1,
        "fingerprinted_related_part_count": 1,
        "uninspected_related_part_count": 0,
        "unrecognized_part_count": 0,
    }
    return {
        "before": before,
        "after": {**before, "auto_load_ole_object_count": 1},
        "worksheet_control_definition_material_changed": True,
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


def _query_table_refresh_details() -> dict[str, object]:
    before = {
        "sheet": "ImportedData",
        "connection_id": 1,
        "refresh_on_load": False,
        "background_refresh": False,
        "refresh_disabled": False,
        "remove_data_on_save": False,
        "fill_formulas": False,
        "connection_edit_disabled": True,
        "growth_behavior": "insert_clear",
        "has_name": True,
        "has_refresh_metadata": False,
        "opaque_metadata": {"present": False, "count": 0},
    }
    return {"before": [before], "after": [{**before, "refresh_on_load": True}]}


def _cell_hyperlink_target_details() -> dict[str, object]:
    profile = {
        "present": True,
        "worksheet_hyperlink_sheet_count": 1,
        "hyperlink_count": 1,
        "hyperlink_with_location_count": 0,
        "hyperlink_with_display_count": 0,
        "hyperlink_with_tooltip_count": 0,
        "binding_relationship_count": 1,
        "external_relationship_count": 1,
        "unrecognized_cell_hyperlink_count": 0,
    }
    return {
        "before": dict(profile),
        "after": dict(profile),
        "cell_hyperlink_binding_changed": True,
        "cell_hyperlink_definition_material_changed": True,
        "cell_hyperlink_relationships_changed": True,
    }


def _external_workbook_link_source_details() -> dict[str, object]:
    profile = {
        "present": True,
        "external_link_count": 1,
        "external_workbook_count": 1,
        "dde_link_count": 0,
        "ole_link_count": 0,
        "unrecognized_link_count": 0,
        "external_workbook_sheet_count": 1,
        "external_defined_name_count": 0,
        "external_workbook_cached_sheet_count": 0,
        "external_workbook_cached_cell_count": 0,
        "external_workbook_cached_refresh_error_count": 0,
        "dde_item_count": 0,
        "dde_advise_item_count": 0,
        "dde_ole_item_count": 0,
        "dde_prefer_picture_item_count": 0,
        "dde_cached_value_count": 0,
        "ole_item_count": 0,
        "ole_advise_item_count": 0,
        "ole_icon_item_count": 0,
        "ole_prefer_picture_item_count": 0,
        "opaque_metadata": {"present": False, "count": 0},
    }
    return {
        "before": dict(profile),
        "after": dict(profile),
        "source_material_changed": True,
    }


def _external_defined_name_source_details() -> dict[str, object]:
    profile = {
        "present": True,
        "surface_count": 1,
        "cell_formula_surface_count": 0,
        "defined_name_surface_count": 1,
        "data_validation_surface_count": 0,
        "chart_formula_surface_count": 0,
        "opaque_chart_part_count": 0,
        "external_reference_count": 1,
    }
    return {
        "before": dict(profile),
        "after": dict(profile),
        "external_workbook_link_surface_material_changed": True,
    }


def _external_defined_name_source_definition_change_details() -> dict[str, object]:
    return {
        "name": "ScenarioRate",
        "before": "'[WCABApprovedSource.xlsx]Inputs'!$B$2",
        "after": "'[WCABReviewSource.xlsx]Inputs'!$B$2",
    }


def _external_defined_name_source_definition_finding_details() -> dict[str, object]:
    return {
        "before": "'[WCABApprovedSource.xlsx]Inputs'!$B$2",
        "after": "'[WCABReviewSource.xlsx]Inputs'!$B$2",
    }


def _named_lambda_definition_change_details() -> dict[str, object]:
    return {
        "name": "ScenarioValue",
        "before": "=LAMBDA(rate,amount,rate*amount)",
        "after": "=LAMBDA(rate,amount,rate*(amount+10))",
    }


def _named_lambda_definition_finding_details() -> dict[str, object]:
    return {
        "before": "=LAMBDA(rate,amount,rate*amount)",
        "after": "=LAMBDA(rate,amount,rate*(amount+10))",
    }


def _table_calculated_column_formula_details() -> dict[str, object]:
    snapshot = {
        "name": "ScenarioLedger",
        "sheet": "Ledger",
        "ref": "A1:C4",
        "columns": ["Units", "Rate", "Calculated amount"],
        "header_row_count": 1,
        "totals_row_count": 0,
    }
    return {
        "before": dict(snapshot),
        "after": dict(snapshot),
        "calculated_column_formula_material_changed": True,
        "name": "ScenarioLedger",
    }


def _power_pivot_data_model_relationship_details() -> dict[str, object]:
    profile = {
        "present": True,
        "data_model_part_count": 1,
        "workbook_binding_count": 1,
        "data_model_declaration_count": 1,
        "model_table_count": 2,
        "model_relationship_count": 1,
        "related_relationship_count": 0,
        "external_relationship_count": 0,
        "fingerprinted_data_part_count": 1,
        "uninspected_data_part_count": 0,
        "unrecognized_part_count": 0,
    }
    return {
        "before": dict(profile),
        "after": dict(profile),
        "workbook_data_model_declaration_changed": True,
    }


def _xlm_auto_open_binding_details() -> dict[str, object]:
    profile = {
        "present": True,
        "automatic_macro_binding_count": 1,
        "auto_open_binding_count": 1,
        "auto_close_binding_count": 0,
        "auto_activate_binding_count": 0,
        "auto_deactivate_binding_count": 0,
    }
    return {
        "before": dict(profile),
        "after": dict(profile),
        "automatic_macro_binding_material_changed": True,
    }


def _sheet_protection_sort_permission_details() -> dict[str, object]:
    credential = {
        "configured": False,
        "has_legacy_verifier": False,
        "has_modern_verifier": False,
        "algorithm": None,
        "spin_count": None,
    }
    before = {
        "sheet": "Controls",
        "sheet_type": "worksheet",
        "enabled": True,
        "locked_actions": [
            "format_cells",
            "format_columns",
            "format_rows",
            "insert_columns",
            "insert_rows",
            "insert_hyperlinks",
            "delete_columns",
            "delete_rows",
            "sort",
            "auto_filter",
            "pivot_tables",
        ],
        "credential": credential,
        "opaque_metadata": {"present": False, "count": 0},
    }
    return {
        "sheet": "Controls",
        "before": before,
        "after": {
            **before,
            "locked_actions": [
                "format_cells",
                "format_columns",
                "format_rows",
                "insert_columns",
                "insert_rows",
                "insert_hyperlinks",
                "delete_columns",
                "delete_rows",
                "auto_filter",
                "pivot_tables",
            ],
        },
    }


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


def _scenario_manager_stored_input_details() -> dict[str, object]:
    profile = {
        "present": True,
        "scenario_sheet_count": 1,
        "scenario_count": 1,
        "input_cell_count": 2,
        "locked_scenario_count": 1,
        "hidden_scenario_count": 0,
        "scenario_with_comment_count": 1,
        "scenario_with_user_count": 1,
        "summary_reference_count": 1,
        "current_scenario_selection_count": 1,
        "shown_scenario_selection_count": 1,
        "deleted_input_cell_count": 0,
        "undone_input_cell_count": 0,
        "formatted_input_cell_count": 1,
        "unrecognized_scenario_count": 0,
    }
    return {
        "before": profile,
        "after": profile,
        "scenario_definition_material_changed": True,
    }


def _what_if_data_table_input_reference_details() -> dict[str, object]:
    profile = {
        "present": True,
        "data_table_count": 1,
        "one_variable_data_table_count": 1,
        "two_variable_data_table_count": 0,
        "one_variable_row_oriented_count": 0,
        "one_variable_column_oriented_count": 1,
        "declared_output_cell_count": 3,
        "recalculation_requested_count": 1,
        "deleted_input_reference_count": 0,
        "unrecognized_data_table_count": 0,
    }
    return {
        "before": profile,
        "after": profile,
        "data_table_definition_material_changed": True,
    }


def _data_validation_list_source_details() -> dict[str, object]:
    rule = {
        "allow_blank": False,
        "dropdown_hidden": False,
        "error": "Choose an approved status.",
        "error_style": "stop",
        "error_title": "Invalid status",
        "formula2": None,
        "ime_mode": "noControl",
        "operator": "between",
        "prompt": "Choose a documented status.",
        "prompt_title": "Approved status",
        "prompts_disabled": False,
        "ranges": ["Inputs!B2"],
        "sheet": "Inputs",
        "show_error_message": True,
        "show_input_message": False,
        "type": "list",
    }
    return {
        "sheet": "Inputs",
        "before": [{**rule, "formula1": "Lists!$A$2:$A$4"}],
        "after": [{**rule, "formula1": "Lists!$B$2:$B$4"}],
    }


def _conditional_formatting_threshold_details() -> dict[str, object]:
    rule = {
        "above_average": True,
        "bottom": False,
        "color_scale": None,
        "data_bar": None,
        "differential_style": {
            "element": "dxf",
            "children": [
                {
                    "element": "fill",
                    "children": [
                        {
                            "element": "patternFill",
                            "attributes": {"patternType": "solid"},
                            "children": [
                                {
                                    "element": "fgColor",
                                    "attributes": {"rgb": "FFFFC7CE"},
                                }
                            ],
                        }
                    ],
                }
            ],
        },
        "equal_average": False,
        "extensions": [],
        "icon_set": None,
        "operator": "greaterThan",
        "percent": False,
        "priority": 1,
        "ranges": ["Operations!B2:B4"],
        "rank": None,
        "sheet": "Operations",
        "std_dev": None,
        "stop_if_true": False,
        "text": None,
        "time_period": None,
        "type": "cellIs",
    }
    return {
        "sheet": "Operations",
        "before": {"rules": [{**rule, "formulas": ["100"]}], "extensions": []},
        "after": {"rules": [{**rule, "formulas": ["50"]}], "extensions": []},
    }


def _number_format_visibility_details() -> dict[str, object]:
    profile = {
        "present": True,
        "default_format_override_count": 0,
        "cell_format_assignment_count": 1,
        "row_format_assignment_count": 0,
        "column_format_assignment_count": 0,
        "built_in_format_assignment_count": 0,
        "custom_format_assignment_count": 1,
        "unrecognized_number_format_count": 0,
    }
    return {
        "before": dict(profile),
        "after": dict(profile),
        "number_format_definition_material_changed": True,
    }


def _ignored_error_formula_range_suppression_details() -> dict[str, object]:
    before = {
        "present": False,
        "worksheet_count": 0,
        "standard_container_count": 0,
        "extension_container_count": 0,
        "ignored_error_rule_count": 0,
        "target_range_count": 0,
        "evaluation_error_count": 0,
        "inconsistent_formula_count": 0,
        "formula_range_omission_count": 0,
        "unlocked_formula_count": 0,
        "empty_cell_reference_count": 0,
        "list_data_validation_count": 0,
        "calculated_column_count": 0,
        "number_stored_as_text_count": 0,
        "two_digit_text_year_count": 0,
        "unrecognized_ignored_error_count": 0,
    }
    return {
        "before": before,
        "after": {
            **before,
            "present": True,
            "worksheet_count": 1,
            "standard_container_count": 1,
            "ignored_error_rule_count": 1,
            "target_range_count": 1,
            "formula_range_omission_count": 1,
        },
        "ignored_error_definition_material_changed": True,
    }


def _workbook_structure_lock_removed_details() -> dict[str, object]:
    credential = {
        "configured": False,
        "has_legacy_verifier": False,
        "has_modern_verifier": False,
        "algorithm": None,
        "spin_count": None,
    }
    before = {
        "enabled": True,
        "lock_structure": True,
        "lock_windows": False,
        "lock_revision": False,
        "workbook_credential": credential,
        "revisions_credential": credential,
        "opaque_metadata": {"present": False, "count": 0},
    }
    return {"before": before, "after": {**before, "enabled": False, "lock_structure": False}}


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
    named_sheet_view_row = next(
        row for row in rows if row["id"] == "operations.named_sheet_view_filter_criterion_changed"
    )
    assert named_sheet_view_row["facts"] == [
        {
            "kind": "named_sheet_view_filter_criterion_changed",
            "sheet": "Report",
            "view_member": "xl/namedSheetViews/namedSheetView1.xml",
            "base_filter_ref": "A1:B5",
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
    xml_map_row = next(
        row for row in rows if row["id"] == "operations.xml_map_table_xpath_retargeted"
    )
    assert xml_map_row["facts"] == [
        {
            "kind": "xml_map_table_column_xpath_retargeted",
            "sheet": "Export",
            "table_member": "xl/tables/table1.xml",
            "table_name": "InvoiceLines",
            "table_ref": "A1:B3",
            "mapped_column_id": 2,
            "mapped_column_name": "Net amount",
            "map_member": "xl/xmlMaps.xml",
            "map_id": 1,
            "schema_id": "WCAB-INVOICE-EXPORT",
            "connection_id": 7,
            "baseline_xpath": "/wcab:Invoice/wcab:Line/wcab:NetAmount",
            "candidate_xpath": "/wcab:Invoice/wcab:Line/wcab:TaxAmount",
            "single_cell_member": "xl/singleCellTables/singleCellTable1.xml",
            "single_cell": "E2",
            "single_cell_xpath": "/wcab:Invoice/wcab:Header/wcab:AsOf",
            "total_cell": "D2",
            "total_formula": "=SUM(InvoiceLines[Net amount])",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Export!$D$2",
        }
    ]
    office_web_addin_row = next(
        row for row in rows if row["id"] == "governance.office_web_addin_auto_show_enabled"
    )
    assert office_web_addin_row["facts"] == [
        {
            "kind": "office_web_addin_auto_show_enabled",
            "taskpane_member": "xl/webextensions/taskpanes.xml",
            "web_extension_member": "xl/webextensions/webextension1.xml",
            "addin_id": "{33333333-3333-3333-3333-333333333333}",
            "reference_id": "{44444444-4444-4444-4444-444444444444}",
            "reference_version": "1.0.0.0",
            "store": "wcab-review-assistant.xml",
            "store_type": "Filesystem",
            "baseline_auto_show": False,
            "candidate_auto_show": True,
            "input_sheet": "Inputs",
            "input_cell": "B2",
            "input_value": 10,
            "model_sheet": "Model",
            "model_cell": "B2",
            "model_formula": "=Inputs!$B$2*2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Model!$B$2",
        }
    ]
    ole_object_row = next(
        row for row in rows if row["id"] == "governance.ole_object_auto_load_enabled"
    )
    assert ole_object_row["facts"] == [
        {
            "kind": "ole_object_auto_load_enabled",
            "sheet": "Inputs",
            "worksheet_member": "xl/worksheets/sheet1.xml",
            "worksheet_relationships_member": "xl/worksheets/_rels/sheet1.xml.rels",
            "relationship_id": "rIdWCABEmbeddedOle",
            "relationship_type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
            "target": "../embeddings/wcab-review-embedded-object.bin",
            "embedded_object_member": "xl/embeddings/wcab-review-embedded-object.bin",
            "content_type": "application/vnd.openxmlformats-officedocument.oleObject",
            "prog_id": "WCAB.Review.Embedded.Object",
            "dv_aspect": "DVASPECT_CONTENT",
            "shape_id": 1026,
            "baseline_auto_load": False,
            "candidate_auto_load": True,
            "input_sheet": "Inputs",
            "input_cell": "B2",
            "input_value": 10,
            "model_sheet": "Model",
            "model_cell": "B2",
            "model_formula": "=Inputs!$B$2*2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Model!$B$2",
        }
    ]
    data_validation_list_source_row = next(
        row for row in rows if row["id"] == "operations.data_validation_list_source_changed"
    )
    assert data_validation_list_source_row["facts"] == [
        {
            "kind": "data_validation_list_source_changed",
            "validation_sheet": "Inputs",
            "target_range": "B2",
            "validation_type": "list",
            "baseline_source_formula": "=Lists!$A$2:$A$4",
            "candidate_source_formula": "=Lists!$B$2:$B$4",
            "allow_blank": False,
            "dropdown_hidden": False,
            "show_error_message": True,
            "error_style": "stop",
            "error_title": "Invalid status",
            "error": "Choose an approved status.",
            "show_input_message": False,
            "prompt_title": "Approved status",
            "prompt": "Choose a documented status.",
            "source_sheet": "Lists",
            "baseline_source_range": "A2:A4",
            "candidate_source_range": "B2:B4",
            "baseline_source_values": ["Draft", "Review", "Approved"],
            "candidate_source_values": ["Draft", "Suspended", "Rejected"],
            "input_cell": "B2",
            "input_value": "Draft",
            "model_sheet": "Model",
            "model_cell": "B2",
            "model_formula": "=Inputs!$B$2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Model!$B$2",
        }
    ]
    conditional_formatting_threshold_row = next(
        row for row in rows if row["id"] == "operations.conditional_formatting_threshold_changed"
    )
    assert conditional_formatting_threshold_row["facts"] == [
        {
            "kind": "conditional_formatting_threshold_changed",
            "sheet": "Operations",
            "target_range": "B2:B4",
            "priority": 1,
            "rule_type": "cellIs",
            "operator": "greaterThan",
            "baseline_formula": "100",
            "candidate_formula": "50",
            "metric_values": [10, 75, 120],
            "fill_rgb": "FFFFC7CE",
        }
    ]
    number_format_visibility_row = next(
        row for row in rows if row["id"] == "operations.number_format_value_hidden"
    )
    assert number_format_visibility_row["facts"] == [
        {
            "kind": "cell_number_format_changed",
            "sheet": "Operations",
            "cell": "B2",
            "value": 0.125,
            "custom_number_format_id": 164,
            "baseline_format": "0.0%;[Red](0.0%);-",
            "candidate_format": ";;;",
            "formula_cell": "B3",
            "formula": "=B2",
        }
    ]
    ignored_error_suppression_row = next(
        row for row in rows if row["id"] == "operations.ignored_error_formula_range_suppressed"
    )
    assert ignored_error_suppression_row["facts"] == [
        {
            "kind": "ignored_error_rule_added",
            "sheet": "Operations",
            "target_range": "B5",
            "warning_flag": "formulaRange",
            "formula": "=SUM(B2:B3)",
            "adjacent_populated_cell": "B4",
            "adjacent_populated_value": 30,
            "downstream_formula_cell": "C5",
            "downstream_formula": "=B5",
        }
    ]
    workbook_structure_protection_row = next(
        row for row in rows if row["id"] == "governance.workbook_structure_lock_removed"
    )
    assert workbook_structure_protection_row["facts"] == [
        {
            "kind": "workbook_structure_lock_removed",
            "baseline_lock_structure": True,
            "candidate_lock_structure": False,
            "hidden_sheet": "ReviewControls",
            "hidden_sheet_state": "hidden",
            "formula_sheet": "Inputs",
            "formula_cell": "D2",
            "formula": "=B2*C2",
        }
    ]
    sheet_protection_sort_row = next(
        row for row in rows if row["id"] == "governance.sheet_protection_sort_permission_enabled"
    )
    assert sheet_protection_sort_row["facts"] == [
        {
            "kind": "sheet_protection_sort_permission_enabled",
            "sheet": "Controls",
            "worksheet_member": "xl/worksheets/sheet1.xml",
            "baseline_sort_locked": True,
            "candidate_sort_locked": False,
            "formula_cell": "D2",
            "formula": "=B2*C2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Controls!$D$2",
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
    query_table_row = next(
        row for row in rows if row["id"] == "governance.query_table_refresh_on_open"
    )
    assert query_table_row["facts"] == [
        {
            "kind": "query_table_refresh_on_load_changed",
            "sheet": "ImportedData",
            "connection_id": 1,
            "connection_member": "xl/connections.xml",
            "connection_url": "https://example.invalid/wcab-query-table-refresh",
            "query_table_member": "xl/queryTables/queryTable1.xml",
            "worksheet_member": "xl/worksheets/sheet1.xml",
            "worksheet_relationships_member": "xl/worksheets/_rels/sheet1.xml.rels",
            "relationship_id": "rIdWCABQueryTable",
            "relationship_type": (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
            ),
            "baseline_refresh_on_load": False,
            "candidate_refresh_on_load": True,
            "background_refresh": False,
            "refresh_disabled": False,
            "remove_data_on_save": False,
            "fill_formulas": False,
            "connection_edit_disabled": True,
            "growth_behavior": "insertClear",
            "saved_value_cell": "B2",
            "saved_value": 100,
            "summary_sheet": "Summary",
            "summary_cell": "B2",
            "summary_formula": "=ImportedData!$B$2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Summary!$B$2",
        }
    ]
    hyperlink_row = next(
        row for row in rows if row["id"] == "governance.cell_hyperlink_target_changed"
    )
    assert hyperlink_row["facts"] == [
        {
            "kind": "cell_hyperlink_target_changed",
            "sheet": "Inputs",
            "cell": "B2",
            "cell_value": "Open vendor portal",
            "worksheet_member": "xl/worksheets/sheet1.xml",
            "worksheet_relationships_member": "xl/worksheets/_rels/sheet1.xml.rels",
            "relationship_id": "rIdWCABVendorPortal",
            "relationship_type": (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            ),
            "target_mode": "External",
            "baseline_target": "https://approved.example.invalid/wcab-vendor-portal",
            "candidate_target": "https://review.example.invalid/wcab-vendor-portal",
            "summary_sheet": "Summary",
            "summary_cell": "B2",
            "summary_formula": "=Inputs!$B$2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Summary!$B$2",
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
    scenario_manager_row = next(
        row for row in rows if row["id"] == "structural.scenario_manager_stored_input_changed"
    )
    assert scenario_manager_row["facts"] == [
        {
            "kind": "scenario_manager_stored_input_value_changed",
            "scenario_sheet": "Inputs",
            "scenario_name": "WCAB downside",
            "changing_cell": "B2",
            "stable_input_cell": "B3",
            "baseline_stored_value": "0.08",
            "candidate_stored_value": "0.16",
            "stable_stored_value": "125",
            "input_number_format_id": 10,
            "summary_ref": "D2",
            "worksheet_input_value": 0.1,
            "worksheet_stable_input_value": 125,
            "result_cell": "D2",
            "result_formula": "=B2*B3",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Inputs!$D$2",
        }
    ]
    data_table_row = next(
        row for row in rows if row["id"] == "structural.what_if_data_table_input_reference_changed"
    )
    assert data_table_row["facts"] == [
        {
            "kind": "what_if_data_table_input_reference_changed",
            "table_sheet": "Sensitivity",
            "master_cell": "D3",
            "output_range": "D3:D5",
            "baseline_input_cell": "B2",
            "candidate_input_cell": "B3",
            "orientation": "column",
            "recalculation_requested": True,
            "input_value_range": "C3:C5",
            "input_values": [0.04, 0.08, 0.12],
            "primary_input_value": 0.08,
            "alternate_input_value": 0.12,
            "scale_cell": "B4",
            "scale_value": 100,
            "output_formula_cell": "D2",
            "output_formula": "=Model!$B$2",
            "model_sheet": "Model",
            "model_cell": "B2",
            "model_formula": "=Sensitivity!$B$2*Sensitivity!$B$3*Sensitivity!$B$4",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Model!$B$2",
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
    external_link_source_row = next(
        row for row in rows if row["id"] == "governance.external_workbook_link_source_changed"
    )
    assert external_link_source_row["facts"] == [
        {
            "kind": "external_workbook_link_source_changed",
            "sheet": "LinkedModel",
            "cell": "B2",
            "formula": "='[WCABSource.xlsx]Inputs'!$B$2",
            "workbook_member": "xl/workbook.xml",
            "workbook_relationships_member": "xl/_rels/workbook.xml.rels",
            "workbook_relationship_id": "rIdWCABExternalLink",
            "workbook_relationship_type": (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
            ),
            "external_link_member": "xl/externalLinks/externalLink1.xml",
            "external_link_relationships_member": ("xl/externalLinks/_rels/externalLink1.xml.rels"),
            "external_link_content_type": (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
            ),
            "external_link_relationship_id": "rIdWCABExternalLinkPath",
            "external_link_relationship_type": (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
            ),
            "external_sheet": "Inputs",
            "target_mode": "External",
            "baseline_target": (
                "https://approved.example.invalid/wcab-external-workbook/WCABSource.xlsx"
            ),
            "candidate_target": (
                "https://review.example.invalid/wcab-external-workbook/WCABSource.xlsx"
            ),
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=LinkedModel!$B$2",
        }
    ]
    external_defined_name_source_row = next(
        row for row in rows if row["id"] == "governance.external_defined_name_source_changed"
    )
    assert external_defined_name_source_row["facts"] == [
        {
            "kind": "external_defined_name_source_changed",
            "name": "ScenarioRate",
            "workbook_member": "xl/workbook.xml",
            "baseline_refers_to": "'[WCABApprovedSource.xlsx]Inputs'!$B$2",
            "candidate_refers_to": "'[WCABReviewSource.xlsx]Inputs'!$B$2",
            "formula_sheet": "Model",
            "formula_cell": "B2",
            "formula": "=ScenarioRate*2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Model!$B$2",
        }
    ]
    named_lambda_row = next(
        row for row in rows if row["id"] == "structural.named_lambda_definition_changed"
    )
    assert named_lambda_row["facts"] == [
        {
            "kind": "named_lambda_definition_changed",
            "name": "ScenarioValue",
            "workbook_member": "xl/workbook.xml",
            "parameters": ["rate", "amount"],
            "baseline_refers_to": "=LAMBDA(rate,amount,rate*amount)",
            "candidate_refers_to": "=LAMBDA(rate,amount,rate*(amount+10))",
            "input_sheet": "Inputs",
            "rate_cell": "B2",
            "rate_value": 0.08,
            "amount_cell": "B3",
            "amount_value": 100,
            "formula_sheet": "Model",
            "formula_cell": "B2",
            "formula": "=ScenarioValue(Inputs!B2,Inputs!B3)",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Model!$B$2",
        }
    ]
    power_pivot_data_model_row = next(
        row for row in rows if row["id"] == "structural.power_pivot_data_model_relationship_changed"
    )
    assert power_pivot_data_model_row["facts"] == [
        {
            "kind": "power_pivot_data_model_relationship_changed",
            "workbook_member": "xl/workbook.xml",
            "workbook_relationships_member": "xl/_rels/workbook.xml.rels",
            "data_model_member": "xl/model/item.data",
            "workbook_relationship_id": "rIdWCABPowerPivotData",
            "workbook_relationship_type": (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData"
            ),
            "workbook_relationship_target": "model/item.data",
            "data_model_content_type": "application/vnd.openxmlformats-officedocument.model+data",
            "extension_uri": "{FCE2AD5D-F65C-4FA6-A056-5C36A1767C68}",
            "min_version_load": "5",
            "model_tables": ["SalesModel", "CalendarModel"],
            "from_table": "SalesModel",
            "from_column": "CalendarKey",
            "to_table": "CalendarModel",
            "baseline_to_column": "DateKey",
            "candidate_to_column": "FiscalDateKey",
            "data_model_payload_sha256": (
                "97fa2a7df8643711934e7cbfc44b360d9faa0f8000211536d24d7454b1354302"
            ),
            "data_model_payload_size": 86,
        }
    ]
    xlm_auto_open_row = next(
        row for row in rows if row["id"] == "governance.xlm_auto_open_binding_retargeted"
    )
    assert xlm_auto_open_row["baseline_files"][0]["path"] == (
        "governance/xlm_auto_open_binding_retargeted/baseline.xlsm"
    )
    assert xlm_auto_open_row["candidate_files"][0]["path"] == (
        "governance/xlm_auto_open_binding_retargeted/candidate.xlsm"
    )
    assert xlm_auto_open_row["facts"] == [
        {
            "kind": "xlm_auto_open_binding_retargeted",
            "workbook_member": "xl/workbook.xml",
            "workbook_relationships_member": "xl/_rels/workbook.xml.rels",
            "macro_sheet_member": "xl/macrosheets/sheet1.xml",
            "macro_sheet_relationship_id": "rIdWCABXlmMacroSheet",
            "macro_sheet_relationship_type": (
                "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"
            ),
            "macro_sheet_relationship_target": "macrosheets/sheet1.xml",
            "macro_sheet_content_type": "application/vnd.ms-excel.macrosheet+xml",
            "workbook_content_type": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "macro_sheet_name": "Macro Automation",
            "macro_sheet_sheet_id": "4",
            "macro_sheet_state": "veryHidden",
            "automatic_macro_name": "_xlnm.Auto_Open",
            "automatic_macro_event": "Auto_Open",
            "baseline_target": "'Macro Automation'!$A$1",
            "candidate_target": "'Macro Automation'!$A$2",
            "macro_sheet_formula": "HALT()",
            "macro_sheet_formula_cells": ["A1", "A2"],
            "macro_sheet_sha256": (
                "41c42af5521da4dd51a8d5a0a271e4ef0fd82040de229fe2461333dfd10d8ba7"
            ),
            "macro_sheet_size": 234,
            "input_sheet": "Inputs",
            "input_cell": "B2",
            "input_value": 10,
            "model_sheet": "Model",
            "model_cell": "B2",
            "model_formula": "=Inputs!$B$2*2",
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=Model!$B$2",
        }
    ]
    table_calculated_column_row = next(
        row for row in rows if row["id"] == "structural.table_calculated_column_formula_changed"
    )
    assert table_calculated_column_row["facts"] == [
        {
            "kind": "table_calculated_column_formula_changed",
            "table_sheet": "Ledger",
            "table_member": "xl/tables/table1.xml",
            "table": "ScenarioLedger",
            "table_ref": "A1:C4",
            "calculated_column_id": 3,
            "calculated_column_name": "Calculated amount",
            "baseline_formula": "A2*B2",
            "candidate_formula": "A2*(B2+1)",
            "stable_formula_cells": ["C2", "C3", "C4"],
            "dashboard_sheet": "Dashboard",
            "dashboard_cell": "B4",
            "dashboard_formula": "=SUM(ScenarioLedger[Calculated amount])",
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


def test_manifest_rejects_a_mixed_extension_workbook_pair(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "xlm_auto_open_binding_retargeted"
    (case / "candidate.xlsm").replace(case / "candidate.xlsx")
    with pytest.raises(ManifestError, match="matching baseline/candidate"):
        case_rows(fixture_root, expected_ids=CASE_IDS)


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


def test_validator_rejects_a_false_query_table_refresh_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "query_table_refresh_on_open"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_refresh_on_load"] = False
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_query_table_refresh_flag(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "query_table_refresh_on_open" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    query_table_member = "xl/queryTables/queryTable1.xml"
    query_table = ElementTree.fromstring(members[query_table_member])
    assert query_table.tag == f"{{{namespace}}}queryTable"
    query_table.set("refreshOnLoad", "0")
    members[query_table_member] = ElementTree.tostring(
        query_table, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_query_table_control_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "query_table_refresh_on_open" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    query_table_member = "xl/queryTables/queryTable1.xml"
    query_table = ElementTree.fromstring(members[query_table_member])
    assert query_table.tag == f"{{{namespace}}}queryTable"
    query_table.set("disableEdit", "0")
    members[query_table_member] = ElementTree.tostring(
        query_table, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_query_table_refresh_pair_changes_only_its_query_table_part(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "query_table_refresh_on_open"
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
    ] == ["xl/queryTables/queryTable1.xml"]


def test_validator_rejects_a_false_cell_hyperlink_target_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "cell_hyperlink_target_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_target"] = "https://approved.example.invalid/wcab-vendor-portal"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_cell_hyperlink_target(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "cell_hyperlink_target_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "xl/worksheets/_rels/sheet1.xml.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationship = relationships[0]
    relationship.set("Target", "https://approved.example.invalid/wcab-vendor-portal")
    members[relationships_member] = ElementTree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_invalid_cell_hyperlink_relationship_mode(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "cell_hyperlink_target_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "xl/worksheets/_rels/sheet1.xml.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationships[0].set("TargetMode", "Internal")
    members[relationships_member] = ElementTree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_cell_hyperlink_target_pair_changes_only_its_relationship_target(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "cell_hyperlink_target_changed"
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
    ] == ["xl/worksheets/_rels/sheet1.xml.rels"]


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


def test_validator_rejects_a_false_data_validation_list_source_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "data_validation_list_source_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_source_formula"] = "=Lists!$A$2:$A$4"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_data_validation_list_source(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "operations" / "data_validation_list_source_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    formula = next(worksheet.iter(f"{{{namespace}}}formula1"))
    formula.text = "=Lists!$A$2:$A$4"
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_data_validation_control_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "operations" / "data_validation_list_source_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    validation = next(worksheet.iter(f"{{{namespace}}}dataValidation"))
    validation.set("showErrorMessage", "0")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_data_validation_list_source_pair_changes_only_its_inputs_worksheet(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "data_validation_list_source_changed"
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
    ] == ["xl/worksheets/sheet1.xml"]


def test_validator_rejects_a_false_conditional_formatting_threshold_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "conditional_formatting_threshold_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_formula"] = "100"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_conditional_formatting_threshold(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "operations" / "conditional_formatting_threshold_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    formula = next(worksheet.iter(f"{{{namespace}}}formula"))
    formula.text = "100"
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_conditional_formatting_control_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "operations" / "conditional_formatting_threshold_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    rule = next(worksheet.iter(f"{{{namespace}}}cfRule"))
    rule.set("operator", "lessThan")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_conditional_formatting_threshold_pair_changes_only_its_operations_worksheet(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "conditional_formatting_threshold_changed"
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
    ] == ["xl/worksheets/sheet1.xml"]


def test_validator_rejects_a_false_number_format_visibility_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "number_format_value_hidden"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_format"] = "0.0%"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_number_format_visibility_code(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "operations" / "number_format_value_hidden" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    styles_member = "xl/styles.xml"
    styles = ElementTree.fromstring(members[styles_member])
    number_format = next(styles.iter(f"{{{namespace}}}numFmt"))
    number_format.set("formatCode", "0.0%")
    members[styles_member] = ElementTree.tostring(styles, encoding="utf-8", xml_declaration=True)
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_number_format_style_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "operations" / "number_format_value_hidden" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    styles_member = "xl/styles.xml"
    styles = ElementTree.fromstring(members[styles_member])
    number_format_xf = next(
        xf
        for xf in styles.findall(f"./{{{namespace}}}cellXfs/{{{namespace}}}xf")
        if xf.get("numFmtId") == "164"
    )
    number_format_xf.set("applyNumberFormat", "0")
    members[styles_member] = ElementTree.tostring(styles, encoding="utf-8", xml_declaration=True)
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_number_format_visibility_pair_changes_only_styles(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "number_format_value_hidden"
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
    ] == ["xl/styles.xml"]


def test_validator_rejects_a_false_ignored_error_suppression_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "ignored_error_formula_range_suppressed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["warning_flag"] = "formula"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_ignored_error_suppression(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "operations" / "ignored_error_formula_range_suppressed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    rule = next(worksheet.iter(f"{{{namespace}}}ignoredError"))
    rule.set("formulaRange", "0")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_ignored_error_control_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "operations" / "ignored_error_formula_range_suppressed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    rule = next(worksheet.iter(f"{{{namespace}}}ignoredError"))
    rule.set("emptyCellReference", "1")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_ignored_error_suppression_pair_changes_only_its_operations_worksheet(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "ignored_error_formula_range_suppressed"
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
    ] == ["xl/worksheets/sheet1.xml"]


def test_validator_rejects_a_false_workbook_structure_lock_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "workbook_structure_lock_removed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_lock_structure"] = True
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_workbook_structure_lock(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "workbook_structure_lock_removed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    workbook_member = "xl/workbook.xml"
    workbook = ElementTree.fromstring(members[workbook_member])
    protection = workbook.find(f"{{{namespace}}}workbookProtection")
    assert protection is not None
    protection.set("lockStructure", "1")
    members[workbook_member] = ElementTree.tostring(
        workbook, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_workbook_protection_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "workbook_structure_lock_removed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    workbook_member = "xl/workbook.xml"
    workbook = ElementTree.fromstring(members[workbook_member])
    protection = workbook.find(f"{{{namespace}}}workbookProtection")
    assert protection is not None
    protection.set("lockWindows", "1")
    members[workbook_member] = ElementTree.tostring(
        workbook, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_workbook_structure_protection_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "workbook_structure_lock_removed"
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
    ] == ["xl/workbook.xml"]


def test_validator_rejects_a_false_scenario_manager_stored_input_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "scenario_manager_stored_input_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_stored_value"] = "0.08"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_scenario_manager_stored_input(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "structural" / "scenario_manager_stored_input_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    input_cell = next(
        cell for cell in worksheet.iter(f"{{{namespace}}}inputCells") if cell.get("r") == "B2"
    )
    input_cell.set("val", "0.08")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_scenario_manager_metadata_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "structural" / "scenario_manager_stored_input_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    scenario = next(worksheet.iter(f"{{{namespace}}}scenario"))
    scenario.set("locked", "0")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_scenario_manager_stored_input_pair_changes_only_its_inputs_worksheet(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "scenario_manager_stored_input_changed"
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
    ] == ["xl/worksheets/sheet1.xml"]


def test_validator_rejects_a_false_what_if_data_table_input_reference_fact(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "what_if_data_table_input_reference_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_input_cell"] = "B2"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_what_if_data_table_input_reference(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root
        / "structural"
        / "what_if_data_table_input_reference_changed"
        / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    formula = next(
        item for item in worksheet.iter(f"{{{namespace}}}f") if item.get("t") == "dataTable"
    )
    formula.set("r1", "B2")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_what_if_data_table_control_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root
        / "structural"
        / "what_if_data_table_input_reference_changed"
        / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    formula = next(
        item for item in worksheet.iter(f"{{{namespace}}}f") if item.get("t") == "dataTable"
    )
    formula.attrib.pop("ca")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_what_if_data_table_input_reference_pair_changes_only_its_sensitivity_worksheet(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "what_if_data_table_input_reference_changed"
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
    ] == ["xl/worksheets/sheet1.xml"]


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


def test_validator_rejects_a_false_external_workbook_link_source_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_workbook_link_source_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_target"] = (
        "https://approved.example.invalid/wcab-external-workbook/WCABSource.xlsx"
    )
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_external_workbook_link_source_target(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "external_workbook_link_source_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "xl/externalLinks/_rels/externalLink1.xml.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationships[0].set(
        "Target",
        "https://approved.example.invalid/wcab-external-workbook/WCABSource.xlsx",
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


def test_validator_rejects_an_internal_external_workbook_link_source_target(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "external_workbook_link_source_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    relationships_member = "xl/externalLinks/_rels/externalLink1.xml.rels"
    relationships = ElementTree.fromstring(members[relationships_member])
    relationships[0].set("TargetMode", "Internal")
    members[relationships_member] = ElementTree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_external_workbook_link_source_pair_changes_only_its_relationship_target(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_workbook_link_source_changed"
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
    ] == ["xl/externalLinks/_rels/externalLink1.xml.rels"]


def test_validator_rejects_a_false_external_defined_name_source_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_defined_name_source_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_refers_to"] = "'[WCABApprovedSource.xlsx]Inputs'!$B$2"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_external_defined_name_source(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "external_defined_name_source_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b"WCABReviewSource.xlsx", b"WCABApprovedSource.xlsx", 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_external_reference_declaration_in_defined_name_pair(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "external_defined_name_source_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b"</workbook>", b"<externalReferences/></workbook>", 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_external_defined_name_source_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_defined_name_source_changed"
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


def test_validator_rejects_a_false_named_lambda_definition_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "named_lambda_definition_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_refers_to"] = "=LAMBDA(rate,amount,rate*amount)"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_named_lambda_definition(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "named_lambda_definition_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b"rate*(amount+10)", b"rate*amount", 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_external_reference_declaration_in_named_lambda_pair(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "named_lambda_definition_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b"</workbook>", b"<externalReferences/></workbook>", 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_named_lambda_definition_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "named_lambda_definition_changed"
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


def test_validator_rejects_a_false_power_pivot_data_model_relationship_fact(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "power_pivot_data_model_relationship_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_to_column"] = "DateKey"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_power_pivot_data_model_relationship(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root
        / "structural"
        / "power_pivot_data_model_relationship_changed"
        / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b' toColumn="FiscalDateKey"', b' toColumn="DateKey"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_power_pivot_data_model_payload_drift(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root
        / "structural"
        / "power_pivot_data_model_relationship_changed"
        / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/model/item.data"] = b"WCAB unexpected opaque Data Model payload drift."
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_power_pivot_data_model_declaration_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root
        / "structural"
        / "power_pivot_data_model_relationship_changed"
        / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    workbook = ElementTree.fromstring(members["xl/workbook.xml"])
    namespace = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
    sales_model = next(
        table
        for table in workbook.iter(f"{{{namespace}}}modelTable")
        if table.get("name") == "SalesModel"
    )
    sales_model.set("connection", "UnexpectedConnection")
    members["xl/workbook.xml"] = ElementTree.tostring(
        workbook, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_power_pivot_data_model_relationship_pair_changes_only_workbook_xml(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "power_pivot_data_model_relationship_changed"
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


def test_validator_rejects_a_false_xlm_auto_open_binding_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "xlm_auto_open_binding_retargeted"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_target"] = "'Macro Automation'!$A$1"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_xlm_auto_open_binding(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "xlm_auto_open_binding_retargeted" / "candidate.xlsm"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b"'Macro Automation'!$A$2", b"'Macro Automation'!$A$1", 1
    )
    staging = candidate.with_suffix(".corrupt.xlsm")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_xlm_auto_open_macro_sheet_payload_drift(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "xlm_auto_open_binding_retargeted" / "candidate.xlsm"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/macrosheets/sheet1.xml"] = members["xl/macrosheets/sheet1.xml"].replace(
        b"HALT()", b"HALT( )", 1
    )
    staging = candidate.with_suffix(".corrupt.xlsm")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_xlm_auto_open_declaration_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "xlm_auto_open_binding_retargeted" / "candidate.xlsm"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    workbook = ElementTree.fromstring(members["xl/workbook.xml"])
    sheet_tag = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"
    macro_sheet = next(
        sheet for sheet in workbook.iter(sheet_tag) if sheet.get("name") == "Macro Automation"
    )
    macro_sheet.set("state", "hidden")
    members["xl/workbook.xml"] = ElementTree.tostring(
        workbook, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsm")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_xlm_auto_open_binding_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "xlm_auto_open_binding_retargeted"
    with ZipFile(case / "baseline.xlsm") as baseline, ZipFile(case / "candidate.xlsm") as candidate:
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


def test_validator_rejects_a_false_table_calculated_column_formula_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "table_calculated_column_formula_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_formula"] = "A2*B2"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_table_calculated_column_formula(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "structural" / "table_calculated_column_formula_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/tables/table1.xml"] = members["xl/tables/table1.xml"].replace(
        b"A2*(B2+1)", b"A2*B2", 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_table_calculated_column_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "structural" / "table_calculated_column_formula_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    table = ElementTree.fromstring(members["xl/tables/table1.xml"])
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    auto_filter = table.find(f"{{{namespace}}}autoFilter")
    assert auto_filter is not None
    auto_filter.set("ref", "A1:C3")
    members["xl/tables/table1.xml"] = ElementTree.tostring(
        table, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_table_calculated_column_formula_pair_changes_only_its_table_part(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "table_calculated_column_formula_changed"
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
    ] == ["xl/tables/table1.xml"]


def test_validator_rejects_a_false_sheet_protection_sort_permission_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "sheet_protection_sort_permission_enabled"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_sort_locked"] = True
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_sheet_protection_sort_permission(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "sheet_protection_sort_permission_enabled" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/worksheets/sheet1.xml"] = members["xl/worksheets/sheet1.xml"].replace(
        b'sort="0"', b'sort="1"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_sheet_protection_action_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "sheet_protection_sort_permission_enabled" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/worksheets/sheet1.xml"] = members["xl/worksheets/sheet1.xml"].replace(
        b'autoFilter="1"', b'autoFilter="0"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_sheet_protection_sort_permission_pair_changes_only_controls_worksheet(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "sheet_protection_sort_permission_enabled"
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


def test_validator_rejects_a_false_named_sheet_view_filter_criterion_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "named_sheet_view_filter_criterion_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_filter_value"] = "West"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_named_sheet_view_filter_criterion(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "operations" / "named_sheet_view_filter_criterion_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    view_member = "xl/namedSheetViews/namedSheetView1.xml"
    named_views = ElementTree.fromstring(members[view_member])
    criterion = next(named_views.iter(f"{{{namespace}}}filter"))
    criterion.set("val", "North")
    members[view_member] = ElementTree.tostring(named_views, encoding="utf-8", xml_declaration=True)
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_named_sheet_view_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "operations" / "named_sheet_view_filter_criterion_changed" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"
    view_member = "xl/namedSheetViews/namedSheetView1.xml"
    named_views = ElementTree.fromstring(members[view_member])
    named_view = named_views.find(f"{{{namespace}}}namedSheetView")
    assert named_view is not None
    named_view.set("name", "Unexpected view")
    members[view_member] = ElementTree.tostring(named_views, encoding="utf-8", xml_declaration=True)
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_named_sheet_view_filter_criterion_pair_changes_only_its_view_part(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "named_sheet_view_filter_criterion_changed"
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
    ] == ["xl/namedSheetViews/namedSheetView1.xml"]


def test_validator_rejects_a_false_xml_map_table_xpath_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "xml_map_table_xpath_retargeted"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_xpath"] = "/wcab:Invoice/wcab:Line/wcab:GrossAmount"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_xml_map_table_xpath(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "operations" / "xml_map_table_xpath_retargeted" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    table = ElementTree.fromstring(members["xl/tables/table1.xml"])
    binding = table.find(f".//{{{namespace}}}xmlColumnPr")
    assert binding is not None
    binding.set("xpath", "/wcab:Invoice/wcab:Line/wcab:NetAmount")
    members["xl/tables/table1.xml"] = ElementTree.tostring(
        table, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_xml_map_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "operations" / "xml_map_table_xpath_retargeted" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    map_info = ElementTree.fromstring(members["xl/xmlMaps.xml"])
    xml_map = map_info.find(f"{{{namespace}}}Map")
    assert xml_map is not None
    xml_map.set("Append", "true")
    members["xl/xmlMaps.xml"] = ElementTree.tostring(
        map_info, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_xml_map_table_xpath_pair_changes_only_its_table_part(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "operations" / "xml_map_table_xpath_retargeted"
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
    ] == ["xl/tables/table1.xml"]


def test_validator_rejects_a_false_office_web_addin_auto_show_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "office_web_addin_auto_show_enabled"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_auto_show"] = False
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_office_web_addin_auto_show_property(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "office_web_addin_auto_show_enabled" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
    extension_member = "xl/webextensions/webextension1.xml"
    extension = ElementTree.fromstring(members[extension_member])
    property_element = next(
        item
        for item in extension.iter(f"{{{namespace}}}property")
        if item.get("name") == "Office.AutoShowTaskpaneWithDocument"
    )
    property_element.set("value", "false")
    members[extension_member] = ElementTree.tostring(
        extension, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_office_web_addin_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "office_web_addin_auto_show_enabled" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
    extension_member = "xl/webextensions/webextension1.xml"
    extension = ElementTree.fromstring(members[extension_member])
    reference = extension.find(f"{{{namespace}}}reference")
    assert reference is not None
    reference.set("store", "different-local-store.xml")
    members[extension_member] = ElementTree.tostring(
        extension, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_office_web_addin_auto_show_pair_changes_only_its_extension_part(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "office_web_addin_auto_show_enabled"
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
    ] == ["xl/webextensions/webextension1.xml"]


def test_validator_rejects_a_false_ole_object_auto_load_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "ole_object_auto_load_enabled"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_auto_load"] = False
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_ole_object_auto_load_flag(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "ole_object_auto_load_enabled" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    ole_object = worksheet.find(f".//{{{namespace}}}oleObject")
    assert ole_object is not None
    ole_object.set("autoLoad", "false")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_ole_object_change(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "ole_object_auto_load_enabled" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    worksheet_member = "xl/worksheets/sheet1.xml"
    worksheet = ElementTree.fromstring(members[worksheet_member])
    ole_object = worksheet.find(f".//{{{namespace}}}oleObject")
    assert ole_object is not None
    ole_object.set("progId", "WCAB.Other.Embedded.Object")
    members[worksheet_member] = ElementTree.tostring(
        worksheet, encoding="utf-8", xml_declaration=True
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_ole_object_auto_load_pair_changes_only_its_inputs_worksheet(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "ole_object_auto_load_enabled"
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


def test_formulafence_adapter_maps_the_exact_query_table_refresh_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _query_table_refresh_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "query_table_refresh_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF023", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "query_table_refresh_on_open"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["query_table_refresh_on_load_changed"]


def test_formulafence_adapter_rejects_an_inexact_query_table_refresh_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _query_table_refresh_details()
        details["after"][0]["connection_edit_disabled"] = False
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "query_table_refresh_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF023", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "query_table_refresh_on_open"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["query_table_refresh_on_load_changed"]


def test_formulafence_adapter_requires_the_query_table_refresh_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "query_table_refresh_controls_changed",
                    "location": None,
                    "details": _query_table_refresh_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "query_table_refresh_on_open"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["query_table_refresh_on_load_changed"]


def test_formulafence_adapter_maps_the_exact_cell_hyperlink_target_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _cell_hyperlink_target_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "cell_hyperlink_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF047", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "cell_hyperlink_target_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["cell_hyperlink_target_changed"]


def test_formulafence_adapter_rejects_an_inexact_cell_hyperlink_target_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _cell_hyperlink_target_details()
        details["after"]["hyperlink_with_tooltip_count"] = 1
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "cell_hyperlink_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF047", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "cell_hyperlink_target_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["cell_hyperlink_target_changed"]


def test_formulafence_adapter_requires_the_cell_hyperlink_target_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "cell_hyperlink_controls_changed",
                    "location": None,
                    "details": _cell_hyperlink_target_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "cell_hyperlink_target_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["cell_hyperlink_target_changed"]


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


def test_formulafence_adapter_maps_the_exact_external_workbook_link_source_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _external_workbook_link_source_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_link_packages_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF025", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_workbook_link_source_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["external_workbook_link_source_changed"]


def test_formulafence_adapter_rejects_an_inexact_external_workbook_link_source_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _external_workbook_link_source_details()
        details["after"]["external_workbook_sheet_count"] = 2
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_link_packages_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF025", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_workbook_link_source_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["external_workbook_link_source_changed"]


def test_formulafence_adapter_requires_the_external_workbook_link_source_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_link_packages_changed",
                    "location": None,
                    "details": _external_workbook_link_source_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_workbook_link_source_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["external_workbook_link_source_changed"]


def test_formulafence_adapter_maps_the_exact_external_defined_name_source_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        surface_details = _external_defined_name_source_details()
        definition_change_details = _external_defined_name_source_definition_change_details()
        definition_finding_details = _external_defined_name_source_definition_finding_details()
        return {
            "summary": {"change_count": 2},
            "changes": [
                {
                    "kind": "external_workbook_link_surfaces_changed",
                    "location": None,
                    "details": surface_details,
                },
                {
                    "kind": "defined_name_changed",
                    "location": None,
                    "details": definition_change_details,
                },
            ],
            "findings": [
                {"rule_id": "FF081", "details": surface_details},
                {"rule_id": "FF008", "details": definition_finding_details},
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_defined_name_source_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["external_defined_name_source_changed"]


def test_formulafence_adapter_requires_the_external_defined_name_definition_evidence(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        surface_details = _external_defined_name_source_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_workbook_link_surfaces_changed",
                    "location": None,
                    "details": surface_details,
                }
            ],
            "findings": [{"rule_id": "FF081", "details": surface_details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_defined_name_source_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["external_defined_name_source_changed"]


def test_formulafence_adapter_maps_the_exact_named_lambda_definition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "defined_name_changed",
                    "location": None,
                    "details": _named_lambda_definition_change_details(),
                }
            ],
            "findings": [
                {
                    "rule_id": "FF008",
                    "details": _named_lambda_definition_finding_details(),
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "named_lambda_definition_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["named_lambda_definition_changed"]


def test_formulafence_adapter_requires_named_lambda_finding(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "defined_name_changed",
                    "location": None,
                    "details": _named_lambda_definition_change_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "named_lambda_definition_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["named_lambda_definition_changed"]


def test_formulafence_adapter_maps_the_exact_power_pivot_data_model_relationship(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _power_pivot_data_model_relationship_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "power_pivot_data_model_changed",
                    "location": None,
                    "severity": "high",
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF033", "severity": "high", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "power_pivot_data_model_relationship_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["power_pivot_data_model_relationship_changed"]


def test_formulafence_adapter_rejects_an_inexact_power_pivot_data_model_relationship(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _power_pivot_data_model_relationship_details()
        details["workbook_data_model_declaration_changed"] = False
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "power_pivot_data_model_changed",
                    "location": None,
                    "severity": "high",
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF033", "severity": "high", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "power_pivot_data_model_relationship_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["power_pivot_data_model_relationship_changed"]


def test_formulafence_adapter_requires_power_pivot_data_model_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "power_pivot_data_model_changed",
                    "location": None,
                    "severity": "high",
                    "details": _power_pivot_data_model_relationship_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "power_pivot_data_model_relationship_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["power_pivot_data_model_relationship_changed"]


def test_formulafence_adapter_maps_the_exact_xlm_auto_open_binding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _xlm_auto_open_binding_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "xlm_automatic_macro_bindings_changed",
                    "location": None,
                    "severity": "high",
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF076", "severity": "high", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "xlm_auto_open_binding_retargeted"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["xlm_auto_open_binding_retargeted"]


def test_formulafence_adapter_rejects_an_inexact_xlm_auto_open_binding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _xlm_auto_open_binding_details()
        details["automatic_macro_binding_material_changed"] = False
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "xlm_automatic_macro_bindings_changed",
                    "location": None,
                    "severity": "high",
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF076", "severity": "high", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "xlm_auto_open_binding_retargeted"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["xlm_auto_open_binding_retargeted"]


def test_formulafence_adapter_requires_xlm_auto_open_binding_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "xlm_automatic_macro_bindings_changed",
                    "location": None,
                    "severity": "high",
                    "details": _xlm_auto_open_binding_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "xlm_auto_open_binding_retargeted"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["xlm_auto_open_binding_retargeted"]


def test_formulafence_adapter_maps_the_exact_table_calculated_column_formula(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _table_calculated_column_formula_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "table_definition_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF013", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "table_calculated_column_formula_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["table_calculated_column_formula_changed"]


def test_formulafence_adapter_rejects_an_inexact_table_calculated_column_formula(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _table_calculated_column_formula_details()
        details["calculated_column_formula_material_changed"] = False
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "table_definition_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF013", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "table_calculated_column_formula_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["table_calculated_column_formula_changed"]


def test_formulafence_adapter_requires_table_calculated_column_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "table_definition_changed",
                    "location": None,
                    "details": _table_calculated_column_formula_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "table_calculated_column_formula_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["table_calculated_column_formula_changed"]


def test_formulafence_adapter_maps_the_exact_sheet_protection_sort_permission(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _sheet_protection_sort_permission_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "sheet_protection_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF022", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "sheet_protection_sort_permission_enabled"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["sheet_protection_sort_permission_enabled"]


def test_formulafence_adapter_rejects_an_inexact_sheet_protection_sort_permission(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _sheet_protection_sort_permission_details()
        details["after"]["locked_actions"].append("sort")
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "sheet_protection_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF022", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "sheet_protection_sort_permission_enabled"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["sheet_protection_sort_permission_enabled"]


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


def test_formulafence_adapter_maps_the_exact_named_sheet_view_filter_criterion_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _named_sheet_view_filter_criterion_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "named_sheet_views_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF038", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "named_sheet_view_filter_criterion_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["named_sheet_view_filter_criterion_changed"]


def test_formulafence_adapter_rejects_an_inexact_named_sheet_view_filter_criterion_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _named_sheet_view_filter_criterion_details()
        details["after"]["sort_rule_count"] = 1
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "named_sheet_views_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF038", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "named_sheet_view_filter_criterion_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["named_sheet_view_filter_criterion_changed"]


def test_formulafence_adapter_requires_the_named_sheet_view_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "named_sheet_views_changed",
                    "location": None,
                    "details": _named_sheet_view_filter_criterion_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "named_sheet_view_filter_criterion_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["named_sheet_view_filter_criterion_changed"]


def test_formulafence_adapter_maps_the_exact_xml_map_table_xpath_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _xml_map_table_xpath_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "xml_mapping_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF049", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "xml_map_table_xpath_retargeted"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["xml_map_table_column_xpath_retargeted"]


def test_formulafence_adapter_rejects_an_inexact_xml_map_table_xpath_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _xml_map_table_xpath_details()
        details["after"]["single_cell_xml_binding_count"] = 2
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "xml_mapping_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF049", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "xml_map_table_xpath_retargeted"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["xml_map_table_column_xpath_retargeted"]


def test_formulafence_adapter_requires_the_xml_map_table_xpath_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "xml_mapping_controls_changed",
                    "location": None,
                    "details": _xml_map_table_xpath_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "xml_map_table_xpath_retargeted"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["xml_map_table_column_xpath_retargeted"]


def test_formulafence_adapter_maps_the_exact_office_web_addin_auto_show_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _office_web_addin_auto_show_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "office_web_addins_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF028", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "office_web_addin_auto_show_enabled"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["office_web_addin_auto_show_enabled"]


def test_formulafence_adapter_rejects_an_inexact_office_web_addin_auto_show_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _office_web_addin_auto_show_details()
        details["after"]["auto_show_taskpane_count"] = 2
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "office_web_addins_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF028", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "office_web_addin_auto_show_enabled"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["office_web_addin_auto_show_enabled"]


def test_formulafence_adapter_requires_the_office_web_addin_auto_show_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "office_web_addins_changed",
                    "location": None,
                    "details": _office_web_addin_auto_show_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "office_web_addin_auto_show_enabled"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["office_web_addin_auto_show_enabled"]


def test_formulafence_adapter_maps_the_exact_ole_object_auto_load_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _ole_object_auto_load_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "worksheet_embedded_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF029", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "ole_object_auto_load_enabled"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["ole_object_auto_load_enabled"]


def test_formulafence_adapter_rejects_an_inexact_ole_object_auto_load_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _ole_object_auto_load_details()
        details["after"]["external_relationship_count"] = 1
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "worksheet_embedded_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF029", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "ole_object_auto_load_enabled"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["ole_object_auto_load_enabled"]


def test_formulafence_adapter_requires_the_ole_object_auto_load_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "worksheet_embedded_controls_changed",
                    "location": None,
                    "details": _ole_object_auto_load_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "ole_object_auto_load_enabled"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["ole_object_auto_load_enabled"]


def test_formulafence_adapter_maps_the_exact_data_validation_list_source_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _data_validation_list_source_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "data_validation_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF020", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "data_validation_list_source_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["data_validation_list_source_changed"]


def test_formulafence_adapter_rejects_an_inexact_data_validation_list_source_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _data_validation_list_source_details()
        details["after"][0]["formula1"] = "Lists!$A$2:$A$4"
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "data_validation_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF020", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "data_validation_list_source_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["data_validation_list_source_changed"]


def test_formulafence_adapter_requires_the_data_validation_list_source_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "data_validation_changed",
                    "location": None,
                    "details": _data_validation_list_source_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "data_validation_list_source_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["data_validation_list_source_changed"]


def test_formulafence_adapter_maps_the_exact_conditional_formatting_threshold_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _conditional_formatting_threshold_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "conditional_formatting_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF021", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "conditional_formatting_threshold_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["conditional_formatting_threshold_changed"]


def test_formulafence_adapter_rejects_an_inexact_conditional_formatting_threshold_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _conditional_formatting_threshold_details()
        details["after"]["rules"][0]["formulas"] = ["100"]
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "conditional_formatting_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF021", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "conditional_formatting_threshold_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["conditional_formatting_threshold_changed"]


def test_formulafence_adapter_requires_the_conditional_formatting_threshold_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "conditional_formatting_changed",
                    "location": None,
                    "details": _conditional_formatting_threshold_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "conditional_formatting_threshold_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["conditional_formatting_threshold_changed"]


def test_formulafence_adapter_maps_the_exact_number_format_visibility_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _number_format_visibility_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "number_format_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF039", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "number_format_value_hidden"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["cell_number_format_changed"]


def test_formulafence_adapter_rejects_an_inexact_number_format_visibility_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _number_format_visibility_details()
        details["after"]["custom_format_assignment_count"] = 2
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "number_format_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF039", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "number_format_value_hidden"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["cell_number_format_changed"]


def test_formulafence_adapter_requires_the_number_format_visibility_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "number_format_controls_changed",
                    "location": None,
                    "details": _number_format_visibility_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "number_format_value_hidden"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["cell_number_format_changed"]


def test_formulafence_adapter_maps_the_exact_ignored_error_suppression(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _ignored_error_formula_range_suppression_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "ignored_error_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF037", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "ignored_error_formula_range_suppressed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["ignored_error_rule_added"]


def test_formulafence_adapter_rejects_an_inexact_ignored_error_suppression(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _ignored_error_formula_range_suppression_details()
        details["after"]["formula_range_omission_count"] = 2
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "ignored_error_controls_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF037", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "ignored_error_formula_range_suppressed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["ignored_error_rule_added"]


def test_formulafence_adapter_requires_the_ignored_error_suppression_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "ignored_error_controls_changed",
                    "location": None,
                    "details": _ignored_error_formula_range_suppression_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "operations" / "ignored_error_formula_range_suppressed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["ignored_error_rule_added"]


def test_formulafence_adapter_maps_the_exact_workbook_structure_lock_removed(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _workbook_structure_lock_removed_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "workbook_protection_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF022", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "workbook_structure_lock_removed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["workbook_structure_lock_removed"]


def test_formulafence_adapter_rejects_an_inexact_workbook_structure_lock_removed(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _workbook_structure_lock_removed_details()
        details["after"]["lock_windows"] = True
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "workbook_protection_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF022", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "workbook_structure_lock_removed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["workbook_structure_lock_removed"]


def test_formulafence_adapter_requires_the_workbook_structure_lock_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "workbook_protection_changed",
                    "location": None,
                    "details": _workbook_structure_lock_removed_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "workbook_structure_lock_removed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["workbook_structure_lock_removed"]


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


def test_formulafence_adapter_maps_the_exact_what_if_data_table_input_reference_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _what_if_data_table_input_reference_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "what_if_data_tables_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF034", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "what_if_data_table_input_reference_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["what_if_data_table_input_reference_changed"]


def test_formulafence_adapter_rejects_an_inexact_what_if_data_table_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _what_if_data_table_input_reference_details()
        details["before"]["recalculation_requested_count"] = 0
        details["after"]["recalculation_requested_count"] = 0
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "what_if_data_tables_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF034", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "what_if_data_table_input_reference_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["what_if_data_table_input_reference_changed"]


def test_formulafence_adapter_requires_the_what_if_data_table_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "what_if_data_tables_changed",
                    "location": None,
                    "details": _what_if_data_table_input_reference_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "what_if_data_table_input_reference_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["what_if_data_table_input_reference_changed"]


def test_formulafence_adapter_maps_the_exact_scenario_manager_stored_input_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _scenario_manager_stored_input_details()
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "scenario_manager_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF035", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "scenario_manager_stored_input_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["scenario_manager_stored_input_value_changed"]


def test_formulafence_adapter_rejects_an_inexact_scenario_manager_stored_input_change(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        details = _scenario_manager_stored_input_details()
        details["selection_material_changed"] = True
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "scenario_manager_changed",
                    "location": None,
                    "details": details,
                }
            ],
            "findings": [{"rule_id": "FF035", "details": details}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "scenario_manager_stored_input_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["scenario_manager_stored_input_value_changed"]


def test_formulafence_adapter_requires_the_scenario_manager_stored_input_finding(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "scenario_manager_changed",
                    "location": None,
                    "details": _scenario_manager_stored_input_details(),
                }
            ],
            "findings": [],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "scenario_manager_stored_input_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["scenario_manager_stored_input_value_changed"]


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
