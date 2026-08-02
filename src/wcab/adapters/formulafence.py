"""Local-only FormulaFence adapter.

FormulaFence is not a dependency of WCAB. This adapter shells out to an
already installed FormulaFence executable and normalizes its JSON evidence so
the benchmark can demonstrate one concrete implementation without making its
report schema normative for other tools.
"""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

from ..build import CASE_IDS, FIXTURE_SCHEMA_VERSION
from ..score import OBSERVATION_SCHEMA_VERSION


class FormulaFenceAdapterError(RuntimeError):
    """FormulaFence could not produce a usable local JSON report."""


_FACT_TO_CHANGE: dict[str, tuple[str, str | None]] = {
    "formula_to_value": ("formula_to_value", "location"),
    "formula_changed": ("formula_changed", "location"),
    "value_changed": ("value_changed", "location"),
    "external_formula_added": ("external_workbook_link_surfaces_changed", None),
    "defined_name_changed": ("defined_name_changed", None),
    "data_validation_count_changed": ("data_validation_changed", None),
    "data_validation_list_source_changed": ("data_validation_changed", None),
    "conditional_formatting_count_changed": ("conditional_formatting_changed", None),
    "conditional_formatting_threshold_changed": ("conditional_formatting_changed", None),
    "cell_number_format_changed": ("number_format_controls_changed", None),
    "auto_filter_criteria_changed": ("filter_visibility_controls_changed", None),
    "sheet_visibility_changed": ("sheet_visibility_changed", None),
    "formula_cell_unlocked": ("cell_protection_assignments_changed", None),
    "manual_calculation_incomplete": ("calculation_settings_changed", None),
    "iterative_calculation_enabled": ("calculation_settings_changed", None),
    "precision_as_displayed_enabled": ("calculation_settings_changed", None),
    "workbook_date_system_changed": ("workbook_date_system_changed", None),
    "formula_cached_result_changed": ("formula_cached_result_changed", None),
    "array_formula_mode_changed": ("array_formula_mode_changed", "location"),
    "chart_series_value_reference_changed": ("chart_definitions_changed", None),
    "pivot_data_field_aggregation_changed": ("pivot_table_definitions_changed", None),
    "pivot_slicer_selection_changed": ("slicer_timeline_cache_definitions_changed", None),
    "power_query_m_filter_changed": ("power_query_changed", None),
    "scenario_manager_stored_input_value_changed": ("scenario_manager_changed", None),
    "what_if_data_table_input_reference_changed": ("what_if_data_tables_changed", None),
    "external_data_connection_refresh_on_load_changed": (
        "external_data_connections_changed",
        None,
    ),
    "pivot_cache_refresh_on_load_changed": ("pivot_cache_refresh_controls_changed", None),
    "external_workbook_link_update_policy_changed": (
        "external_data_refresh_settings_changed",
        None,
    ),
    "static_cycle_introduced": ("formula_changed", None),
    "three_d_scope_changed": ("three_d_reference_scope_changed", "formula_location"),
    "structured_table_scope_changed": ("table_definition_changed", None),
    "dynamic_formula_reference_added": ("dynamic_formula_reference_added", "location"),
}

_COVERAGE_EXPECTATION_TO_RULE: dict[str, tuple[str, str | None]] = {
    "dynamic_reference_static_coverage": ("FF012", "location"),
}

_COVERAGE_EXPECTATION_TO_PROFILE_FEATURE = {
    "dynamic_reference_driver_changed": "dynamic_reference_cells",
}

_LINT_EXPECTATIONS = {
    "operations.copied_formula_interruption": "FF082",
    "operations.sumifs_range_shape": "FF093",
    "governance.formula_cell_unlocked": "FF085",
    "governance.manual_calculation_incomplete": "FF086",
    "governance.static_cycle_introduced": "FF090",
}

_PORTFOLIO_FACT_EVIDENCE = {
    "portfolio_value_changed": {"native_change_kind": "value_changed"},
    "portfolio_external_reference": {"native_rule_id": "FF079"},
}


def _resolve_executable(executable: str) -> str:
    resolved = shutil.which(executable)
    if resolved is None:
        candidate = Path(executable)
        if candidate.is_file():
            return str(candidate)
        raise FormulaFenceAdapterError(
            f"FormulaFence executable {executable!r} was not found; install it or pass --executable."
        )
    return resolved


def _run_json(arguments: list[str], *, executable: str, timeout: int = 120) -> dict[str, Any]:
    command = [_resolve_executable(executable), *arguments]
    with TemporaryDirectory(prefix="wcab-formulafence-") as temporary:
        report_path = Path(temporary) / "report.json"
        try:
            completed = subprocess.run(
                [*command, "--format", "json", "--output", str(report_path)],
                text=True,
                capture_output=True,
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired as error:
            raise FormulaFenceAdapterError(
                f"FormulaFence exceeded the {timeout}-second adapter limit"
            ) from error
        if completed.returncode != 0:
            raise FormulaFenceAdapterError(
                f"FormulaFence exited {completed.returncode}: {completed.stderr.strip()}"
            )
        if not report_path.is_file():
            raise FormulaFenceAdapterError(
                f"FormulaFence did not write JSON (exit {completed.returncode}): {completed.stderr.strip()}"
            )
        try:
            report = json.loads(report_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise FormulaFenceAdapterError(f"FormulaFence emitted invalid JSON: {error}") from error
    if not isinstance(report, dict):
        raise FormulaFenceAdapterError("FormulaFence report root must be a JSON object")
    return report


def diff(
    baseline: str | Path, candidate: str | Path, *, executable: str = "formulafence"
) -> dict[str, Any]:
    """Run a local FormulaFence semantic diff and return its JSON report."""

    return _run_json(["diff", str(baseline), str(candidate)], executable=executable)


def profile(workbook: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Run a local FormulaFence profile and return its JSON report."""

    return _run_json(["profile", str(workbook)], executable=executable)


def lint(workbook: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Run local FormulaFence lint and return its JSON report."""

    return _run_json(["lint", str(workbook)], executable=executable)


def portfolio(
    baseline: str | Path, candidate: str | Path, *, executable: str = "formulafence"
) -> dict[str, Any]:
    """Run FormulaFence's local directory-portfolio comparison."""

    return _run_json(["portfolio", str(baseline), str(candidate)], executable=executable)


def _load_truth(case_dir: Path) -> dict[str, Any]:
    result = json.loads((case_dir / "truth.json").read_text(encoding="utf-8"))
    if not isinstance(result, dict):
        raise FormulaFenceAdapterError(f"{case_dir}: truth root must be an object")
    return result


def _fact_location(fact: dict[str, Any]) -> str | None:
    if fact["kind"] == "three_d_scope_changed":
        return f"{fact['formula_sheet']}!{fact['formula_cell']}"
    sheet = fact.get("sheet")
    cell = fact.get("cell")
    if isinstance(sheet, str) and isinstance(cell, str):
        return f"{sheet}!{cell}"
    return None


def _nested_location(expectation: dict[str, Any], field: str) -> str | None:
    value = expectation.get(field)
    if not isinstance(value, dict):
        return None
    sheet = value.get("sheet")
    cell = value.get("cell")
    if isinstance(sheet, str) and isinstance(cell, str):
        return f"{sheet}!{cell}"
    return None


def _coverage_evidence(expectation: dict[str, Any]) -> dict[str, str] | None:
    kind = str(expectation.get("kind"))
    rule_mapping = _COVERAGE_EXPECTATION_TO_RULE.get(kind)
    if rule_mapping is not None:
        return {"native_rule_id": rule_mapping[0]}
    profile_feature = _COVERAGE_EXPECTATION_TO_PROFILE_FEATURE.get(kind)
    if profile_feature is not None:
        return {
            "native_change_kind": "value_changed",
            "native_profile_feature": profile_feature,
        }
    return None


def _external_data_connection_refresh_on_load_observed(
    change: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Match FormulaFence's private control summary to the exact WCAB fact."""

    if change.get("kind") != "external_data_connections_changed":
        return False
    connection_id = fact.get("connection_id")
    before_refresh_on_load = fact.get("baseline_refresh_on_load")
    after_refresh_on_load = fact.get("candidate_refresh_on_load")
    if (
        not isinstance(connection_id, int)
        or not isinstance(before_refresh_on_load, bool)
        or not isinstance(after_refresh_on_load, bool)
    ):
        return False
    details = change.get("details")
    if not isinstance(details, dict):
        return False
    before_connections = details.get("before")
    after_connections = details.get("after")
    if not isinstance(before_connections, list) or not isinstance(after_connections, list):
        return False

    def refresh_value(connections: list[Any]) -> bool | None:
        matches = [
            connection
            for connection in connections
            if isinstance(connection, dict) and connection.get("id") == connection_id
        ]
        if len(matches) != 1 or not isinstance(matches[0].get("refresh_on_load"), bool):
            return None
        return matches[0]["refresh_on_load"]

    return (
        refresh_value(before_connections) is before_refresh_on_load
        and refresh_value(after_connections) is after_refresh_on_load
    )


def _external_workbook_link_update_policy_observed(
    change: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Require FormulaFence's exact global external-link policy transition."""

    if change.get("kind") != "external_data_refresh_settings_changed":
        return False
    before_update_links = fact.get("baseline_update_links")
    after_update_links = fact.get("candidate_update_links")
    if before_update_links != "never" or after_update_links != "always":
        return False
    details = change.get("details")
    if not isinstance(details, dict):
        return False
    expected_before = {
        "update_links": before_update_links,
        "allow_refresh_query": False,
        "refresh_all_connections": False,
        "save_external_link_values": True,
    }
    expected_after = {**expected_before, "update_links": after_update_links}
    return details.get("before") == expected_before and details.get("after") == expected_after


def _iterative_calculation_enabled_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require FormulaFence's exact iteration switch and unchanged bounds."""

    if change.get("kind") != "calculation_settings_changed":
        return False
    count = fact.get("iteration_count")
    delta = fact.get("iteration_delta")
    if (
        fact.get("baseline_iterate") is not False
        or fact.get("candidate_iterate") is not True
        or not isinstance(count, int)
        or isinstance(count, bool)
        or not isinstance(delta, (int, float))
        or isinstance(delta, bool)
    ):
        return False
    details = change.get("details")
    if not isinstance(details, dict):
        return False
    expected_before = {
        "forceFullCalc": False,
        "fullCalcOnLoad": False,
        "iterate": False,
        "iterateCount": count,
        "iterateDelta": delta,
    }
    expected_after = {**expected_before, "iterate": True}
    return details.get("before") == expected_before and details.get("after") == expected_after


def _precision_as_displayed_enabled_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require FormulaFence's exact full-precision switch and stable controls."""

    if change.get("kind") != "calculation_settings_changed":
        return False
    if (
        fact.get("baseline_full_precision") is not True
        or fact.get("candidate_full_precision") is not False
    ):
        return False
    details = change.get("details")
    if not isinstance(details, dict):
        return False
    expected_before = {
        "calcCompleted": True,
        "calcMode": "auto",
        "calcOnSave": True,
        "forceFullCalc": False,
        "fullCalcOnLoad": False,
        "fullPrecision": True,
    }
    expected_after = {**expected_before, "fullPrecision": False}
    return details.get("before") == expected_before and details.get("after") == expected_after


def _workbook_date_system_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Require FormulaFence's exact normalized date-control transition."""

    if (
        fact.get("baseline_date_1904") is not False
        or fact.get("candidate_date_1904") is not True
        or fact.get("date_compatibility") is not True
    ):
        return False
    if not isinstance(details, dict):
        return False
    expected_before = {
        "date_1904": False,
        "date_compatibility": True,
        "date_compatibility_declared": True,
        "unrecognized_control_count": 0,
    }
    expected_after = {**expected_before, "date_1904": True}
    return details.get("before") == expected_before and details.get("after") == expected_after


def _workbook_date_system_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's dedicated global date-system change record."""

    return change.get("kind") == "workbook_date_system_changed" and (
        _workbook_date_system_details_observed(change.get("details"), fact)
    )


def _workbook_date_system_finding_observed(finding: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require FormulaFence's matching high-severity FF117 finding."""

    return finding.get("rule_id") == "FF117" and _workbook_date_system_details_observed(
        finding.get("details"), fact
    )


def _auto_filter_criteria_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Require the exact redacted FormulaFence AutoFilter evidence shape."""

    if (
        fact.get("sheet") != "Report"
        or fact.get("filter_ref") != "A1:B5"
        or fact.get("filter_column_id") != 0
        or fact.get("baseline_filter_value") != "North"
        or fact.get("candidate_filter_value") != "South"
    ):
        return False
    expected_surface = {
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
    return details == {
        "before": expected_surface,
        "after": expected_surface,
        "filter_visibility_definition_material_changed": True,
    }


def _auto_filter_criteria_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's active-filter control record."""

    return change.get("kind") == "filter_visibility_controls_changed" and (
        _auto_filter_criteria_details_observed(change.get("details"), fact)
    )


def _auto_filter_criteria_finding_observed(finding: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require FormulaFence's matching high-severity FF036 finding."""

    return finding.get("rule_id") == "FF036" and _auto_filter_criteria_details_observed(
        finding.get("details"), fact
    )


def _pivot_cache_refresh_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Require FormulaFence's exact redacted PivotCache control transition."""

    if (
        fact.get("cache_id") != 1
        or fact.get("source_type") != "worksheet"
        or fact.get("source_sheet") != "Source"
        or fact.get("source_ref") != "A1:B5"
        or fact.get("pivot_sheet") != "Report"
        or fact.get("pivot_ref") != "A1:B2"
        or fact.get("pivot_output_cell") != "B2"
        or fact.get("dashboard_sheet") != "Dashboard"
        or fact.get("dashboard_cell") != "B4"
        or fact.get("dashboard_formula") != "=Report!$B$2"
        or fact.get("baseline_refresh_on_load") is not False
        or fact.get("candidate_refresh_on_load") is not True
    ):
        return False
    if not isinstance(details, dict):
        return False
    expected_before = [
        {
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
    ]
    expected_after = [{**expected_before[0], "refresh_on_load": True}]
    return details == {"before": expected_before, "after": expected_after}


def _pivot_cache_refresh_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's dedicated PivotCache control record."""

    return change.get("kind") == "pivot_cache_refresh_controls_changed" and (
        _pivot_cache_refresh_details_observed(change.get("details"), fact)
    )


def _pivot_cache_refresh_finding_observed(finding: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require FormulaFence's matching high-severity FF023 finding."""

    return finding.get("rule_id") == "FF023" and _pivot_cache_refresh_details_observed(
        finding.get("details"), fact
    )


def _chart_series_value_reference_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Match FormulaFence's redacted single-chart definition transition.

    FormulaFence intentionally does not expose chart formulas or visual output.
    WCAB therefore requires its exact one-chart structural profile and FF030,
    then independently establishes the title/category/value references through
    raw package validation.
    """

    required_fact_fields = (
        "chart_sheet",
        "chart_anchor",
        "source_sheet",
        "series_title_ref",
        "category_ref",
        "baseline_value_ref",
        "candidate_value_ref",
    )
    if (
        any(not isinstance(fact.get(field), str) for field in required_fact_fields)
        or fact.get("baseline_value_ref") == fact.get("candidate_value_ref")
        or not isinstance(details, dict)
    ):
        return False
    expected_profile = {
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
    return (
        details.get("before") == expected_profile
        and details.get("after") == expected_profile
        and details.get("chart_definition_material_changed") is True
        and set(details) == {"before", "after", "chart_definition_material_changed"}
    )


def _chart_series_value_reference_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's dedicated chart-definition change record."""

    return change.get("kind") == "chart_definitions_changed" and (
        _chart_series_value_reference_details_observed(change.get("details"), fact)
    )


def _chart_series_value_reference_finding_observed(
    finding: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Require FormulaFence's matching high-severity chart finding."""

    return finding.get("rule_id") == "FF030" and _chart_series_value_reference_details_observed(
        finding.get("details"), fact
    )


def _pivot_data_field_aggregation_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Match FormulaFence's redacted one-PivotTable aggregation evidence.

    FormulaFence intentionally reports only PivotTable structure, not source
    labels, cached values, or the selected aggregate function. WCAB therefore
    requires exactly its one-PivotTable layout material change and FF031, then
    independently validates the stored ``sum -> average`` declaration.
    """

    if (
        fact.get("cache_id") != 1
        or fact.get("source_type") != "worksheet"
        or fact.get("source_sheet") != "Source"
        or fact.get("source_ref") != "A1:B5"
        or fact.get("pivot_sheet") != "Report"
        or fact.get("pivot_ref") != "A1:B2"
        or fact.get("pivot_output_cell") != "B2"
        or fact.get("dashboard_sheet") != "Dashboard"
        or fact.get("dashboard_cell") != "B4"
        or fact.get("dashboard_formula") != "=Report!$B$2"
        or fact.get("data_field_source_index") != 1
        or fact.get("baseline_subtotal") != "sum"
        or fact.get("candidate_subtotal") != "average"
        or not isinstance(details, dict)
    ):
        return False
    expected_profile = {
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
    return (
        details.get("before") == expected_profile
        and details.get("after") == expected_profile
        and details.get("pivot_table_layout_material_changed") is True
        and set(details) == {"before", "after", "pivot_table_layout_material_changed"}
    )


def _pivot_data_field_aggregation_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's PivotTable-definition change record."""

    return change.get("kind") == "pivot_table_definitions_changed" and (
        _pivot_data_field_aggregation_details_observed(change.get("details"), fact)
    )


def _pivot_data_field_aggregation_finding_observed(
    finding: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Require FormulaFence's matching high-severity PivotTable finding."""

    return finding.get("rule_id") == "FF031" and _pivot_data_field_aggregation_details_observed(
        finding.get("details"), fact
    )


def _pivot_slicer_selection_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Match FormulaFence's redacted one-Slicer selection evidence.

    FormulaFence keeps Slicer names, selected values, and report output private.
    WCAB therefore requires its exact one-Slicer structural profile and FF032,
    then independently validates the stored cache-item selection and local
    PivotTable relationship graph.
    """

    if (
        fact.get("cache_id") != 1
        or fact.get("source_type") != "worksheet"
        or fact.get("source_sheet") != "Source"
        or fact.get("source_ref") != "A1:B5"
        or fact.get("pivot_sheet") != "Report"
        or fact.get("pivot_ref") != "A1:B2"
        or fact.get("pivot_output_cell") != "B2"
        or fact.get("dashboard_sheet") != "Dashboard"
        or fact.get("dashboard_cell") != "B4"
        or fact.get("dashboard_formula") != "=Report!$B$2"
        or fact.get("slicer_name") != "WCAB Region slicer"
        or fact.get("slicer_source_name") != "Region"
        or fact.get("slicer_pivot_table_name") != "WCAB Pivot Report"
        or fact.get("slicer_pivot_tab_id") != 2
        or fact.get("item_count") != 2
        or fact.get("baseline_selected_item_index") != 0
        or fact.get("candidate_selected_item_index") != 1
        or fact.get("baseline_selected_value") != "North"
        or fact.get("candidate_selected_value") != "South"
        or not isinstance(details, dict)
    ):
        return False
    expected_profile = {
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
    return (
        details.get("before") == expected_profile
        and details.get("after") == expected_profile
        and details.get("slicer_filter_state_or_definition_material_changed") is True
        and set(details)
        == {"before", "after", "slicer_filter_state_or_definition_material_changed"}
    )


def _pivot_slicer_selection_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's Slicer/Timeline cache definition change record."""

    return change.get("kind") == "slicer_timeline_cache_definitions_changed" and (
        _pivot_slicer_selection_details_observed(change.get("details"), fact)
    )


def _pivot_slicer_selection_finding_observed(finding: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require FormulaFence's matching high-severity Slicer finding."""

    return finding.get("rule_id") == "FF032" and _pivot_slicer_selection_details_observed(
        finding.get("details"), fact
    )


def _power_query_m_filter_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Match FormulaFence's redacted one-query M-definition evidence.

    FormulaFence deliberately reports the shape of a Power Query surface, not
    M source text, local-table values, or query results. WCAB therefore
    requires this exact one-query profile and FF024, while its raw validator
    establishes the stored M transition and connection-only controls.
    """

    if (
        fact.get("data_mashup_part") != "customXml/item1.xml"
        or fact.get("source_sheet") != "Source"
        or fact.get("source_table") != "SourceData"
        or fact.get("source_ref") != "A1:B5"
        or fact.get("query_section") != "Section1"
        or fact.get("query_name") != "RegionQuery"
        or fact.get("filter_column") != "Region"
        or fact.get("baseline_filter_value") != "North"
        or fact.get("candidate_filter_value") != "South"
        or fact.get("fill_enabled") is not False
        or fact.get("firewall_enabled") is not True
        or fact.get("future_packages_allowed") is not False
        or not isinstance(details, dict)
    ):
        return False
    expected_profile = {
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
    return (
        details.get("before") == expected_profile
        and details.get("after") == expected_profile
        and details.get("formula_material_changed") is True
        and set(details) == {"before", "after", "formula_material_changed"}
    )


def _power_query_m_filter_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's Power Query definition change record."""

    return change.get("kind") == "power_query_changed" and (
        _power_query_m_filter_details_observed(change.get("details"), fact)
    )


def _power_query_m_filter_finding_observed(finding: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require FormulaFence's matching high-severity Power Query finding."""

    return finding.get("rule_id") == "FF024" and _power_query_m_filter_details_observed(
        finding.get("details"), fact
    )


def _data_validation_list_source_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Require FormulaFence's exact list-source transition and rule shape.

    FormulaFence exposes the stored validation source and its entry-control
    metadata, but it appropriately does not evaluate a list formula or decide
    whether a future input would be accepted. WCAB's raw validator establishes
    the compact package-only change; this adapter requires its exact native
    evidence and high-severity FF020 finding.
    """

    if (
        fact.get("validation_sheet") != "Inputs"
        or fact.get("target_range") != "B2"
        or fact.get("validation_type") != "list"
        or fact.get("baseline_source_formula") != "=Lists!$A$2:$A$4"
        or fact.get("candidate_source_formula") != "=Lists!$B$2:$B$4"
        or fact.get("allow_blank") is not False
        or fact.get("dropdown_hidden") is not False
        or fact.get("show_error_message") is not True
        or fact.get("error_style") != "stop"
        or fact.get("error_title") != "Invalid status"
        or fact.get("error") != "Choose an approved status."
        or fact.get("show_input_message") is not False
        or fact.get("prompt_title") != "Approved status"
        or fact.get("prompt") != "Choose a documented status."
        or fact.get("source_sheet") != "Lists"
        or fact.get("baseline_source_range") != "A2:A4"
        or fact.get("candidate_source_range") != "B2:B4"
        or fact.get("baseline_source_values") != ["Draft", "Review", "Approved"]
        or fact.get("candidate_source_values") != ["Draft", "Suspended", "Rejected"]
        or fact.get("input_cell") != "B2"
        or fact.get("input_value") != "Draft"
        or fact.get("model_sheet") != "Model"
        or fact.get("model_cell") != "B2"
        or fact.get("model_formula") != "=Inputs!$B$2"
        or fact.get("dashboard_sheet") != "Dashboard"
        or fact.get("dashboard_cell") != "B4"
        or fact.get("dashboard_formula") != "=Model!$B$2"
        or not isinstance(details, dict)
    ):
        return False
    expected_rule = {
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
    return details == {
        "sheet": "Inputs",
        "before": [{**expected_rule, "formula1": "Lists!$A$2:$A$4"}],
        "after": [{**expected_rule, "formula1": "Lists!$B$2:$B$4"}],
    }


def _data_validation_list_source_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's stored data-validation control-change record."""

    return change.get("kind") == "data_validation_changed" and (
        _data_validation_list_source_details_observed(change.get("details"), fact)
    )


def _data_validation_list_source_finding_observed(
    finding: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Require FormulaFence's matching high-severity data-validation finding."""

    return finding.get("rule_id") == "FF020" and _data_validation_list_source_details_observed(
        finding.get("details"), fact
    )


def _conditional_formatting_threshold_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Require FormulaFence's exact one-rule threshold transition and FF021.

    FormulaFence exposes the stored conditional-formatting rule rather than a
    rendered workbook. WCAB independently validates the raw cellIs formula,
    stable differential fill, values, and package boundary; this adapter
    requires FormulaFence to show that same one-rule formula transition.
    """

    if (
        fact.get("sheet") != "Operations"
        or fact.get("target_range") != "B2:B4"
        or fact.get("priority") != 1
        or fact.get("rule_type") != "cellIs"
        or fact.get("operator") != "greaterThan"
        or fact.get("baseline_formula") != "100"
        or fact.get("candidate_formula") != "50"
        or fact.get("metric_values") != [10, 75, 120]
        or fact.get("fill_rgb") != "FFFFC7CE"
        or not isinstance(details, dict)
        or set(details) != {"sheet", "before", "after"}
        or details.get("sheet") != "Operations"
    ):
        return False
    before = details.get("before")
    after = details.get("after")
    if (
        not isinstance(before, dict)
        or not isinstance(after, dict)
        or set(before) != {"rules", "extensions"}
        or set(after) != {"rules", "extensions"}
        or before.get("extensions") != []
        or after.get("extensions") != []
        or not isinstance(before.get("rules"), list)
        or not isinstance(after.get("rules"), list)
        or len(before["rules"]) != 1
        or len(after["rules"]) != 1
    ):
        return False
    before_rule = before["rules"][0]
    after_rule = after["rules"][0]
    if not isinstance(before_rule, dict) or not isinstance(after_rule, dict):
        return False
    expected_controls = {
        "sheet": "Operations",
        "ranges": ["Operations!B2:B4"],
        "priority": 1,
        "type": "cellIs",
        "operator": "greaterThan",
        "stop_if_true": False,
        "above_average": True,
        "percent": False,
        "bottom": False,
        "rank": None,
        "std_dev": None,
        "equal_average": False,
        "text": None,
        "time_period": None,
    }
    if any(before_rule.get(key) != value for key, value in expected_controls.items()) or any(
        after_rule.get(key) != value for key, value in expected_controls.items()
    ):
        return False
    before_without_formula = dict(before_rule)
    after_without_formula = dict(after_rule)
    before_formula = before_without_formula.pop("formulas", None)
    after_formula = after_without_formula.pop("formulas", None)
    return (
        before_formula == ["100"]
        and after_formula == ["50"]
        and before_without_formula == after_without_formula
    )


def _conditional_formatting_threshold_observed(
    change: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Match FormulaFence's stored conditional-formatting change record."""

    return change.get("kind") == "conditional_formatting_changed" and (
        _conditional_formatting_threshold_details_observed(change.get("details"), fact)
    )


def _conditional_formatting_threshold_finding_observed(
    finding: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Require FormulaFence's matching high-severity conditional-format finding."""

    return finding.get("rule_id") == "FF021" and _conditional_formatting_threshold_details_observed(
        finding.get("details"), fact
    )


def _number_format_visibility_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Require FormulaFence's exact custom-number-format structural evidence.

    FormulaFence intentionally redacts number-format codes and cell targets from
    its shared report. WCAB independently validates the one custom code,
    referenced style, target, and package boundary; this adapter requires the
    matching one-custom-assignment structural transition and FF039 finding.
    """

    if (
        fact.get("sheet") != "Operations"
        or fact.get("cell") != "B2"
        or fact.get("value") != 0.125
        or fact.get("custom_number_format_id") != 164
        or fact.get("baseline_format") != "0.0%;[Red](0.0%);-"
        or fact.get("candidate_format") != ";;;"
        or fact.get("formula_cell") != "B3"
        or fact.get("formula") != "=B2"
        or not isinstance(details, dict)
        or set(details) != {"before", "after", "number_format_definition_material_changed"}
        or details.get("number_format_definition_material_changed") is not True
    ):
        return False
    expected = {
        "present": True,
        "default_format_override_count": 0,
        "cell_format_assignment_count": 1,
        "row_format_assignment_count": 0,
        "column_format_assignment_count": 0,
        "built_in_format_assignment_count": 0,
        "custom_format_assignment_count": 1,
        "unrecognized_number_format_count": 0,
    }
    return details.get("before") == expected and details.get("after") == expected


def _number_format_visibility_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's stored number-format change record."""

    return change.get("kind") == "number_format_controls_changed" and (
        _number_format_visibility_details_observed(change.get("details"), fact)
    )


def _number_format_visibility_finding_observed(
    finding: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Require FormulaFence's matching high-severity number-format finding."""

    return finding.get("rule_id") == "FF039" and _number_format_visibility_details_observed(
        finding.get("details"), fact
    )


def _what_if_data_table_input_reference_details_observed(
    details: Any, fact: dict[str, Any]
) -> bool:
    """Match FormulaFence's redacted one-variable Data Table evidence.

    FormulaFence deliberately omits the Data Table's output range and input
    references. WCAB requires the exact structural profile and FF034, while
    its raw validator establishes the generated ``r1`` transition and stable
    table surroundings independently.
    """

    if (
        fact.get("table_sheet") != "Sensitivity"
        or fact.get("master_cell") != "D3"
        or fact.get("output_range") != "D3:D5"
        or fact.get("baseline_input_cell") != "B2"
        or fact.get("candidate_input_cell") != "B3"
        or fact.get("orientation") != "column"
        or fact.get("recalculation_requested") is not True
        or fact.get("input_value_range") != "C3:C5"
        or fact.get("input_values") != [0.04, 0.08, 0.12]
        or fact.get("primary_input_value") != 0.08
        or fact.get("alternate_input_value") != 0.12
        or fact.get("scale_cell") != "B4"
        or fact.get("scale_value") != 100
        or fact.get("output_formula_cell") != "D2"
        or fact.get("output_formula") != "=Model!$B$2"
        or fact.get("model_sheet") != "Model"
        or fact.get("model_cell") != "B2"
        or fact.get("model_formula") != "=Sensitivity!$B$2*Sensitivity!$B$3*Sensitivity!$B$4"
        or fact.get("dashboard_sheet") != "Dashboard"
        or fact.get("dashboard_cell") != "B4"
        or fact.get("dashboard_formula") != "=Model!$B$2"
        or not isinstance(details, dict)
    ):
        return False
    expected_profile = {
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
    return (
        details.get("before") == expected_profile
        and details.get("after") == expected_profile
        and details.get("data_table_definition_material_changed") is True
        and set(details) == {"before", "after", "data_table_definition_material_changed"}
    )


def _what_if_data_table_input_reference_observed(
    change: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Match FormulaFence's What-If Data Table control-change record."""

    return change.get("kind") == "what_if_data_tables_changed" and (
        _what_if_data_table_input_reference_details_observed(change.get("details"), fact)
    )


def _what_if_data_table_input_reference_finding_observed(
    finding: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Require FormulaFence's matching high-severity What-If Data Table finding."""

    return finding.get("rule_id") == "FF034" and (
        _what_if_data_table_input_reference_details_observed(finding.get("details"), fact)
    )


def _scenario_manager_stored_input_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Match FormulaFence's redacted one-scenario stored-input evidence.

    FormulaFence reports only Scenario Manager structure, deliberately omitting
    the scenario name, stored values, input references, comment, and user
    metadata. WCAB's raw validator verifies that one generated stored value
    changed; the adapter requires this exact redacted profile and FF035.
    """

    if (
        fact.get("scenario_sheet") != "Inputs"
        or fact.get("scenario_name") != "WCAB downside"
        or fact.get("changing_cell") != "B2"
        or fact.get("stable_input_cell") != "B3"
        or fact.get("baseline_stored_value") != "0.08"
        or fact.get("candidate_stored_value") != "0.16"
        or fact.get("stable_stored_value") != "125"
        or fact.get("input_number_format_id") != 10
        or fact.get("summary_ref") != "D2"
        or fact.get("worksheet_input_value") != 0.1
        or fact.get("worksheet_stable_input_value") != 125
        or fact.get("result_cell") != "D2"
        or fact.get("result_formula") != "=B2*B3"
        or fact.get("dashboard_sheet") != "Dashboard"
        or fact.get("dashboard_cell") != "B4"
        or fact.get("dashboard_formula") != "=Inputs!$D$2"
        or not isinstance(details, dict)
    ):
        return False
    expected_profile = {
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
    return (
        details.get("before") == expected_profile
        and details.get("after") == expected_profile
        and details.get("scenario_definition_material_changed") is True
        and set(details) == {"before", "after", "scenario_definition_material_changed"}
    )


def _scenario_manager_stored_input_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match FormulaFence's Scenario Manager control-change record."""

    return change.get("kind") == "scenario_manager_changed" and (
        _scenario_manager_stored_input_details_observed(change.get("details"), fact)
    )


def _scenario_manager_stored_input_finding_observed(
    finding: dict[str, Any], fact: dict[str, Any]
) -> bool:
    """Require FormulaFence's matching high-severity Scenario Manager finding."""

    return finding.get("rule_id") == "FF035" and _scenario_manager_stored_input_details_observed(
        finding.get("details"), fact
    )


def _formula_cached_result_details_observed(details: Any, fact: dict[str, Any]) -> bool:
    """Match FormulaFence's intentionally redacted saved-result details.

    FormulaFence does not disclose formula-cell locations or cached values for
    this detector. The adapter therefore requires the exact one-result numeric
    profile and the native material-change count, while WCAB's raw validator
    establishes the declared formula, location, and values independently.
    """

    if (
        fact.get("result_type") != "numeric"
        or not isinstance(fact.get("sheet"), str)
        or not isinstance(fact.get("cell"), str)
        or not isinstance(fact.get("formula"), str)
        or fact.get("baseline_cached_result") == fact.get("candidate_cached_result")
    ):
        return False
    if not isinstance(details, dict):
        return False
    expected_profile = {
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
    return (
        details.get("before") == expected_profile
        and details.get("after") == expected_profile
        and details.get("unexplained_cached_result_change_count") == 1
        and details.get("cached_result_material_changed") is True
    )


def _formula_cached_result_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Match the dedicated FormulaFence change record."""

    return change.get("kind") == "formula_cached_result_changed" and (
        _formula_cached_result_details_observed(change.get("details"), fact)
    )


def _formula_cached_result_finding_observed(finding: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require the corresponding high-severity FormulaFence finding."""

    return finding.get("rule_id") == "FF042" and _formula_cached_result_details_observed(
        finding.get("details"), fact
    )


def _array_formula_mode_observed(change: dict[str, Any], fact: dict[str, Any]) -> bool:
    """Require FormulaFence to identify this exact CSE-to-dynamic transition."""

    if change.get("kind") != "array_formula_mode_changed":
        return False
    sheet = fact.get("sheet")
    cell = fact.get("cell")
    if not isinstance(sheet, str) or not isinstance(cell, str):
        return False
    if change.get("location") != f"{sheet}!{cell}":
        return False
    details = change.get("details")
    if not isinstance(details, dict):
        return False
    before = details.get("before")
    after = details.get("after")
    expected_before = {
        "mode": fact.get("baseline_mode"),
        "output_range": fact.get("baseline_output_range"),
    }
    expected_after = {
        "mode": fact.get("candidate_mode"),
        "output_range": fact.get("candidate_output_range"),
    }
    return before == expected_before and after == expected_after


def _portfolio_entry(report: dict[str, Any], path: str) -> dict[str, Any] | None:
    workbooks = report.get("workbooks", [])
    if not isinstance(workbooks, list):
        raise FormulaFenceAdapterError("FormulaFence portfolio report has no workbooks list")
    for entry in workbooks:
        if isinstance(entry, dict) and entry.get("path") == path:
            return entry
    return None


def _evaluate_portfolio_case(
    directory: Path, truth: dict[str, Any], *, executable: str
) -> dict[str, Any]:
    report = portfolio(directory / "baseline", directory / "candidate", executable=executable)
    matched: list[str] = []
    missed: list[str] = []
    unmapped: list[str] = []
    for fact in truth.get("facts", []):
        kind = fact.get("kind")
        if kind == "portfolio_value_changed":
            entry = _portfolio_entry(report, str(fact.get("workbook")))
            changes = [] if entry is None else entry.get("changes", [])
            location = f"{fact.get('sheet')}!{fact.get('cell')}"
            observed = any(
                isinstance(change, dict)
                and change.get("kind") == "value_changed"
                and change.get("location") == location
                for change in changes
            )
        elif kind == "portfolio_external_reference":
            source_entry = _portfolio_entry(report, str(fact.get("target_workbook")))
            expected_downstream = fact.get("workbook")
            observed = False
            findings = [] if source_entry is None else source_entry.get("findings", [])
            for finding in findings:
                if not isinstance(finding, dict) or finding.get("rule_id") != "FF079":
                    continue
                samples = finding.get("details", {}).get("sample_impacts", [])
                if any(
                    isinstance(sample, dict) and sample.get("workbook") == expected_downstream
                    for sample in samples
                ):
                    observed = True
                    break
        else:
            unmapped.append(str(kind))
            continue
        (matched if observed else missed).append(str(kind))
    return {
        "case_id": truth.get("id"),
        "status": "matched" if not missed else "missed",
        "matched": matched,
        "missed": missed,
        "unmapped": unmapped,
        "matched_coverage_expectations": [],
        "missed_coverage_expectations": [],
        "unmapped_coverage_expectations": [],
        "change_count": report.get("summary", {}).get("change_count"),
    }


def evaluate_diff_case(case_dir: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Compare FormulaFence's diff evidence with mappable WCAB facts.

    Unmappable facts and coverage expectations remain visible in the result
    instead of being silently counted as passes. Portfolio facts are handled
    through FormulaFence's separate portfolio command.
    """

    directory = Path(case_dir)
    truth = _load_truth(directory)
    if truth.get("topology") == "portfolio":
        return _evaluate_portfolio_case(directory, truth, executable=executable)
    if truth.get("topology") != "pair":
        raise FormulaFenceAdapterError(
            f"{directory}: unsupported topology {truth.get('topology')!r}"
        )

    report = diff(directory / "baseline.xlsx", directory / "candidate.xlsx", executable=executable)
    changes = report.get("changes", [])
    if not isinstance(changes, list):
        raise FormulaFenceAdapterError("FormulaFence report has no changes list")
    findings = report.get("findings", [])
    if not isinstance(findings, list):
        raise FormulaFenceAdapterError("FormulaFence report has no findings list")

    matched: list[str] = []
    missed: list[str] = []
    unmapped: list[str] = []
    for fact in truth.get("facts", []):
        kind = fact.get("kind")
        expectation = _FACT_TO_CHANGE.get(kind)
        if expectation is None:
            unmapped.append(str(kind))
            continue
        expected_kind, location_mode = expectation
        expected_location = _fact_location(fact) if location_mode is not None else None
        expected_functions = fact.get("functions")
        observed = any(
            change.get("kind") == expected_kind
            and (expected_location is None or change.get("location") == expected_location)
            and (
                expected_functions is None
                or (
                    isinstance(change.get("details"), dict)
                    and change["details"].get("functions") == expected_functions
                )
            )
            for change in changes
            if isinstance(change, dict)
        )
        if kind == "external_data_connection_refresh_on_load_changed":
            observed = any(
                _external_data_connection_refresh_on_load_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            )
        if kind == "external_workbook_link_update_policy_changed":
            observed = any(
                _external_workbook_link_update_policy_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            )
        if kind == "iterative_calculation_enabled":
            observed = any(
                _iterative_calculation_enabled_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            )
        if kind == "precision_as_displayed_enabled":
            observed = any(
                _precision_as_displayed_enabled_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            )
        if kind == "workbook_date_system_changed":
            observed = any(
                _workbook_date_system_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _workbook_date_system_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "auto_filter_criteria_changed":
            observed = any(
                _auto_filter_criteria_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _auto_filter_criteria_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "data_validation_list_source_changed":
            observed = any(
                _data_validation_list_source_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _data_validation_list_source_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "conditional_formatting_threshold_changed":
            observed = any(
                _conditional_formatting_threshold_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _conditional_formatting_threshold_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "cell_number_format_changed":
            observed = any(
                _number_format_visibility_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _number_format_visibility_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "pivot_cache_refresh_on_load_changed":
            observed = any(
                _pivot_cache_refresh_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _pivot_cache_refresh_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "chart_series_value_reference_changed":
            observed = any(
                _chart_series_value_reference_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _chart_series_value_reference_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "pivot_data_field_aggregation_changed":
            observed = any(
                _pivot_data_field_aggregation_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _pivot_data_field_aggregation_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "pivot_slicer_selection_changed":
            observed = any(
                _pivot_slicer_selection_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _pivot_slicer_selection_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "power_query_m_filter_changed":
            observed = any(
                _power_query_m_filter_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _power_query_m_filter_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "what_if_data_table_input_reference_changed":
            observed = any(
                _what_if_data_table_input_reference_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _what_if_data_table_input_reference_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "scenario_manager_stored_input_value_changed":
            observed = any(
                _scenario_manager_stored_input_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _scenario_manager_stored_input_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "formula_cached_result_changed":
            observed = any(
                _formula_cached_result_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            ) and any(
                _formula_cached_result_finding_observed(finding, fact)
                for finding in findings
                if isinstance(finding, dict)
            )
        if kind == "array_formula_mode_changed":
            observed = any(
                _array_formula_mode_observed(change, fact)
                for change in changes
                if isinstance(change, dict)
            )
        (matched if observed else missed).append(str(kind))

    coverage_expectations = truth.get("coverage_expectations", [])
    if not isinstance(coverage_expectations, list):
        raise FormulaFenceAdapterError(f"{directory}: coverage_expectations must be an array")
    candidate_profile: dict[str, Any] | None = None
    matched_coverage_expectations: list[dict[str, Any]] = []
    missed_coverage_expectations: list[dict[str, Any]] = []
    unmapped_coverage_expectations: list[dict[str, Any]] = []
    for expectation in coverage_expectations:
        if not isinstance(expectation, dict):
            raise FormulaFenceAdapterError(f"{directory}: coverage expectation must be an object")
        expectation_kind = str(expectation.get("kind"))
        rule_mapping = _COVERAGE_EXPECTATION_TO_RULE.get(expectation_kind)
        if rule_mapping is not None:
            expected_rule, location_mode = rule_mapping
            expected_location = _fact_location(expectation) if location_mode is not None else None
            expected_functions = expectation.get("functions")
            observed = any(
                isinstance(finding, dict)
                and finding.get("rule_id") == expected_rule
                and (expected_location is None or finding.get("location") == expected_location)
                and (
                    expected_functions is None
                    or (
                        isinstance(finding.get("details"), dict)
                        and finding["details"].get("functions") == expected_functions
                    )
                )
                for finding in findings
            )
        else:
            profile_feature = _COVERAGE_EXPECTATION_TO_PROFILE_FEATURE.get(expectation_kind)
            if profile_feature is None:
                unmapped_coverage_expectations.append(expectation)
                continue
            if candidate_profile is None:
                candidate_profile = profile(directory / "candidate.xlsx", executable=executable)
            features = candidate_profile.get("features")
            if not isinstance(features, dict):
                raise FormulaFenceAdapterError("FormulaFence profile has no features object")
            profile_entries = features.get(profile_feature)
            if not isinstance(profile_entries, list):
                raise FormulaFenceAdapterError(
                    f"FormulaFence profile feature {profile_feature!r} is not a list"
                )
            driver_location = _nested_location(expectation, "driver")
            formula_location = _nested_location(expectation, "formula")
            expected_functions = expectation.get("functions")
            driver_changed = any(
                isinstance(change, dict)
                and change.get("kind") == "value_changed"
                and change.get("location") == driver_location
                for change in changes
            )
            profile_observed = any(
                isinstance(entry, dict)
                and entry.get("location") == formula_location
                and entry.get("functions") == expected_functions
                for entry in profile_entries
            )
            observed = driver_changed and profile_observed
        (matched_coverage_expectations if observed else missed_coverage_expectations).append(
            expectation
        )
    return {
        "case_id": truth.get("id"),
        "status": "matched" if not missed else "missed",
        "matched": matched,
        "missed": missed,
        "unmapped": unmapped,
        "matched_coverage_expectations": matched_coverage_expectations,
        "missed_coverage_expectations": missed_coverage_expectations,
        "unmapped_coverage_expectations": unmapped_coverage_expectations,
        "change_count": report.get("summary", {}).get("change_count"),
    }


def evaluate_reference_suite(
    root: str | Path, *, executable: str = "formulafence"
) -> dict[str, Any]:
    """Run the documented FormulaFence reference pass across WCAB fixtures."""

    fixture_root = Path(root)
    cases: list[dict[str, Any]] = []
    lint_results: list[dict[str, Any]] = []
    for truth_path in sorted(fixture_root.rglob("truth.json")):
        case_dir = truth_path.parent
        case_result = evaluate_diff_case(case_dir, executable=executable)
        cases.append(case_result)
        expected_rule = _LINT_EXPECTATIONS.get(str(case_result["case_id"]))
        if expected_rule is not None:
            report = lint(case_dir / "candidate.xlsx", executable=executable)
            findings = report.get("findings", [])
            observed_rules = {
                finding.get("rule_id") for finding in findings if isinstance(finding, dict)
            }
            lint_results.append(
                {
                    "case_id": case_result["case_id"],
                    "expected_rule": expected_rule,
                    "matched": expected_rule in observed_rules,
                    "observed_rules": sorted(
                        rule for rule in observed_rules if isinstance(rule, str)
                    ),
                }
            )
    missed = [case for case in cases if case["missed"]]
    coverage_missed = [case for case in cases if case["missed_coverage_expectations"]]
    lint_missed = [case for case in lint_results if not case["matched"]]
    mapped_fact_count = sum(len(case["matched"]) + len(case["missed"]) for case in cases)
    matched_fact_count = sum(len(case["matched"]) for case in cases)
    unmapped_fact_count = sum(len(case["unmapped"]) for case in cases)
    mapped_coverage_expectation_count = sum(
        len(case["matched_coverage_expectations"]) + len(case["missed_coverage_expectations"])
        for case in cases
    )
    matched_coverage_expectation_count = sum(
        len(case["matched_coverage_expectations"]) for case in cases
    )
    unmapped_coverage_expectation_count = sum(
        len(case["unmapped_coverage_expectations"]) for case in cases
    )
    return {
        "tool": "FormulaFence",
        "case_count": len(cases),
        "matched_diff_case_count": sum(case["status"] == "matched" for case in cases),
        "mapped_diff_fact_count": mapped_fact_count,
        "matched_diff_fact_count": matched_fact_count,
        "unmapped_diff_fact_count": unmapped_fact_count,
        "diff_cases_with_misses": missed,
        "diff_cases": cases,
        "mapped_coverage_expectation_count": mapped_coverage_expectation_count,
        "matched_coverage_expectation_count": matched_coverage_expectation_count,
        "unmapped_coverage_expectation_count": unmapped_coverage_expectation_count,
        "coverage_expectations_with_misses": coverage_missed,
        "lint_rule_count": len(lint_results),
        "lint_rule_misses": lint_missed,
        "lint_rules": lint_results,
    }


def reference_observations(root: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Emit scoreable normalized observations from the local FormulaFence adapter.

    This reuses the adapter's documented mapping from native FormulaFence
    evidence to WCAB facts. A matched fact is copied from WCAB truth only after
    the native change/rule mapping matched; the optional evidence object names
    that native mapping. FormulaFence is an analyzer rather than a policy
    engine, so this adapter intentionally leaves every review disposition null.
    """

    fixture_root = Path(root)
    directories: dict[str, tuple[Path, dict[str, Any]]] = {}
    for truth_path in fixture_root.rglob("truth.json"):
        truth = _load_truth(truth_path.parent)
        case_id = truth.get("id")
        if not isinstance(case_id, str):
            raise FormulaFenceAdapterError(f"{truth_path}: fixture case id must be a string")
        directories[case_id] = (truth_path.parent, truth)
    if set(directories) != set(CASE_IDS):
        raise FormulaFenceAdapterError("fixture tree does not match the WCAB case catalogue")

    cases: list[dict[str, Any]] = []
    for case_id in CASE_IDS:
        directory, truth = directories[case_id]
        result = evaluate_diff_case(directory, executable=executable)
        matched = set(result["matched"])
        facts: list[dict[str, Any]] = []
        for fact in truth.get("facts", []):
            kind = str(fact.get("kind"))
            if kind not in matched:
                continue
            if kind in _PORTFOLIO_FACT_EVIDENCE:
                evidence = _PORTFOLIO_FACT_EVIDENCE[kind]
            else:
                expectation = _FACT_TO_CHANGE.get(kind)
                if expectation is None:
                    continue
                evidence = {"native_change_kind": expectation[0]}
            facts.append({"evidence": evidence, "fact": fact})
        coverage_notes: list[str] = []
        if result["unmapped"]:
            coverage_notes.append(
                "FormulaFence adapter mapping intentionally unavailable for: "
                + ", ".join(sorted(result["unmapped"]))
            )
        unmapped_coverage_expectations = result.get("unmapped_coverage_expectations", [])
        if unmapped_coverage_expectations:
            coverage_notes.append(
                "FormulaFence adapter coverage mapping intentionally unavailable for: "
                + ", ".join(
                    sorted(
                        str(expectation.get("kind"))
                        for expectation in unmapped_coverage_expectations
                        if isinstance(expectation, dict)
                    )
                )
            )
        coverage_declarations: list[dict[str, Any]] = []
        for expectation in result.get("matched_coverage_expectations", []):
            if not isinstance(expectation, dict):
                continue
            evidence = _coverage_evidence(expectation)
            if evidence is None:
                continue
            coverage_declarations.append(
                {
                    "evidence": evidence,
                    "expectation": expectation,
                }
            )
        cases.append(
            {
                "coverage": {"declarations": coverage_declarations, "notes": coverage_notes},
                "facts": facts,
                "id": case_id,
                "review": None,
                "status": "analyzed",
            }
        )
    return {
        "benchmark": {"fixture_schema_versions": [FIXTURE_SCHEMA_VERSION]},
        "cases": cases,
        "schema_version": OBSERVATION_SCHEMA_VERSION,
        "tool": {"name": "FormulaFence"},
    }
