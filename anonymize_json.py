#!/usr/bin/env python3
"""Anonymize downloaded Scoutnet JSON files while preserving member links."""

import argparse
import json
import secrets
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, Set


EXACT_FIELDS = {
    "member_no",
    "first_name",
    "last_name",
    "ssno",
    "nickname",
    "date_of_birth",
    "kid",
    "latest_police_check",
    "pc_course_89",
    "prev_term_due_date",
    "current_term_due_date",
    "created_at",
    "confirmed_at",
    "postcode",
    "unit",
    "group",
    "note",
    "town",
    "country",
}
FIELD_PREFIXES = (
    "address_",
    "email",
    "contact_",
    "extra_info",
    "avatar_",
)


def is_sensitive_field(field: str) -> bool:
    return field in EXACT_FIELDS or field == "extra_emails" or field.startswith(FIELD_PREFIXES)


def collect_member_ids(value: Any, member_ids: Set[str]) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            if key == "member_no" and isinstance(child, dict) and "value" in child:
                member_ids.add(str(child["value"]))
            collect_member_ids(child, member_ids)
    elif isinstance(value, list):
        for child in value:
            collect_member_ids(child, member_ids)


def create_member_id_map(member_ids: Set[str]) -> Dict[str, str]:
    replacements = secrets.SystemRandom().sample(
        range(10_000_000, 100_000_000),
        len(member_ids),
    )
    return dict(zip(member_ids, (str(value) for value in replacements)))


def create_member_profiles(member_ids: Set[str]) -> Dict[str, Dict[str, str]]:
    member_id_map = create_member_id_map(member_ids)
    profiles = {}
    first_birth_date = date(1800, 1, 1)
    for index, member_id in enumerate(sorted(member_ids)):
        anonymized_id = member_id_map[member_id]
        birth_date = first_birth_date + timedelta(days=index)
        profiles[member_id] = {
            "member_no": anonymized_id,
            "first_name": "Scout",
            "last_name": f"Member-{anonymized_id}",
            "email": f"member-{anonymized_id}@example.invalid",
            "contact_alt_email": f"alt-{anonymized_id}@example.invalid",
            "date_of_birth": birth_date.isoformat(),
            "ssno": f"{birth_date:%Y%m%d}-{index % 10000:04d}",
            "pc_course_89": (date(1800, 1, 1) + timedelta(days=index)).isoformat(),
            "prev_term_due_date": (date(1800, 1, 1) + timedelta(days=index)).isoformat(),
            "current_term_due_date": (date(1800, 1, 1) + timedelta(days=index)).isoformat(),
            "created_at": (date(1800, 1, 1) + timedelta(days=index)).isoformat(),
            "confirmed_at": (date(1800, 1, 1) + timedelta(days=index)).isoformat(),
            "postcode": f"{10000 + index:05d}"[:3] + " " + f"{10000 + index:05d}"[3:],
        }
    return profiles


def collect_unit_values(value: Any, unit_values: Set[str]) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            if key == "unit" and isinstance(child, dict) and "raw_value" in child:
                unit_values.add(str(child["raw_value"]))
            collect_unit_values(child, unit_values)
    elif isinstance(value, list):
        for child in value:
            collect_unit_values(child, unit_values)


def collect_group_values(value: Any, group_values: Set[str]) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            if key == "group" and isinstance(child, dict) and "raw_value" in child:
                group_values.add(str(child["raw_value"]))
            collect_group_values(child, group_values)
    elif isinstance(value, list):
        for child in value:
            collect_group_values(child, group_values)


def create_raw_value_map(raw_values: Set[str]) -> Dict[str, str]:
    sorted_values = sorted(
        raw_values,
        key=lambda value: (not value.isdigit(), int(value) if value.isdigit() else value),
    )
    return {raw_value: str(index) for index, raw_value in enumerate(sorted_values, start=1)}


def replacement_for(
    field: str,
    original: Any,
    member_profiles: Dict[str, Dict[str, str]],
    member_id: str,
    unit_map: Dict[str, str],
    group_map: Dict[str, str],
    list_index: int = 0,
) -> Any:
    if field == "member_no":
        return member_profiles[member_id]["member_no"]
    if field == "note":
        return ""
    if original is None or (isinstance(original, str) and not original.strip()):
        return original
    if field in (
        "first_name",
        "last_name",
        "email",
        "contact_alt_email",
        "ssno",
        "date_of_birth",
        "pc_course_89",
        "prev_term_due_date",
        "current_term_due_date",
        "created_at",
        "confirmed_at",
        "postcode",
    ):
        return member_profiles[member_id][field]
    if field == "unit":
        raw_value = unit_map[str(original.get("raw_value", ""))]
        return {"raw_value": raw_value, "value": f"Test patrull{raw_value}"}
    if field == "group":
        raw_value = group_map[str(original.get("raw_value", ""))]
        return {"raw_value": raw_value, "value": f"Test kår{raw_value}"}
    if field == "extra_emails":
        return f"extra-{member_profiles[member_id]['member_no']}-{list_index}@example.invalid"
    if "email" in field:
        return f"{field}-{member_profiles[member_id]['member_no']}@example.invalid"
    if "date" in field or field == "latest_police_check" or field == "avatar_updated":
        return "1800-01-01"
    if "phone" in field or "mobile" in field or "telephone" in field:
        return "0000000000"
    if field == "ssno":
        return "000000-0000"
    if field == "avatar_url":
        return "https://example.invalid/avatar"
    if field == "town":
        return "Anonymized town"
    if field == "country":
        return "XX"
    if isinstance(original, bool):
        return False
    if isinstance(original, (int, float)):
        return 0
    return "Anonymized"


def anonymize_value(
    field: str,
    value: Any,
    member_profiles: Dict[str, Dict[str, str]],
    member_id: str,
    unit_map: Dict[str, str],
    group_map: Dict[str, str],
) -> Any:
    if isinstance(value, dict):
        if field in ("unit", "group") and "raw_value" in value:
            return replacement_for(
                field, value, member_profiles, member_id, unit_map, group_map
            )
        if "value" in value:
            return {
                key: (
                    anonymize_value(
                        field, child, member_profiles, member_id, unit_map, group_map
                    )
                    if key == "value"
                    else anonymize_object(
                        child, member_profiles, member_id, unit_map, group_map
                    )
                )
                for key, child in value.items()
            }
        return anonymize_object(value, member_profiles, member_id, unit_map, group_map)
    if isinstance(value, list):
        return [
            replacement_for(
                field,
                child,
                member_profiles,
                member_id,
                unit_map,
                group_map,
                index,
            )
            if not isinstance(child, (dict, list))
            else anonymize_value(
                field, child, member_profiles, member_id, unit_map, group_map
            )
            for index, child in enumerate(value)
        ]
    return replacement_for(
        field, value, member_profiles, member_id, unit_map, group_map
    )


def anonymize_object(
    value: Any,
    member_profiles: Dict[str, Dict[str, str]],
    member_id: str,
    unit_map: Dict[str, str],
    group_map: Dict[str, str],
) -> Any:
    if isinstance(value, dict):
        return {
            key: (
                anonymize_roles(
                    child, member_profiles, member_id, unit_map, group_map
                )
                if key == "roles"
                else anonymize_value(
                    key, child, member_profiles, member_id, unit_map, group_map
                )
                if is_sensitive_field(key)
                else anonymize_object(
                    child, member_profiles, member_id, unit_map, group_map
                )
            )
            for key, child in value.items()
        }
    if isinstance(value, list):
        return [
            anonymize_object(child, member_profiles, member_id, unit_map, group_map)
            for child in value
        ]
    return value


def anonymize_roles(
    value: Any,
    member_profiles: Dict[str, Dict[str, str]],
    member_id: str,
    unit_map: Dict[str, str],
    group_map: Dict[str, str],
) -> Any:
    if not isinstance(value, dict):
        return anonymize_object(value, member_profiles, member_id, unit_map, group_map)

    anonymized = {}
    for key, child in value.items():
        if key != "value" or not isinstance(child, dict):
            anonymized[key] = anonymize_object(
                child, member_profiles, member_id, unit_map, group_map
            )
            continue

        role_values = {}
        for category, entities in child.items():
            entity_map = unit_map if category == "troop" else group_map if category == "group" else None
            if entity_map is None or not isinstance(entities, dict):
                role_values[category] = anonymize_object(
                    entities, member_profiles, member_id, unit_map, group_map
                )
                continue

            role_values[category] = {
                entity_map[str(raw_value)]: anonymize_object(
                    roles, member_profiles, member_id, unit_map, group_map
                )
                for raw_value, roles in entities.items()
            }
        anonymized[key] = role_values
    return anonymized


def anonymize_file_data(
    value: Any,
    member_profiles: Dict[str, Dict[str, str]],
    unit_map: Dict[str, str],
    group_map: Dict[str, str],
) -> Any:
    if not isinstance(value, dict) or not isinstance(value.get("data"), dict):
        return value

    anonymized = dict(value)
    anonymized["data"] = {
        member_profiles[str(member_no)]["member_no"]: anonymize_object(
            record, member_profiles, str(member_no), unit_map, group_map
        )
        for member_no, record in value["data"].items()
    }
    return anonymized


def anonymize_customlist_metadata(value: Any, field: str = "") -> Any:
    if isinstance(value, dict):
        return {
            key: anonymize_customlist_metadata(child, key)
            for key, child in value.items()
        }
    if isinstance(value, list):
        return [anonymize_customlist_metadata(child, field) for child in value]
    if field == "owning_body_id":
        return 0
    if isinstance(value, str):
        return ""
    return value


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create anonymized copies of downloaded Scoutnet JSON files."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("json"),
        help="Directory containing the downloaded JSON files (default: json).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("json-anonymized"),
        help="Directory for anonymized copies (default: json-anonymized).",
    )
    args = parser.parse_args()

    input_dir = args.input.resolve()
    output_dir = args.output.resolve()
    if input_dir == output_dir:
        parser.error("Input and output directories must be different.")
    if not input_dir.is_dir():
        parser.error(f"Input directory does not exist: {input_dir}")

    source_files = sorted(input_dir.glob("*.json"))
    if not source_files:
        parser.error(f"No JSON files found in: {input_dir}")

    documents = []
    member_ids: Set[str] = set()
    unit_values: Set[str] = set()
    group_values: Set[str] = set()
    for source in source_files:
        with source.open(encoding="utf-8-sig") as file:
            document = json.load(file)
        documents.append((source.name, document))
        data = document.get("data") if isinstance(document, dict) else None
        if isinstance(data, dict):
            member_ids.update(str(member_no) for member_no in data)
            collect_member_ids(data, member_ids)
        collect_unit_values(document, unit_values)
        collect_group_values(document, group_values)

    member_profiles = create_member_profiles(member_ids)
    unit_map = create_raw_value_map(unit_values)
    group_map = create_raw_value_map(group_values)
    output_dir.mkdir(parents=True, exist_ok=True)
    for filename, document in documents:
        if filename == "customlist.json":
            anonymized = anonymize_customlist_metadata(document)
        else:
            anonymized = anonymize_file_data(
                document, member_profiles, unit_map, group_map
            )
        destination = output_dir / filename
        with destination.open("w", encoding="utf-8", newline="\n") as file:
            json.dump(anonymized, file, ensure_ascii=False, indent=2)
            file.write("\n")

    print(f"Anonymized {len(source_files)} JSON files in {output_dir}")


if __name__ == "__main__":
    main()
