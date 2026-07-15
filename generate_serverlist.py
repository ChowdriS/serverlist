"""
Builds a serverlist output file (same shape as ExampleSL.xlsx) by combining:
  - a user-uploaded per-wave input file (serverlist/inputFile/{wave_name}.xlsx)
  - rightSizing.xlsx  (sheet "in")               -> OS family/version lookup
  - phoenixtracker.xlsx (sheet "ServerTracker(working)") -> FQDN/region/tags/UEFI lookup

Run:
    python generate_serverlist.py
It will prompt for wave_name and environment (dev/stage/prod).
"""

import os
import re
import sys

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "inputFile")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
RIGHTSIZING_FILE = os.path.join(BASE_DIR, "rightSizing.xlsx")
PHOENIXTRACKER_FILE = os.path.join(BASE_DIR, "phoenixtracker.xlsx")

RIGHTSIZING_SHEET = "in"
PHOENIXTRACKER_SHEET = "ServerTracker(working)"

# Hardcoded per-environment AWS account IDs.
ACCOUNT_ID_MAP = {
    "dev": "447648296582",
    "stage": "726725834586",
    "prod": "355564824768",
}

OUTPUT_COLUMNS = [
    "wave_name", "app_name", "aws_region", "aws_accountid", "server_name",
    "server_os_family", "server_os_version", "server_fqdn", "server_tier",
    "server_environment", "r_type", "subnet_IDs", "securitygroup_IDs",
    "subnet_IDs_test", "securitygroup_IDs_test", "instanceType", "tenancy",
    "tags", "private_ip", "iamRole", "server_boot_mode_uefi",
    "ebs_kms_key_id", "ebs_encrypted",
]

# Columns whose real logic hasn't been defined yet - left blank for now.
PENDING_COLUMNS = [
    "subnet_IDs", "securitygroup_IDs", "subnet_IDs_test",
    "securitygroup_IDs_test", "instanceType", "private_ip", "ebs_kms_key_id",
]


def _norm(value) -> str:
    """Uppercases and strips a hostname-like value for case-insensitive matching."""
    if value is None:
        return ""
    return str(value).strip().upper()


def load_rightsizing_lookup() -> dict:
    """Maps normalized hostname -> {os_family, os_version} from rightSizing.xlsx."""
    df = pd.read_excel(RIGHTSIZING_FILE, sheet_name=RIGHTSIZING_SHEET)
    lookup = {}
    for _, row in df.iterrows():
        key = _norm(row.get("Name"))
        if not key:
            continue
        lookup[key] = {
            "os_family": row.get("OS Name"),
            "os_version": row.get("OS Description"),
        }
    return lookup


def load_phoenixtracker_lookup() -> dict:
    """Maps normalized hostname -> {fqdn, region, tags, uefi} from phoenixtracker.xlsx.

    Row 1 of this sheet is a title row, so the real header is row 2 (header=1).
    """
    df = pd.read_excel(PHOENIXTRACKER_FILE, sheet_name=PHOENIXTRACKER_SHEET, header=1)
    lookup = {}
    for _, row in df.iterrows():
        key = _norm(row.get("f"))
        if not key:
            continue
        lookup[key] = {
            "fqdn": row.get("FQDN"),
            "region": row.get("Region"),
            "tags": row.get("Tags"),
            "uefi_enabled": row.get("UEFI Enabled"),
            "app_id": row.get("App ID"),
        }
    return lookup


def dedupe_by_phoenix_app_id(input_df: pd.DataFrame, phoenix_lookup: dict) -> pd.DataFrame:
    """When one server_name has multiple input rows (different app_name/app_id),
    keep only the row whose app_id matches phoenixtracker's App ID for that server -
    that's the row actually aligned to this server, per the source data. Drops the rest.
    If phoenixtracker has no App ID for the server, or none/multiple rows match, keeps
    every candidate row and prints a warning instead of guessing."""
    keep_rows = []
    for server_name, group in input_df.groupby("servername", sort=False):
        if len(group) == 1:
            keep_rows.append(group.iloc[0])
            continue

        key = _norm(server_name)
        px_app_id = phoenix_lookup.get(key, {}).get("app_id")
        if px_app_id is None or (isinstance(px_app_id, float) and pd.isna(px_app_id)):
            print(f"WARNING: {server_name!r} has {len(group)} candidate rows and no App ID "
                  f"in phoenixtracker to disambiguate - keeping all of them")
            keep_rows.extend(r for _, r in group.iterrows())
            continue

        px_app_id_norm = str(px_app_id).strip().upper()
        matches = group[group["app_id"].astype(str).str.strip().str.upper() == px_app_id_norm]

        if len(matches) == 1:
            dropped = len(group) - 1
            print(f"INFO: {server_name!r} had {len(group)} candidate rows - kept app_id="
                  f"{px_app_id!r} (phoenixtracker match), dropped {dropped} other(s)")
            keep_rows.append(matches.iloc[0])
        elif len(matches) > 1:
            print(f"WARNING: {server_name!r} has {len(matches)} rows all matching phoenixtracker "
                  f"app_id {px_app_id!r} - keeping all of them, cannot disambiguate further")
            keep_rows.extend(r for _, r in matches.iterrows())
        else:
            print(f"WARNING: {server_name!r} has {len(group)} candidate rows but none match "
                  f"phoenixtracker app_id {px_app_id!r} - keeping all of them")
            keep_rows.extend(r for _, r in group.iterrows())

    return pd.DataFrame(keep_rows, columns=input_df.columns).reset_index(drop=True)


def resolve_uefi_flag(raw_value) -> object:
    """'efi'-ish values -> True, 'bios'-ish values -> False, unknown/missing -> None."""
    if raw_value is None or (isinstance(raw_value, float) and pd.isna(raw_value)):
        return None
    text = str(raw_value).strip().lower()
    if "efi" in text:
        return True
    if "bios" in text:
        return False
    return None


_SANITIZE_RE = re.compile(r"\s*[|-]\s*")


def sanitize(value) -> object:
    """server_name, app_name, and tags must not contain '|' or '-' - replaced with '_',
    also absorbing any surrounding spaces (e.g. 'microsoft | ar23990' -> 'microsoft_ar23990')."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return value
    return _SANITIZE_RE.sub("_", str(value))


def load_input_file(wave_name: str) -> pd.DataFrame:
    path = os.path.join(INPUT_DIR, f"{wave_name}.xlsx")
    if not os.path.isfile(path):
        print(f"ERROR: input file not found: {path}")
        sys.exit(1)
    df = pd.read_excel(path)
    required = {"servername", "app_name", "app_id", "app_tier"}
    missing = required - set(df.columns)
    if missing:
        print(f"ERROR: input file is missing expected columns: {sorted(missing)}")
        sys.exit(1)
    return df


def build_serverlist(wave_name: str, environment: str) -> pd.DataFrame:
    input_df = load_input_file(wave_name)
    rightsizing_lookup = load_rightsizing_lookup()
    phoenix_lookup = load_phoenixtracker_lookup()
    input_df = dedupe_by_phoenix_app_id(input_df, phoenix_lookup)
    account_id = ACCOUNT_ID_MAP[environment]

    rows = []
    for _, in_row in input_df.iterrows():
        server_name = in_row.get("servername")
        key = _norm(server_name)

        rs = rightsizing_lookup.get(key, {})
        px = phoenix_lookup.get(key, {})

        if key and key not in rightsizing_lookup:
            print(f"WARNING: {server_name!r} not found in rightSizing - OS fields left blank")
        if key and key not in phoenix_lookup:
            print(f"WARNING: {server_name!r} not found in phoenixtracker - FQDN/region/tags/UEFI left blank")

        row = {
            "wave_name": wave_name,
            "app_name": sanitize(in_row.get("app_name")),
            "aws_region": px.get("region"),
            "aws_accountid": account_id,
            "server_name": sanitize(server_name),
            "server_os_family": rs.get("os_family"),
            "server_os_version": rs.get("os_version"),
            "server_fqdn": px.get("fqdn"),
            "server_tier": in_row.get("app_tier"),
            "server_environment": environment,
            "r_type": "Rehost",
            "subnet_IDs": None,
            "securitygroup_IDs": None,
            "subnet_IDs_test": None,
            "securitygroup_IDs_test": None,
            "instanceType": None,
            "tenancy": "Shared",
            "tags": sanitize(px.get("tags")),
            "private_ip": None,
            "iamRole": "AonEC2DefaultRole",
            "server_boot_mode_uefi": resolve_uefi_flag(px.get("uefi_enabled")),
            "ebs_kms_key_id": None,
            "ebs_encrypted": True,
        }
        rows.append(row)

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


def main():
    wave_name = input("Enter wave_name: ").strip()
    environment = input("Enter environment (dev/stage/prod): ").strip().lower()

    if environment not in ACCOUNT_ID_MAP:
        print(f"ERROR: environment must be one of {list(ACCOUNT_ID_MAP)}, got {environment!r}")
        sys.exit(1)

    out_df = build_serverlist(wave_name, environment)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f"{wave_name}_serverlist.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Sheet1", index=False)

    if PENDING_COLUMNS:
        print(f"NOTE: these columns are still placeholders (logic not defined yet): {PENDING_COLUMNS}")
    print(f"Wrote {len(out_df)} rows to {out_path}")


if __name__ == "__main__":
    main()
