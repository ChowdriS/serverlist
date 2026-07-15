"""
Builds a serverlist output file (same shape as ExampleSL.xlsx) by combining:
  - a user-uploaded per-wave input file (serverlist/inputFile/{wave_name}.xlsx)
  - rightSizing.xlsx  (sheet "in")               -> OS family/version lookup
  - phoenixtracker.xlsx (sheet "ServerTracker(working)") -> FQDN/region/domain/facing/tags/UEFI lookup
  - subnet.xlsx (sheet "Sheet1")                 -> subnet_IDs lookup, keyed on
    (environment, region, domain, facing) from phoenixtracker
  - sg_mapping.csv                               -> securitygroup_IDs lookup, keyed on
    (environment, region, domain, facing, tier) from phoenixtracker + the input file's app_tier

Run:
    python generate_serverlist.py
It will prompt for wave_name and environment (dev/stage/prod).
"""

import collections
import os
import re
import sys

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "inputFile")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
RIGHTSIZING_FILE = os.path.join(BASE_DIR, "rightSizing.xlsx")
PHOENIXTRACKER_FILE = os.path.join(BASE_DIR, "tracker.xlsx")
SUBNET_FILE = os.path.join(BASE_DIR, "subnet.xlsx")
SG_FILE = os.path.join(BASE_DIR, "sg_mapping.csv")

RIGHTSIZING_SHEET = "in"
PHOENIXTRACKER_SHEET = "ServerTracker(working)"

# Hardcoded per-environment AWS account IDs.
ACCOUNT_ID_MAP = {
    "dev": "447648296582",
    "stage": "726725834586",
    "prod": "355564824768",
}

# subnet.xlsx's "env" column uses QA/DEV/Prod, not our dev/stage/prod input - this maps
# between the two. "stage" -> "QA" is a best guess (no third env name matched) - correct
# if subnet.xlsx's env taxonomy turns out to mean something else.
ENV_TO_SUBNET_ENV = {
    "dev": "DEV",
    "stage": "QA",
    "prod": "PROD",
}

# Hardcoded per-(environment, region) test subnet/SG, from the aws_lookup.py run
# across dev/stage/prod x eu-central-1/eu-west-2/us-east-1.
TEST_SUBNET_MAP = {
    ("dev", "eu-central-1"): "subnet-077c6c459d5278082",
    ("dev", "eu-west-2"): "subnet-0eddd1e9363b119f8",
    ("dev", "us-east-1"): "subnet-0eda0c24afa95b976",
    ("stage", "eu-central-1"): "subnet-0119bbd31d4f8ba98",
    ("stage", "eu-west-2"): "subnet-091da59e75b333952",
    ("stage", "us-east-1"): "subnet-0cc5c3e3a43596e5c",
    ("prod", "eu-central-1"): "subnet-07f050523062e9102",
    ("prod", "eu-west-2"): "subnet-076e1004e56eb10e0",
    ("prod", "us-east-1"): "subnet-04a7797ffda4e08d0",
}
TEST_SG_MAP = {
    ("dev", "eu-central-1"): "sg-0b7135bad9588e836",
    ("dev", "eu-west-2"): "sg-094dc13866b9884e3",
    ("dev", "us-east-1"): "sg-0127714ef3e32e742",
    ("stage", "eu-central-1"): "sg-01ae6893030c17f76",
    ("stage", "eu-west-2"): "sg-0735e412fac63c3e9",
    ("stage", "us-east-1"): "sg-0944e1a65155a5421",
    ("prod", "eu-central-1"): "sg-07f7b27a3c8ecba22",
    ("prod", "eu-west-2"): "sg-08e5e73534517ca21",
    ("prod", "us-east-1"): "sg-060772fff174b3777",
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
    "instanceType", "private_ip", "ebs_kms_key_id",
]


def _norm(value) -> str:
    """Uppercases and strips a hostname-like value for case-insensitive matching."""
    if value is None:
        return ""
    return str(value).strip().upper()


def _norm_facing(value) -> str:
    """Normalizes 'internal_facing'/'internal'/'external_facing'/'external' -> 'INTERNAL'/'EXTERNAL'."""
    text = _norm(value)
    if "INTERNAL" in text:
        return "INTERNAL"
    if "EXTERNAL" in text:
        return "EXTERNAL"
    return text


def _resolve_sheet_name(path: str, expected_name: str) -> str:
    """Finds the sheet matching expected_name ignoring spacing/case differences -
    real source files sometimes have 'ServerTracker (working)' vs 'ServerTracker(working)'."""
    def norm(s: str) -> str:
        return re.sub(r"\s+", "", s).lower()

    target = norm(expected_name)
    sheet_names = pd.ExcelFile(path).sheet_names
    for name in sheet_names:
        if norm(name) == target:
            return name
    print(f"ERROR: no sheet resembling {expected_name!r} found in {path}. Available sheets: {sheet_names}")
    sys.exit(1)


def load_rightsizing_lookup() -> dict:
    """Maps normalized hostname -> {os_family, os_version} from rightSizing.xlsx."""
    sheet_name = _resolve_sheet_name(RIGHTSIZING_FILE, RIGHTSIZING_SHEET)
    df = pd.read_excel(RIGHTSIZING_FILE, sheet_name=sheet_name)
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
    """Maps normalized hostname -> {fqdn, region, domain, facing, tags, uefi, app_id}
    from phoenixtracker.xlsx. domain/facing feed the subnet.xlsx lookup below.

    Row 1 of this sheet is a title row, so the real header is row 2 (header=1).
    """
    sheet_name = _resolve_sheet_name(PHOENIXTRACKER_FILE, PHOENIXTRACKER_SHEET)
    df = pd.read_excel(PHOENIXTRACKER_FILE, sheet_name=sheet_name, header=1)
    lookup = {}
    for _, row in df.iterrows():
        key = _norm(row.get("f"))
        if not key:
            continue
        lookup[key] = {
            "fqdn": row.get("FQDN"),
            "region": row.get("Region"),
            "domain": row.get("Domain"),
            "facing": row.get("Internal/External"),
            "tags": row.get("Tags"),
            "uefi_enabled": row.get("UEFI Enabled"),
            "app_id": row.get("App ID"),
        }
    return lookup


def load_subnet_lookup() -> dict:
    """Maps (env, region, domain, facing) -> list of subnet IDs from subnet.xlsx.
    Multiple subnet IDs can share one combo (comma-separated in the "subnets" column)."""
    df = pd.read_excel(SUBNET_FILE)
    lookup = {}
    for _, row in df.iterrows():
        key = (
            _norm(row.get("env")),
            _norm(row.get("region")),
            _norm(row.get("domain")),
            _norm_facing(row.get("facing")),
        )
        subnets_raw = row.get("subnets")
        if subnets_raw is None or (isinstance(subnets_raw, float) and pd.isna(subnets_raw)):
            continue
        lookup[key] = [s.strip() for s in str(subnets_raw).split(",") if s.strip()]
    return lookup


def resolve_subnet_id(environment: str, px: dict, subnet_lookup: dict,
                       subnet_usage_counter: dict, server_name) -> object:
    """Picks a subnet ID for subnet_IDs, matched on (environment, region, domain, facing)
    from the caller's phoenixtracker row. When a combo has multiple candidate subnets,
    cycles through them round-robin (via subnet_usage_counter) so servers sharing the
    same combo get spread evenly across those subnets instead of all landing on the
    first one. Does not touch subnet_IDs_test."""
    subnet_env = ENV_TO_SUBNET_ENV.get(environment)
    key = (
        _norm(subnet_env),
        _norm(px.get("region")),
        _norm(px.get("domain")),
        _norm_facing(px.get("facing")),
    )
    candidates = subnet_lookup.get(key)
    if not candidates:
        print(f"WARNING: {server_name!r} - no subnet.xlsx match for env={subnet_env!r} "
              f"region={px.get('region')!r} domain={px.get('domain')!r} facing={px.get('facing')!r}")
        return None
    index = subnet_usage_counter[key] % len(candidates)
    subnet_usage_counter[key] += 1
    return candidates[index]


def resolve_test_value(test_map: dict, environment: str, px: dict, server_name, field_label: str) -> object:
    """Looks up TEST_SUBNET_MAP/TEST_SG_MAP by (environment, region), region taken
    from the caller's phoenixtracker row. Warns instead of guessing if unmatched."""
    region = px.get("region")
    region_key = str(region).strip().lower() if region and not (isinstance(region, float) and pd.isna(region)) else ""
    value = test_map.get((environment, region_key))
    if value is None:
        print(f"WARNING: {server_name!r} - no {field_label} for env={environment!r} region={region!r}")
    return value


def load_sg_lookup() -> dict:
    """Maps (env, region, domain, facing, tier) -> security_group_ids string from sg_mapping.csv."""
    df = pd.read_csv(SG_FILE)
    lookup = {}
    for _, row in df.iterrows():
        key = (
            str(row.get("env")).strip().lower(),
            _norm(row.get("region")),
            _norm(row.get("domain")),
            _norm_facing(row.get("facing")),
            str(row.get("tier")).strip().lower(),
        )
        lookup[key] = row.get("security_group_ids")
    return lookup


def resolve_security_group_ids(environment: str, px: dict, tier, sg_lookup: dict, server_name) -> object:
    """Looks up securitygroup_IDs from sg_mapping.csv, matched on
    (environment, region, domain, facing, tier) - tier is the row's app/db server_tier."""
    tier_key = str(tier).strip().lower() if tier else ""
    key = (
        str(environment).strip().lower(),
        _norm(px.get("region")),
        _norm(px.get("domain")),
        _norm_facing(px.get("facing")),
        tier_key,
    )
    value = sg_lookup.get(key)
    if value is None:
        print(f"WARNING: {server_name!r} - no sg_mapping.csv match for env={environment!r} "
              f"region={px.get('region')!r} domain={px.get('domain')!r} facing={px.get('facing')!r} tier={tier!r}")
    return value


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
    subnet_lookup = load_subnet_lookup()
    subnet_usage_counter = collections.defaultdict(int)
    sg_lookup = load_sg_lookup()
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
            "subnet_IDs": resolve_subnet_id(environment, px, subnet_lookup, subnet_usage_counter, server_name),
            "securitygroup_IDs": resolve_security_group_ids(environment, px, in_row.get("app_tier"), sg_lookup, server_name),
            "subnet_IDs_test": resolve_test_value(TEST_SUBNET_MAP, environment, px, server_name, "subnet_IDs_test"),
            "securitygroup_IDs_test": resolve_test_value(TEST_SG_MAP, environment, px, server_name, "securitygroup_IDs_test"),
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
