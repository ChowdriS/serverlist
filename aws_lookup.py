"""
Looks up AWS subnets and security groups by name filter (server-side wildcard match)
and prints their Name/ID. Defaults: subnets whose Name tag contains "mgn", security
groups whose name contains "ssm" - matches AWS's own EC2 filter wildcard syntax
(*substring*), so this runs entirely server-side, no local post-filtering.

Requires AWS credentials to already be configured (env vars, ~/.aws/credentials,
SSO, etc.) - this script doesn't handle auth itself.

Run:
    python aws_lookup.py --region eu-central-1 --env dev
    python aws_lookup.py --region eu-central-1 --profile my-profile
    python aws_lookup.py --region eu-central-1 --subnet-filter mgn --sg-filter ssm
    python aws_lookup.py --region eu-central-1 --env prod --output results.xlsx
    python aws_lookup.py --all-regions --all-envs --output results.xlsx
"""

import argparse
import sys

import boto3
import pandas as pd

# Maps our dev/stage/prod naming to the actual AWS CLI/SSO profile name configured
# locally for each environment (~/.aws/credentials or ~/.aws/config on Mac/Linux,
# C:\Users\<you>\.aws\ on Windows). Edit these if your profile names differ.
PROFILE_MAP = {
    "dev": "dev",
    "stage": "stage",
    "prod": "prod",
}


def _tag_name(tags) -> str:
    """Pulls the 'Name' tag value out of an AWS resource's Tags list."""
    for tag in tags or []:
        if tag.get("Key") == "Name":
            return tag.get("Value", "")
    return ""


def find_subnets(ec2_client, name_filter: str) -> list[dict]:
    """Returns [{name, id, vpc_id, cidr}] for subnets whose Name tag matches *name_filter*."""
    resp = ec2_client.describe_subnets(
        Filters=[{"Name": "tag:Name", "Values": [f"*{name_filter}*"]}]
    )
    return [
        {
            "name": _tag_name(s.get("Tags")),
            "id": s["SubnetId"],
            "vpc_id": s.get("VpcId"),
            "cidr": s.get("CidrBlock"),
        }
        for s in resp.get("Subnets", [])
    ]


def find_security_groups(ec2_client, name_filter: str) -> list[dict]:
    """Returns [{name, id, vpc_id}] for security groups whose name matches *name_filter*."""
    resp = ec2_client.describe_security_groups(
        Filters=[{"Name": "group-name", "Values": [f"*{name_filter}*"]}]
    )
    return [
        {
            "name": sg.get("GroupName"),
            "id": sg["GroupId"],
            "vpc_id": sg.get("VpcId"),
        }
        for sg in resp.get("SecurityGroups", [])
    ]


def run_for_region(session: boto3.Session, region: str, subnet_filter: str, sg_filter: str) -> dict:
    ec2 = session.client("ec2", region_name=region)
    return {
        "region": region,
        "subnets": find_subnets(ec2, subnet_filter),
        "security_groups": find_security_groups(ec2, sg_filter),
    }


def print_results(results: list[dict]):
    for result in results:
        region = result["region"]
        env = result.get("env")
        header = f"{env} / {region}" if env else region
        print(f"\n=== {header} ===")

        print(f"Subnets (Name contains {result.get('subnet_filter', '')!r}):")
        if not result["subnets"]:
            print("  (none found)")
        for s in result["subnets"]:
            print(f"  {s['name']:<40} {s['id']}  vpc={s['vpc_id']}  cidr={s['cidr']}")

        print(f"Security groups (name contains {result.get('sg_filter', '')!r}):")
        if not result["security_groups"]:
            print("  (none found)")
        for sg in result["security_groups"]:
            print(f"  {sg['name']:<40} {sg['id']}  vpc={sg['vpc_id']}")


def write_output(results: list[dict], output_path: str):
    subnet_rows = []
    sg_rows = []
    for result in results:
        prefix = {"env": result.get("env"), "region": result["region"]}
        for s in result["subnets"]:
            subnet_rows.append({**prefix, **s})
        for sg in result["security_groups"]:
            sg_rows.append({**prefix, **sg})

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(subnet_rows).to_excel(writer, sheet_name="subnets", index=False)
        pd.DataFrame(sg_rows).to_excel(writer, sheet_name="security_groups", index=False)
    print(f"\nWrote results to {output_path}")


def _regions_for_session(session: boto3.Session, args) -> list[str]:
    if args.all_regions:
        return [r["RegionName"] for r in session.client("ec2", region_name="us-east-1").describe_regions()["Regions"]]
    return args.regions


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--region", action="append", dest="regions",
                         help="AWS region to search. Repeatable. Required unless --all-regions.")
    parser.add_argument("--all-regions", action="store_true",
                         help="Search every region enabled for this account instead of specific --region values.")

    profile_group = parser.add_mutually_exclusive_group()
    profile_group.add_argument("--env", choices=sorted(PROFILE_MAP), default=None,
                                help="Use the AWS profile mapped for this environment (see PROFILE_MAP).")
    profile_group.add_argument("--all-envs", action="store_true",
                                help="Run against dev, stage, and prod profiles in one go.")
    profile_group.add_argument("--profile", default=None,
                                help="AWS named profile to use directly, bypassing --env's dev/stage/prod mapping.")

    parser.add_argument("--subnet-filter", default="mgn", help="Substring to match in subnet Name tag (default: mgn).")
    parser.add_argument("--sg-filter", default="ssm", help="Substring to match in security group name (default: ssm).")
    parser.add_argument("--output", default=None, help="Optional path to write results as .xlsx (sheets: subnets, security_groups).")
    args = parser.parse_args()

    if not args.regions and not args.all_regions:
        print("ERROR: pass at least one --region, or use --all-regions")
        sys.exit(1)

    if args.all_envs:
        envs_to_run = sorted(PROFILE_MAP)
    elif args.env:
        envs_to_run = [args.env]
    else:
        envs_to_run = [None]  # plain --profile (or default profile) run, no env label

    results = []
    for env in envs_to_run:
        profile_name = args.profile if env is None else PROFILE_MAP[env]
        session = boto3.Session(profile_name=profile_name)
        regions = _regions_for_session(session, args)

        for region in regions:
            result = run_for_region(session, region, args.subnet_filter, args.sg_filter)
            result["env"] = env
            result["subnet_filter"] = args.subnet_filter
            result["sg_filter"] = args.sg_filter
            results.append(result)

    print_results(results)

    if args.output:
        write_output(results, args.output)


if __name__ == "__main__":
    main()
