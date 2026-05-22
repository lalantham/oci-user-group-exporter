import json
import subprocess
import os
import configparser
from collections import defaultdict
from openpyxl import Workbook


def run_oci_command(cmd_args):
    result = subprocess.run(cmd_args, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Command failed: {result.stderr.strip()}")
        return None
    if not result.stdout.strip():
        return None
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as e:
        print(f"Failed to parse JSON for: {' '.join(cmd_args)}")
        print(f"Error: {e}")
        return None


def get_tenancy_ocid_from_config():
    config_path = os.path.expanduser("~/.oci/config")
    if not os.path.exists(config_path):
        print("OCI config file not found at ~/.oci/config")
        exit(1)
    config = configparser.ConfigParser()
    config.read(config_path)
    if 'DEFAULT' not in config or 'tenancy' not in config['DEFAULT']:
        print("Tenancy OCID not found in ~/.oci/config under DEFAULT profile")
        exit(1)
    return config['DEFAULT']['tenancy']


def fetch_compartments(tenancy_ocid):
    compartments = {}

    def _fetch(compartment_id):
        data = run_oci_command([
            "oci", "iam", "compartment", "list",
            "--compartment-id", compartment_id,
            "--all",
            "--query", 'data[?"lifecycle-state"==`ACTIVE`]'
        ])
        if not data:
            return
        for c in data:
            comp_id = c.get("id")
            comp_name = c.get("name", "Unknown")
            compartments[comp_id] = comp_name
            _fetch(comp_id)

    compartments[tenancy_ocid] = "Root Compartment"
    _fetch(tenancy_ocid)
    return compartments


def fetch_domains_for_compartment(compartment_ocid):
    data = run_oci_command([
        "oci", "iam", "domain", "list",
        "--compartment-id", compartment_ocid
    ])
    return data.get("data", []) if data else []


def fetch_users_for_domain(domain_url):
    return run_oci_command([
        "oci", "identity-domains", "users", "list",
        "--endpoint", domain_url,
        "--attribute-sets", "all",
        "--all",
        "--output", "json"
    ])


def fetch_groups_for_domain(domain_url):
    data = run_oci_command([
        "oci", "identity-domains", "groups", "list",
        "--endpoint", domain_url,
        "--all",
        "--output", "json"
    ])
    if not data:
        return []
    raw = data.get("data") or data
    return raw.get("resources") or raw.get("Resources") or []


def fetch_group_members(domain_url, group_id):
    """
    Fetch members of a specific group using group get with --attributes members.
    This is the only reliable way to get members via OCI CLI.
    """
    data = run_oci_command([
        "oci", "identity-domains", "group", "get",
        "--endpoint", domain_url,
        "--group-id", group_id,
        "--attributes", "members",
        "--output", "json"
    ])
    if not data:
        return []
    return data.get("data", {}).get("members") or []


def build_user_to_groups_map(domain_url):
    user_to_groups = defaultdict(list)

    groups = fetch_groups_for_domain(domain_url)
    print(f"  Groups found: {len(groups)}")

    for group in groups:
        group_name = group.get("display-name") or group.get("displayName") or "Unknown Group"
        group_id = group.get("id")
        if not group_id:
            continue

        members = fetch_group_members(domain_url, group_id)
        for member in members:
            # only process User type members, skip nested groups
            if member.get("type") != "User":
                continue
            member_id = member.get("value")
            if member_id:
                user_to_groups[member_id].append(group_name)

    return user_to_groups


def extract_resources(data):
    if not data:
        return []
    raw = data.get("data") or data
    return raw.get("resources") or raw.get("Resources") or []


def main():
    tenancy_ocid = get_tenancy_ocid_from_config()
    print(f"Tenancy OCID: {tenancy_ocid}")

    compartments = fetch_compartments(tenancy_ocid)
    print(f"Found {len(compartments)} compartments")

    wb = Workbook()
    ws = wb.active
    ws.title = "User Groups"
    ws.append(["Compartment", "Domain", "Username", "Email", "Group", "Status"])

    for comp_ocid, comp_name in compartments.items():
        domains = fetch_domains_for_compartment(comp_ocid)
        if not domains:
            continue

        for domain in domains:
            domain_name = domain.get("display-name", "Unknown")
            domain_url = domain.get("url")
            if not domain_url:
                continue

            print(f"Processing domain: {domain_name}")

            user_data = fetch_users_for_domain(domain_url)
            user_to_groups = build_user_to_groups_map(domain_url)

            users = extract_resources(user_data)
            print(f"  Users: {len(users)}, Users with groups: {len(user_to_groups)}")

            for user in users:
                username = user.get("user-name") or ""
                user_id = user.get("id", "")
                active_field = user.get("active")
                status = (
                    "Active" if active_field is True
                    else "Inactive" if active_field is False
                    else "Unknown"
                )

                emails = user.get("emails") or []
                email = next(
                    (e.get("value") for e in emails if e.get("primary")),
                    ""
                )

                groups = user_to_groups.get(user_id, [])

                if groups:
                    for group_name in groups:
                        ws.append([comp_name, domain_name, username, email, group_name, status])
                else:
                    ws.append([comp_name, domain_name, username, email, "", status])

    output_file = "user_groups.xlsx"
    wb.save(output_file)
    print(f"Output saved to {output_file}")


if __name__ == "__main__":
    main()
