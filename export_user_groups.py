import json
import subprocess
import os
import configparser
from openpyxl import Workbook


def run_oci_command(cmd_args, allow_empty=False):
    result = subprocess.run(cmd_args, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"❌ Command failed:\n{result.stderr.strip()}")
        exit(1)
    if not result.stdout.strip():
        if allow_empty:
            return None
        print(f"❌ OCI CLI returned no output for command: {' '.join(cmd_args)}")
        exit(1)
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as e:
        print("❌ Failed to parse JSON output.")
        print("Output was:", result.stdout.strip())
        print("Error:", e)
        exit(1)


def get_tenancy_ocid_from_config():
    config_path = os.path.expanduser("~/.oci/config")
    if not os.path.exists(config_path):
        print("❌ OCI config file not found at ~/.oci/config")
        exit(1)
    config = configparser.ConfigParser()
    config.read(config_path)
    if 'DEFAULT' not in config or 'tenancy' not in config['DEFAULT']:
        print("❌ Tenancy OCID not found in ~/.oci/config under DEFAULT profile")
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
        ], allow_empty=True)
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
    print(f"\n🔍 Fetching domains for compartment: {compartment_ocid}")
    domain_data = run_oci_command([
        "oci", "iam", "domain", "list",
        "--compartment-id", compartment_ocid
    ], allow_empty=True)
    return domain_data.get("data", []) if domain_data else []


def fetch_users_for_domain(domain_url):
    print(f"📥 Fetching users from domain endpoint: {domain_url}")
    return run_oci_command([
        "oci", "identity-domains", "users", "list",
        "--endpoint", domain_url,
        "--attribute-sets", "all",
        "--all",
        "--output", "json"
    ], allow_empty=True)


def main():
    print("🔍 Reading tenancy OCID from ~/.oci/config ...")
    tenancy_ocid = get_tenancy_ocid_from_config()
    print(f"✅ Tenancy OCID: {tenancy_ocid}")

    print("\n📂 Fetching all compartments (including root)...")
    compartments = fetch_compartments(tenancy_ocid)
    print(f"✅ Found {len(compartments)} compartments.")

    wb = Workbook()
    ws = wb.active
    ws.title = "User Groups"
    ws.append(['Domain', 'Username', 'Group', 'Status'])

    json_files = []

    for comp_ocid, comp_name in compartments.items():
        domains = fetch_domains_for_compartment(comp_ocid)
        if not domains:
            print(f"⚠️ No domains found in compartment: {comp_ocid} ({comp_name})")
            continue

        for domain in domains:
            domain_name = domain.get("display-name")
            domain_url = domain.get("url")
            if not domain_url:
                print(f"⚠️ No URL found for domain {domain_name} in compartment {comp_name}")
                continue

            try:
                user_data = fetch_users_for_domain(domain_url)
            except Exception as e:
                print(f"⚠️ Error fetching users for domain {domain_name}: {e}")
                continue

            json_filename = f"users_{domain_name.replace(' ', '_')}.json"
            with open(json_filename, "w") as jf:
                json.dump(user_data, jf, indent=2)
            json_files.append(json_filename)

            for user in user_data.get('data', {}).get('resources', []):
                username = user.get('user-name')
                groups = user.get('groups', [])

                active_field = user.get('active')
                if active_field is True:
                    status = "Active"
                elif active_field is False:
                    status = "Inactive"
                else:
                    status = "Unknown"

                if groups:
                    for group in groups:
                        ws.append([domain_name, username, group.get('display'), status])
                else:
                    ws.append([domain_name, username, '', status])

    excel_file = "user_groups.xlsx"
    wb.save(excel_file)
    print(f"\n✅ All done. Output saved to '{excel_file}'")

    # Cleanup temporary JSON files
    for f in json_files:
        try:
            os.remove(f)
            print(f"Deleted temporary file: {f}")
        except Exception as e:
            print(f"Failed to delete {f}: {e}")


if __name__ == "__main__":
    main()
