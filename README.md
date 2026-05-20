# OCI User Group Exporter

Fetches all users and their group memberships across all compartments and identity domains in an OCI tenancy, and exports the results to an Excel file.

## Output

`user_groups.xlsx` with columns:

| Domain | Username | Group | Status |
|---|---|---|---|

## Prerequisites

- Python 3.x
- OCI CLI installed and configured (`~/.oci/config`)
- Permissions to list compartments, domains, and users

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python export_user_groups.py
```

The script will:
1. Read your tenancy OCID from `~/.oci/config`
2. Recursively fetch all active compartments
3. Fetch all identity domains in each compartment
4. Fetch all users and group memberships from each domain
5. Save results to `user_groups.xlsx`
6. Clean up any temporary JSON files

## IAM permissions required

```
Allow group <your-group> to read compartments in tenancy
Allow group <your-group> to read domains in tenancy
Allow group <your-group> to read users in tenancy
```
