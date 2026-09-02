#!/usr/bin/env python3
"""
Azure Enum — Azure Active Directory & Resource Enumeration Tool
Enumerates AAD users, groups, service principals, and Azure resource groups.
Requires azure-identity, azure-mgmt-resource, msgraph-sdk: pip install azure-identity azure-mgmt-resource msal
Author: Omar Khalid (amooryx) | github.com/amooryx/azure-enum
AUTHORIZED USE ONLY — for authorized red team engagements and cloud pentests.
"""

import argparse
import json
import sys
import urllib.parse
import urllib.request

try:
    import msal
    HAS_MSAL = True
except ImportError:
    HAS_MSAL = False

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

def get_token_password(tenant_id: str, client_id: str, username: str, password: str) -> str | None:
    if not HAS_MSAL:
        print("[!] msal not installed: pip install msal")
        sys.exit(1)
    app = msal.PublicClientApplication(
        client_id, authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    result = app.acquire_token_by_username_password(
        username=username, password=password,
        scopes=["https://graph.microsoft.com/.default"]
    )
    return result.get("access_token")

def get_token_device(tenant_id: str, client_id: str) -> str | None:
    if not HAS_MSAL:
        sys.exit(1)
    app  = msal.PublicClientApplication(
        client_id, authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])
    print(f"[*] Device flow: {flow['message']}")
    result = app.acquire_token_by_device_flow(flow)
    return result.get("access_token")

def graph_get(endpoint: str, token: str) -> dict | list:
    url = f"{GRAPH_BASE}/{endpoint}"
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type",  "application/json")
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read())
    except Exception as e:
        return {"error": str(e)}

def list_users(token: str) -> list[dict]:
    data = graph_get("users?$select=id,displayName,userPrincipalName,jobTitle,department,accountEnabled&$top=999", token)
    return data.get("value", [])

def list_groups(token: str) -> list[dict]:
    data = graph_get("groups?$select=id,displayName,description,groupTypes&$top=999", token)
    return data.get("value", [])

def list_service_principals(token: str) -> list[dict]:
    data = graph_get("servicePrincipals?$select=id,displayName,appId,servicePrincipalType,appRoles&$top=100", token)
    return data.get("value", [])

def list_apps(token: str) -> list[dict]:
    data = graph_get("applications?$select=id,displayName,appId,passwordCredentials,keyCredentials&$top=100", token)
    return data.get("value", [])

def get_me(token: str) -> dict:
    return graph_get("me", token)

def main():
    parser = argparse.ArgumentParser(
        description="Azure Enum — AAD Enumeration (Authorized use only)",
    )
    parser.add_argument("--tenant",   required=True, help="Azure tenant ID")
    parser.add_argument("--client-id", default="04b07795-8ddb-461a-bbee-02f9e1bf7b46",
                        help="Client ID (default: Azure CLI app)")
    parser.add_argument("--username", help="Username for password flow")
    parser.add_argument("--password", help="Password for password flow")
    parser.add_argument("--device",   action="store_true", help="Use device code flow")
    parser.add_argument("--users",    action="store_true")
    parser.add_argument("--groups",   action="store_true")
    parser.add_argument("--sps",      action="store_true", help="Service principals")
    parser.add_argument("--apps",     action="store_true", help="App registrations")
    parser.add_argument("--me",       action="store_true", help="Current identity")
    parser.add_argument("--all", "-a", action="store_true")
    parser.add_argument("--out",      help="Output JSON file")
    args = parser.parse_args()

    print(f"[*] Authenticating to tenant: {args.tenant}")
    if args.device:
        token = get_token_device(args.tenant, args.client_id)
    elif args.username and args.password:
        token = get_token_password(args.tenant, args.client_id, args.username, args.password)
    else:
        parser.error("Provide --username/--password or --device for authentication")

    if not token:
        print("[!] Authentication failed")
        sys.exit(1)
    print("[+] Token obtained")

    results = {}
    if args.me or args.all:
        me = get_me(token)
        print(f"[+] Signed in as: {me.get('userPrincipalName')} ({me.get('displayName')})")
        results["me"] = me

    if args.users or args.all:
        users = list_users(token)
        print(f"[+] Users: {len(users)}")
        disabled = sum(1 for u in users if not u.get("accountEnabled"))
        print(f"    Enabled: {len(users)-disabled}  Disabled: {disabled}")
        results["users"] = users

    if args.groups or args.all:
        groups = list_groups(token)
        print(f"[+] Groups: {len(groups)}")
        results["groups"] = groups

    if args.sps or args.all:
        sps = list_service_principals(token)
        print(f"[+] Service principals: {len(sps)}")
        results["service_principals"] = sps

    if args.apps or args.all:
        apps = list_apps(token)
        with_creds = [a for a in apps if a.get("passwordCredentials") or a.get("keyCredentials")]
        print(f"[+] App registrations: {len(apps)} ({len(with_creds)} with credentials)")
        for a in with_creds[:5]:
            print(f"    [!!!] {a['displayName']} has stored credentials (passwordCredentials/keyCredentials)")
        results["apps"] = apps

    if args.out:
        with open(args.out, "w") as f:
            json.dump(results, f, indent=2)
        print(f"[*] Results → {args.out}")

if __name__ == "__main__":
    main()
