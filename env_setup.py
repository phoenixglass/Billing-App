#!/usr/bin/env python3
"""
Generate PBKDF2 password hashes for BILLING_APP_CREDENTIALS environment variable.

Usage:
    python3 env_setup.py

This script will:
1. Prompt for each username and password
2. Generate PBKDF2-HMAC-SHA256 hashes (600,000 iterations)
3. Output a JSON string ready to use in .env file
"""

import json
import sys
from getpass import getpass
from pathlib import Path

# Import hash_password from app.py
sys.path.insert(0, str(Path(__file__).parent))
from app import hash_password, PBKDF2_ITERATIONS


def main():
    print("=" * 70)
    print("BILLING APP CREDENTIAL SETUP")
    print("=" * 70)
    print(f"\nGenerating PBKDF2 hashes with {PBKDF2_ITERATIONS:,} iterations")
    print("Passwords will NOT be displayed or stored.\n")

    credentials = {}

    while True:
        username = input("Username (or press Enter to finish): ").strip()
        if not username:
            break

        if username in credentials:
            print(f"  ✗ User '{username}' already added")
            continue

        # Get password twice to confirm
        while True:
            pwd1 = getpass(f"  Password for '{username}': ")
            pwd2 = getpass(f"  Confirm password: ")
            if pwd1 == pwd2:
                break
            print("  ✗ Passwords don't match, try again")

        if not pwd1:
            print("  ✗ Password cannot be empty")
            continue

        hash_val = hash_password(pwd1)
        credentials[username] = hash_val
        print(f"  ✓ User '{username}' added\n")

    if not credentials:
        print("\n✗ No users added. Exiting.")
        return

    print("\n" + "=" * 70)
    print("ADD THIS TO YOUR .env FILE:")
    print("=" * 70)
    env_var = json.dumps(credentials)
    print(f"\nBILLING_APP_CREDENTIALS='{env_var}'\n")

    print("=" * 70)
    print("SAVE THIS IN A SECURE LOCATION:")
    print("=" * 70)
    print("1. Copy the BILLING_APP_CREDENTIALS line above")
    print("2. Paste into .env (already in .gitignore)")
    print("3. Run: source .env && streamlit run app.py")
    print("4. Delete any terminal history showing the credentials")
    print("5. DO NOT commit .env to git")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    main()
