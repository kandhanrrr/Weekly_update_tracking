"""
save_password.py - Securely store your email password in Windows Credential Manager
====================================================================================
Run this ONCE to save your password. After that, weekly_update_tracker.py reads it
automatically. Your password is never stored in any file.

Usage:
    python save_password.py
"""

import sys
import getpass
import win32cred

SERVICE_NAME = "Weekly_Update_Tracker"

def save_password():
    print("=" * 55)
    print("  Weekly Update Tracker - Secure Password Setup")
    print("=" * 55)
    print()
    print("This stores your email password in Windows Credential")
    print("Manager (same secure vault used by Outlook/Chrome).")
    print("Your password will NOT be saved in any text file.")
    print()

    username = input("Enter your email address (smtp_user from config.ini): ").strip()
    if not username:
        print("[ERROR] Email address cannot be empty.")
        sys.exit(1)

    password = getpass.getpass("Enter your email password (input is hidden): ")
    if not password:
        print("[ERROR] Password cannot be empty.")
        sys.exit(1)

    confirm = getpass.getpass("Confirm password: ")
    if password != confirm:
        print("[ERROR] Passwords do not match. Please try again.")
        sys.exit(1)

    # Store in Windows Credential Manager
    credential = {
        "Type":           win32cred.CRED_TYPE_GENERIC,
        "TargetName":     SERVICE_NAME,
        "UserName":       username,
        "CredentialBlob": password,
        "Persist":        win32cred.CRED_PERSIST_LOCAL_MACHINE,
    }
    win32cred.CredWrite(credential, 0)

    print()
    print("[SUCCESS] Password saved securely in Windows Credential Manager.")
    print(f"          Service name : {SERVICE_NAME}")
    print(f"          Username     : {username}")
    print()
    print("You can view/delete it anytime via:")
    print("  Control Panel > Credential Manager > Windows Credentials")
    print()
    print("The tracker script will now read it automatically.")

def verify_password():
    """Check if a password is already stored."""
    try:
        cred = win32cred.CredRead(SERVICE_NAME, win32cred.CRED_TYPE_GENERIC)
        print(f"[INFO] Password already stored for: {cred['UserName']}")
        choice = input("Overwrite it? (y/n): ").strip().lower()
        return choice == "y"
    except Exception:
        return True  # Nothing stored yet — proceed

if __name__ == "__main__":
    if verify_password():
        save_password()
    else:
        print("Keeping existing password. No changes made.")
