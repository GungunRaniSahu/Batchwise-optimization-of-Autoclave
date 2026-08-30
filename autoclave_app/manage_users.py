"""
Command-line user management for the Autoclave Batch Optimizer.

Usage:
    python manage_users.py add <username> [--role admin|operator]
    python manage_users.py list
    python manage_users.py remove <username>
"""
import argparse
import getpass

from config import Config
from auth import create_user, load_users, save_users


def main():
    parser = argparse.ArgumentParser(description="Manage Autoclave Optimizer users")
    sub = parser.add_subparsers(dest="cmd", required=True)

    add_p = sub.add_parser("add", help="Add or update a user")
    add_p.add_argument("username")
    add_p.add_argument("--role", default="operator", choices=["admin", "operator"])

    sub.add_parser("list", help="List all users")

    rm_p = sub.add_parser("remove", help="Remove a user")
    rm_p.add_argument("username")

    args = parser.parse_args()

    if args.cmd == "add":
        pw = getpass.getpass("Password: ")
        pw2 = getpass.getpass("Confirm password: ")
        if not pw:
            print("Password cannot be empty.")
            return
        if pw != pw2:
            print("Passwords do not match.")
            return
        pin = getpass.getpass("Recovery PIN (used for 'Forgot password', optional, press Enter to skip): ")
        create_user(Config.USERS_FILE, args.username, pw, args.role, recovery_pin=pin or None)
        print(f"User '{args.username}' saved with role '{args.role}'.")

    elif args.cmd == "list":
        users = load_users(Config.USERS_FILE)
        if not users:
            print("No users found. Create one with: python manage_users.py add <username>")
        for u, rec in users.items():
            print(f"- {u}  (role: {rec.get('role', 'operator')})")

    elif args.cmd == "remove":
        users = load_users(Config.USERS_FILE)
        if args.username in users:
            del users[args.username]
            save_users(Config.USERS_FILE, users)
            print(f"User '{args.username}' removed.")
        else:
            print(f"User '{args.username}' not found.")


if __name__ == "__main__":
    main()
