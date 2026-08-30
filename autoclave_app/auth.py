"""
Authentication helpers.

Users are stored in a local JSON file (users.json). Passwords and recovery
PINs are hashed with werkzeug's generate_password_hash — nothing is ever
stored or logged in plaintext.

Password reset works via a "recovery PIN" set at account-creation time
(there's no email server configured for this internal tool). Anyone who
knows a username's recovery PIN can set a new password for that account,
so treat the PIN like a second password.
"""
import json
import os
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin


class User(UserMixin):
    def __init__(self, username, role="operator"):
        self.id = username
        self.username = username
        self.role = role


def load_users(users_file):
    if not os.path.exists(users_file):
        return {}
    try:
        with open(users_file, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_users(users_file, users):
    with open(users_file, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2)


def user_exists(users_file, username):
    return username in load_users(users_file)


def any_users_exist(users_file):
    return len(load_users(users_file)) > 0


def get_user(users_file, username):
    users = load_users(users_file)
    rec = users.get(username)
    if not rec:
        return None
    return User(username, rec.get("role", "operator"))


def verify_user(users_file, username, password):
    users = load_users(users_file)
    rec = users.get(username)
    if not rec:
        return None
    if check_password_hash(rec["password_hash"], password):
        return User(username, rec.get("role", "operator"))
    return None


def create_user(users_file, username, password, role="operator", recovery_pin=None):
    users = load_users(users_file)
    record = {
        "password_hash": generate_password_hash(password),
        "role": role,
    }
    if recovery_pin:
        record["recovery_pin_hash"] = generate_password_hash(recovery_pin)
    users[username] = record
    save_users(users_file, users)


def verify_recovery_pin(users_file, username, pin):
    users = load_users(users_file)
    rec = users.get(username)
    if not rec or "recovery_pin_hash" not in rec:
        return False
    return check_password_hash(rec["recovery_pin_hash"], pin)


def reset_password(users_file, username, new_password):
    users = load_users(users_file)
    if username not in users:
        return False
    users[username]["password_hash"] = generate_password_hash(new_password)
    save_users(users_file, users)
    return True
