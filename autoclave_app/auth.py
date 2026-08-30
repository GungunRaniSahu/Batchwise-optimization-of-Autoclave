"""
Authentication helpers.

Users are stored in a local JSON file (users.json) with hashed passwords.
No plaintext passwords are ever stored or logged.
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


def create_user(users_file, username, password, role="operator"):
    users = load_users(users_file)
    users[username] = {
        "password_hash": generate_password_hash(password),
        "role": role,
    }
    save_users(users_file, users)
