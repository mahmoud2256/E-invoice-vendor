"""
auth.py
-------
Lightweight username/password authentication for the Streamlit app.
Users are stored in a local JSON file (users.json) next to this script.

NOTE ON PERSISTENCE: on Streamlit Community Cloud the filesystem is reset
whenever the app restarts/redeploys, so users added at runtime through
"Manage Users" will be lost on the next redeploy unless users.json is
committed back to the GitHub repo. For permanent multi-user management
that survives redeploys, move this to a small database (e.g. Supabase,
same approach as the NovaERP app) later on.
"""

import hashlib
import json
from pathlib import Path

USERS_FILE = Path(__file__).parent / "users.json"

DEFAULT_USERS = {
    "Mahmoud": {"password_hash": None, "is_admin": True},
}


def _hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _ensure_users_file():
    if not USERS_FILE.exists():
        users = {
            "Mahmoud": {"password_hash": _hash_password("1234"), "is_admin": True},
        }
        USERS_FILE.write_text(json.dumps(users, indent=2))


def load_users() -> dict:
    _ensure_users_file()
    return json.loads(USERS_FILE.read_text())


def save_users(users: dict):
    USERS_FILE.write_text(json.dumps(users, indent=2))


def verify_login(username: str, password: str) -> bool:
    users = load_users()
    user = users.get(username)
    if not user:
        return False
    return user["password_hash"] == _hash_password(password)


def is_admin(username: str) -> bool:
    users = load_users()
    user = users.get(username, {})
    return bool(user.get("is_admin", False))


def add_user(username: str, password: str, is_admin_user: bool = False) -> bool:
    """Returns False if the username already exists."""
    users = load_users()
    if username in users:
        return False
    users[username] = {"password_hash": _hash_password(password), "is_admin": is_admin_user}
    save_users(users)
    return True


def delete_user(username: str) -> bool:
    users = load_users()
    if username not in users or username == "Mahmoud":
        return False
    del users[username]
    save_users(users)
    return True


def list_usernames() -> list:
    return list(load_users().keys())
