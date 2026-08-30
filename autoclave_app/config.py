import os
from datetime import timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class Config:
    BASE_DIR = BASE_DIR

    # Change this in production. Best practice: set the SECRET_KEY environment
    # variable instead of editing this file.
    SECRET_KEY = os.environ.get("SECRET_KEY", "change-this-secret-key")

    EXCEL_FILE = os.path.join(BASE_DIR, "parts.xlsx")
    BACKUP_DIR = os.path.join(BASE_DIR, "backups")
    SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
    OUTPUT_FILE = os.path.join(BASE_DIR, "Optimized_Batches.xlsx")
    USERS_FILE = os.path.join(BASE_DIR, "users.json")
    LOGO_FILE = os.path.join(BASE_DIR, "static", "img", "tata_logo.png")

    # Data retention: files older than this are auto-deleted (backups + output).
    DATA_RETENTION_DAYS = int(os.environ.get("DATA_RETENTION_DAYS", "60"))

    # How often (hours) the background cleanup job checks for expired files.
    CLEANUP_INTERVAL_HOURS = int(os.environ.get("CLEANUP_INTERVAL_HOURS", "24"))

    PERMANENT_SESSION_LIFETIME = timedelta(hours=8)
