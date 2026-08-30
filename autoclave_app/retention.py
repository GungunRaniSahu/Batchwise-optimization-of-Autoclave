"""
Data retention utilities.

Anything older than DATA_RETENTION_DAYS (default 60) is permanently deleted:
  - timestamped Excel backups in BACKUP_DIR
  - the generated Optimized_Batches.xlsx output file

This is deliberate deletion based on file modification time, run on a
schedule (see scheduler.py) and also once at application startup.
"""
import os
import time
import logging

logger = logging.getLogger("retention")


def cleanup_old_files(directory, max_age_days, extra_files=None):
    """Delete files older than max_age_days (by mtime).

    Args:
        directory: folder whose files should be checked (non-recursive).
        max_age_days: age threshold in days.
        extra_files: optional list of individual file paths to also check.

    Returns:
        list of deleted file paths.
    """
    deleted = []
    cutoff = time.time() - (max_age_days * 86400)

    if directory and os.path.isdir(directory):
        for name in os.listdir(directory):
            if name.startswith("."):
                continue
            path = os.path.join(directory, name)
            if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                try:
                    os.remove(path)
                    deleted.append(path)
                except OSError as e:
                    logger.warning("Could not delete %s: %s", path, e)

    if extra_files:
        for path in extra_files:
            if path and os.path.exists(path) and os.path.getmtime(path) < cutoff:
                try:
                    os.remove(path)
                    deleted.append(path)
                except OSError as e:
                    logger.warning("Could not delete %s: %s", path, e)

    if deleted:
        logger.info("Data retention cleanup removed %d file(s) older than %d days.",
                     len(deleted), max_age_days)
    return deleted
