"""
Background scheduler that runs the 60-day data retention cleanup
automatically, without needing a user to click anything.
"""
import logging
from apscheduler.schedulers.background import BackgroundScheduler

from config import Config
from retention import cleanup_old_files

logger = logging.getLogger("scheduler")

_scheduler = None


def _run_cleanup():
    cleanup_old_files(
        Config.BACKUP_DIR,
        Config.DATA_RETENTION_DAYS,
        extra_files=[Config.OUTPUT_FILE],
    )


def start_scheduler():
    global _scheduler
    if _scheduler is not None:
        return _scheduler

    # Run once immediately at startup so stale data doesn't wait for the
    # first interval to elapse.
    _run_cleanup()

    _scheduler = BackgroundScheduler(daemon=True)
    _scheduler.add_job(
        _run_cleanup,
        "interval",
        hours=Config.CLEANUP_INTERVAL_HOURS,
        id="data_retention_cleanup",
        replace_existing=True,
    )
    _scheduler.start()
    logger.info(
        "Data retention scheduler started (retention=%s days, check every %s hours).",
        Config.DATA_RETENTION_DAYS, Config.CLEANUP_INTERVAL_HOURS,
    )
    return _scheduler
