import os
import time
from openpyxl import load_workbook, Workbook


def append_to_xlsx(xlsx_path: str, row: list, headers: list | None = None) -> None:
    """Append a row to an xlsx workbook with file-level locking for parallel-safe writes.

    Raises TimeoutError if the lock cannot be acquired within ~10 seconds. Writing
    without the lock would risk corrupting the xlsx, so we fail loudly instead.
    """
    path = os.path.abspath(xlsx_path)
    lock_path = path + ".lock"
    acquired = False
    for _ in range(20):
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            acquired = True
            break
        except FileExistsError:
            time.sleep(0.5)
    if not acquired:
        raise TimeoutError(f"Could not acquire lock on {lock_path} after 10 seconds")
    try:
        if os.path.exists(path):
            wb = load_workbook(path)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            if headers:
                ws.append(headers)
        ws.append(row)
        wb.save(path)
    finally:
        try:
            os.unlink(lock_path)
        except Exception:
            pass


def make_xlsx_appender(path: str, headers: list):
    """Return a parallel-safe row-appender bound to a specific xlsx file and header schema."""
    def appender(row: list) -> None:
        append_to_xlsx(path, row, headers)
    return appender
