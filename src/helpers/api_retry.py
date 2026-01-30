"""
Centralized retry utility for Google Sheets API calls.

Google Sheets API has a quota of ~60 write requests per minute per user.
This module provides retry-with-backoff for any API call that might hit 429.
"""

import time
import gspread


def is_rate_limit_error(e):
    """
    Returns True if the exception is a Google Sheets API rate limit (HTTP 429).
    """
    return "429" in str(e)


def retry_on_429(fn, *, max_attempts=8, base_sleep_seconds=15, max_sleep_seconds=90):
    """
    Run a function, retrying with exponential backoff on HTTP 429.
    
    Args:
        fn: Callable to execute (should make a Sheets API call)
        max_attempts: Maximum number of retry attempts (default 8)
        base_sleep_seconds: Initial sleep duration on first 429 (default 15s)
        max_sleep_seconds: Maximum sleep duration (default 90s, covers the 60s quota window)
    
    Returns:
        The return value of fn() if successful
        
    Raises:
        The last exception if all retries are exhausted
    """
    last_exc = None
    for attempt in range(max_attempts):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            last_exc = e
            if not is_rate_limit_error(e):
                raise
            # Calculate sleep with exponential backoff, capped at max
            sleep_s = min(max_sleep_seconds, base_sleep_seconds * (1.5 ** attempt))
            time.sleep(sleep_s)
    # Exhausted retries
    raise last_exc


def batch_update_with_retry(spreadsheet, requests, *, chunk_size=200):
    """
    Execute a Sheets API batchUpdate with automatic 429 retry.
    
    Args:
        spreadsheet: gspread.Spreadsheet object
        requests: List of request dictionaries for batch_update
        chunk_size: Max requests per batch call (default 200)
    """
    if not requests:
        return
    for i in range(0, len(requests), chunk_size):
        batch = requests[i:i + chunk_size]
        retry_on_429(lambda b=batch: spreadsheet.batch_update({"requests": b}))
