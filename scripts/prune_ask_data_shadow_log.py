"""Prune old Ask Data v2/v3 shadow comparison log entries (read-only diagnostics; see
services/ask_data/shadow.py and migrations/create_ask_data_shadow_log.sql).

Dry-run by default: prints how many entries are older than --older-than-days and would be deleted, without
deleting anything. Pass --apply to actually delete them. Only ever touches the ask_data_shadow_log table/
local file -- never invoices, payments, projects, or any other production data.

Usage:
    python scripts/prune_ask_data_shadow_log.py                        # dry run, default 30 days
    python scripts/prune_ask_data_shadow_log.py --older-than-days 14   # dry run, custom cutoff
    python scripts/prune_ask_data_shadow_log.py --older-than-days 14 --apply
"""
from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def _cutoff(older_than_days: int) -> datetime.datetime:
    return datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=older_than_days)


def _parse_timestamp(value) -> datetime.datetime | None:
    if not value:
        return None
    try:
        text = str(value).replace("Z", "+00:00")
        dt = datetime.datetime.fromisoformat(text)
        return dt if dt.tzinfo else dt.replace(tzinfo=datetime.timezone.utc)
    except Exception:
        return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--older-than-days", type=int, default=30, help="delete entries older than this many days (default 30)")
    parser.add_argument("--apply", action="store_true", help="actually delete; omit for a dry run")
    args = parser.parse_args()

    from config.settings import load_ask_data_shadow_log

    cutoff = _cutoff(args.older_than_days)
    entries = load_ask_data_shadow_log()
    stale = [e for e in entries if (_parse_timestamp(e.get("created_at") or e.get("timestamp")) or cutoff) < cutoff]

    print(f"Ask Data shadow log: {len(entries)} total entries, {len(stale)} older than {args.older_than_days} day(s) "
          f"(cutoff {cutoff.isoformat()}).")

    if not stale:
        print("Nothing to prune.")
        return 0

    if not args.apply:
        print(f"Dry run only -- would delete {len(stale)} entries. Re-run with --apply to actually delete them.")
        return 0

    deleted = 0
    try:
        from services.supabase_service import _get_client, ASK_DATA_SHADOW_LOG_TABLE

        client = _get_client()
        for entry in stale:
            row_id = entry.get("id")
            if row_id is None:
                continue
            client.table(ASK_DATA_SHADOW_LOG_TABLE).delete().eq("id", row_id).execute()
            deleted += 1
    except Exception as exc:
        print(f"Supabase prune skipped/failed ({exc}); local-only entries (if any) are kept as-is -- "
              f"the local file is already size-bounded (see ASK_DATA_SHADOW_LOG_LOCAL_MAX_ENTRIES).")

    print(f"Deleted {deleted} entries from Supabase.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
