#!/usr/bin/env python3
"""
Download race results from Weibu (微步, https://sport.luojiweiye.com) and summarise them.

Default mode (--event-id N): for each race (distance) in the event, fetch every runner's result
and, unless --no-details, their checkpoint times. Prints per-race statistics (status counts,
finish rate, finish-time spread and histogram) and writes {event name}.xlsx with sheets:
- Summary: per race - entrants, DNS, starters, timed out, withdrew, DSQ, finishers,
  finish rate and finish-time quartiles
- By gender: the same per race and sex
- one sheet per race: every runner, with cumulative time and rank at each checkpoint
- Checkpoints: one row per runner per checkpoint (only when details are fetched)

Other modes:
- --all-events: list every event id found so far (cache/events.json) with its name
- --find-player NAME: search downloaded results for a runner's name across events
- --attachments: download the event's attachments (e.g. GPX/KML course files)

API responses are cached under cache/ (see --refresh and --clear-cache).
"""

# Weibu endpoints (base https://sport.luojiweiye.com):
#
# Get race name for a race:
#   /api/sport/activity/for-result?activity_id={activity_id}
#   data: {id, name, team_grade}
#
# Get race distances for a race:
#   /api/sport/packages/for-result?activity_id={activity_id}
#   data: [{name, value (= package_id), qiang_at (gun start), ready_qiang_at}]
#   ready_qiang_at '1970-01-01 00:00:00' means not set
#
# Get results for a race + distance (page_size max 20):
#   /api/web/players?activity_id={activity_id}&package_id={package_id}&page=1&page_size=20
#   data: {total, per_page, current_page, last_page, data: [one row per runner:
#          id (= player_id), name, sex, no (bib), status, status_name, order,
#          time_used, result_qiang (gun time), result_jing (chip time), cp_name (last CP), ...]}
#
# Get results for a particular runner:
#   /api/sport/admin/result?player_id={player_id}
#   data: {player: {...}, results: [one entry per CP: cp_name, cp_mileage,
#          when (wall clock), time_used (leg time), closing_time (mins), status_label, ...]}
#
# Get attachments for a race (web page):
#   /web/weibu/#/pagesA/pages/event/attachment?game_id={activity_id}
#
# Get attachment lists for a race (fetch_attachments() uses and merges both):
#   /api/file/getFileListH5?game_id={activity_id}
#   /api/web/track_file/web_list?game_id={activity_id}
#   data: [{name, url, type, ...}]; track_file entries also have track_url, track_url_gd, json_url
#
# Personal result page (HTML):
#   /web/h5/new_gerenchengji.html?id={player_id}

# Standard library
import argparse
import json
import math
import os
import pathlib
import random
import re
import shutil
import statistics
import sys
import time
import urllib.parse
from collections import Counter

# Third party
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from tabulate import tabulate
from wcwidth import wcswidth

BASE = "https://sport.luojiweiye.com"
PAGE_SIZE = 20  # server rejects >20 with generic 51000 error
MAX_RANGE = 500 # refuse accidental huge ranges
TIMEOUT = 15
PROGRESS_BAR_WIDTH = 24
SLEEP_BETWEEN_PAGES = 0.3
TABLEFMT = "github"
FETCH_CHECKPOINTS = True
CACHE_DIR        = pathlib.Path("cache")
RACE_CACHE_DIR   = CACHE_DIR / "races"    # /api/web/players, one file per race
PLAYER_CACHE_DIR = CACHE_DIR / "players"  # /api/sport/admin/result, one file per player
ATTACHMENTS_CACHE_DIR = CACHE_DIR / "attachments"
EVENTS_REGISTRY = CACHE_DIR / "events.json"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; results-scraper/1.0)",
    "Accept": "application/json",
}

class EventNotFound(RuntimeError):
    """API code 51200: no event with that id (未找到该比赛)."""

def configure(fetch_details, cache_dir):
    global FETCH_CHECKPOINTS, CACHE_DIR, RACE_CACHE_DIR, PLAYER_CACHE_DIR, ATTACHMENTS_CACHE_DIR, EVENTS_REGISTRY
    FETCH_CHECKPOINTS = fetch_details
    CACHE_DIR = pathlib.Path(cache_dir)
    RACE_CACHE_DIR = CACHE_DIR / "races"
    PLAYER_CACHE_DIR = CACHE_DIR / "players"
    ATTACHMENTS_CACHE_DIR = CACHE_DIR / "attachments"
    EVENTS_REGISTRY = CACHE_DIR / "events.json"


def render_table(rows, headers, colalign=None, maxwidths=None):
    """Render a console table with CJK-aware alignment.

    `maxwidths` truncates long cells with an ellipsis so one enormous
    race/event name can't blow the table out. Requires wcwidth to be
    effective; without it tabulate falls back to code-point widths.
    """
    return tabulate(
        rows,
        headers=headers,
        tablefmt=TABLEFMT,
        colalign=colalign,
        maxcolwidths=maxwidths,
        disable_numparse=True,   # keep "88.8%" and times as-is
    )


def _event_id_spec(s):
    """Parse '6394' or '6000-6500' (inclusive) into a list of ints."""
    try:
        if "-" in s:
            lo, _, hi = s.partition("-")
            lo, hi = int(lo), int(hi)
            if lo > hi:
                raise ValueError(f"empty range {lo}-{hi}")
            if hi - lo + 1 > MAX_RANGE:
                raise ValueError(f"range {lo}-{hi} has {hi - lo + 1 } ids (max {MAX_RANGE})")
            return list(range(lo, hi + 1))
        return [int(s)]
    except ValueError as e:
        raise argparse.ArgumentTypeError(
            f"invalid event id spec {s!r} (use N or LO-HI): {e}"
        )
    
def clear_cache():
    """Delete cached API responses. Safe to call when the cache doesn't exist.

    Keeps the event-id registry (events.json): it lists the event ids found so far,
    which could otherwise only be rebuilt by scanning id ranges again.
    """
    if not CACHE_DIR.exists():
        print(f"Nothing to clear ({CACHE_DIR}/ does not exist)")
        return
    for child in CACHE_DIR.iterdir():
        if child == EVENTS_REGISTRY:
            continue
        if child.is_dir() and not child.is_symlink():
            shutil.rmtree(child)
        else:
            child.unlink()
    print(f"Cleared {CACHE_DIR}/ (kept {EVENTS_REGISTRY.name})")

        
def _event_cache_path(activity_id):
    """Path to the cached event-name payload for the current activity."""
    return CACHE_DIR / f"event_{activity_id}.json"


def _races_cache_path(activity_id):
    """Path to the cached race-list payload for the current activity."""
    return CACHE_DIR / f"races_{activity_id}.json"


def _race_cache_path(activity_id, package_id, sex=""):
    """Path to the cached list payload for one race (and gender filter)."""
    return RACE_CACHE_DIR / f"{activity_id}_{package_id}_{sex or 'all'}.json"


def _player_cache_path(player_id):
    """Path to the cached detail payload for one athlete."""
    return PLAYER_CACHE_DIR / f"{player_id}.json"


def _load_known_events():
    """Sorted list of event ids seen so far (empty if no registry yet)."""
    if not EVENTS_REGISTRY.exists():
        return []
    try:
        return sorted(set(json.loads(EVENTS_REGISTRY.read_text(encoding="utf-8"))["event_ids"]))
    except (json.JSONDecodeError, KeyError, OSError) as e:
        print(f"  warning: bad event registry ({e}) — starting fresh")
        return []


def _save_known_events(ids):
    """Persist the event-id registry."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    EVENTS_REGISTRY.write_text(
        json.dumps({"event_ids": sorted(set(ids))}, ensure_ascii=False),
        encoding="utf-8",
    )


def _race_meta_lookup(activity_id):
    """{package_id: {name, qiang_at, ...}} from races_<id>.json, or {}."""
    path = _races_cache_path(activity_id)
    if not path.exists():
        return {}
    try:
        races = json.loads(path.read_text(encoding="utf-8"))["races"]
    except (json.JSONDecodeError, KeyError, OSError):
        return {}
    return {r["package_id"]: r for r in races}


def _url_extension(url):
    """Lowercased file extension from a URL's path, ignoring query/fragment.

    Returns '' when the path has no suffix (e.g. '.../download?id=3').
    """
    suffix = pathlib.PurePosixPath(urllib.parse.urlparse(url).path).suffix
    return suffix.lstrip(".").lower()


def _wait_before_retry(attempt, retries, backoff, url, reason):
    """Exponential backoff with jitter, plus a visible console note."""
    delay = backoff * (2 ** (attempt - 1)) + random.uniform(0, 0.5)
    sys.stdout.write(
        f"\r  retry {attempt}/{retries - 1} in {delay:.1f}s "
        f"({reason}) — {url}\n"
    )
    sys.stdout.flush()
    time.sleep(delay)


def _line_width(s):
    """Display cells occupied by `s`; falls back to len() for control chars."""
    w = wcswidth(s)
    return w if w >= 0 else len(s)


def _progress(race_name, page, last_page, got, total):
    """Draw a self-rewriting progress line; return its length so the caller
    can blank it with spaces."""
    total = total or 0
    frac = (got / total) if total else 0
    width = PROGRESS_BAR_WIDTH
    filled = int(frac * width)
    bar = "█" * filled + "░" * (width - filled)
    pct = f"{frac*100:5.1f}%"
    line = f"  {race_name:<10} [{bar}] {pct}  page {page}/{last_page}  rows {got}/{total}"
    sys.stdout.write("\r" + line)
    sys.stdout.flush()
    return _line_width(line)


def _detail_progress(race_name, i, total):
    frac = i / total if total else 0
    width = PROGRESS_BAR_WIDTH
    filled = int(frac * width)
    bar = "█" * filled + "░" * (width - filled)
    line = f"  {race_name:<10} details [{bar}] {i}/{total}"
    sys.stdout.write("\r" + line)
    sys.stdout.flush()
    return _line_width(line)


def _event_progress(i, total, event_id, source, last_len):
    """One self-rewriting line: 'event 123/500  id 6394 [net]'."""
    line = f"  events {i}/{total}  id {event_id} [{source}]"
    line_w = _line_width(line)
    # pad to the previous line length so a shorter line doesn't leave residue
    pad = max(0, last_len - line_w)
    sys.stdout.write("\r" + line + " " * pad)
    sys.stdout.flush()
    return line_w


# TODO: HTTP 429 (rate limited) and non-JSON replies (e.g. an HTML error page) aren't retried and raise
#       HTTPError / JSONDecodeError rather than RuntimeError, so they abort an --event-id range scan even
#       with skip_errors=True. Retry 429 with backoff like 5xx, and raise RuntimeError for bad JSON.
def get_json(path, params=None, retries=6, backoff=1.5):
    url = BASE + path
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
        except (requests.Timeout, requests.ConnectionError) as e:
            if attempt == retries:
                raise RuntimeError(
                    f"Giving up on {url} after {retries} attempts: {e}"
                ) from e
            _wait_before_retry(attempt, retries, backoff, url, type(e).__name__)
            continue

        if 500 <= r.status_code < 600:
            if attempt < retries:
                _wait_before_retry(attempt, retries, backoff, url, f"HTTP {r.status_code}")
                continue

        r.raise_for_status()
        payload = r.json()
        code = payload.get("msg", {}).get("code")
        if code == 51105:
            raise RuntimeError("Results not yet published: " + payload["msg"].get("info", ""))
        if code == 51200:
            raise EventNotFound(payload["msg"].get("info", ""))
        if code != 0:
            raise RuntimeError(f"API error on {url}: {payload.get('msg')}")
        return payload["data"]


def _download_file(url, target, retries=4, backoff=1.5):
    """Stream `url` to `target` atomically (temp file + os.replace).

    On any failure the temp file is removed, so a later run re-downloads
    cleanly instead of trusting a truncated file.
    """
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(target.name + ".part")
    for attempt in range(1, retries + 1):
        try:
            with requests.get(url, headers=HEADERS, timeout=TIMEOUT, stream=True) as r:
                r.raise_for_status()
                with tmp.open("wb") as fh:
                    for chunk in r.iter_content(chunk_size=1 << 16):
                        fh.write(chunk)
            os.replace(tmp, target)   # atomic same-dir rename
            return target
        except (requests.Timeout, requests.ConnectionError,
                requests.HTTPError, OSError) as e:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
            if attempt == retries:
                raise RuntimeError(
                    f"Giving up downloading {url} after {retries} attempts: {e}"
                ) from e
            _wait_before_retry(attempt, retries, backoff, url, type(e).__name__)


def enumerate_events(event_ids=None, use_cache=True, skip_errors=False):
    """Resolve a list of event ids to [(event_id, name)] pairs.

    Reads the registry when `event_ids` is None. Prints an in-place
    progress line tagged [cache] or [net] per id, so a 500-id range is
    visibly working and you can see how much is being served from cache.
    """ 
    event_ids = sorted(set(event_ids if event_ids is not None else _load_known_events()))
    if not event_ids:
        raise RuntimeError(
            "No event ids known. Seed with --event-id or run a single-event "
            "scrape first (the id is recorded automatically)."
        )
    out = []
    total = len(event_ids)
    last_len = 0
    # Record newly found ids in the registry once at the end (also on Ctrl+C or an error),
    # rather than rewriting the whole registry file for every id
    known = set(_load_known_events())
    new_ids = set()
    try:
        for i, eid in enumerate(event_ids, 1):
            try:
                name, source = _get_event_name(eid, use_cache=use_cache)
            except EventNotFound:
                # A dead id is not an error — just skip it, and don't record it.
                last_len = _event_progress(i, total, eid, "missing", last_len)
                continue
            except RuntimeError as e:
                if not skip_errors:
                    sys.stdout.write("\r" + " " * last_len + "\r")
                    sys.stdout.flush()
                    raise
                print(f"  warning: event {eid} skipped: {e}")
                last_len = _event_progress(i, total, eid, "skip", last_len)
                continue
            out.append((eid, name))
            if eid not in known:
                new_ids.add(eid)
            last_len = _event_progress(i, total, eid, source, last_len)
            if source == "net" and i < total:
                time.sleep(SLEEP_BETWEEN_PAGES)
    finally:
        if new_ids:
            _save_known_events(known | new_ids)
    sys.stdout.write("\r" + " " * last_len + "\r")
    sys.stdout.flush()
    return out


def _get_event_name(activity_id, use_cache=True):
    """Event name for activity_id, cached to CACHE_DIR/event_{id}.json."""
    path = _event_cache_path(activity_id)
    if use_cache and path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))["name"], "cache"
        except (json.JSONDecodeError, KeyError, OSError) as e:
            print(f"  warning: bad event cache ({e}) — refetching")
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass

    data = get_json("/api/sport/activity/for-result", {"activity_id": activity_id})
    name = data.get("name", f"activity {activity_id}")
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"name": name}, ensure_ascii=False), encoding="utf-8")
    return name, "net"


def get_event_name(activity_id, use_cache=True):
    """Event name for activity_id (cached). Returns just the name."""
    return _get_event_name(activity_id, use_cache=use_cache)[0]

def get_races(activity_id, use_cache=True):
    """Race list for activity_id, cached to CACHE_DIR/races_{id}.json."""
    path = _races_cache_path(activity_id)
    if use_cache and path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))["races"]
        except (json.JSONDecodeError, KeyError, OSError) as e:
            print(f"  warning: bad race-list cache ({e}) — refetching")
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass

    data = get_json("/api/sport/packages/for-result", {"activity_id": activity_id})
    out = []
    for item in data:
        qiang_at = item.get("qiang_at") or ""
        ready = item.get("ready_qiang_at") or ""
        # '1970-01-01 00:00:00' is the API's null placeholder
        if ready.startswith("1970-01-01"):
            ready = ""
        out.append({
            "package_id": item.get("value"),
            "name": item.get("name") or item.get("package_name") or str(item.get("value")),
            "qiang_at": qiang_at,
            "ready_qiang_at": ready,
        })
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"races": out}, ensure_ascii=False), encoding="utf-8")
    return out


def find_player(name, case_sensitive=False, exact=False):
    """Search every cached race roster for athletes whose name matches.

    Reads RACE_CACHE_DIR/*_all.json (the /api/web/players cache) — no network.
    Filenames encode activity_id and package_id; race names come from
    races_<activity_id>.json when present, else fall back to the ids.

    Args:
        name: substring (or full name when exact=True)
        case_sensitive: match case-sensitively
        exact: match the whole name instead of a substring

    Returns a list of dicts sorted by (event, race, name):
        {event_id, event, package_id, race, player_id, name, bib, sex,
         status, status_name, time_used, order}
    """
    needle = name if case_sensitive else name.lower()
    hits = []
    meta_cache = {}   # activity_id -> {package_id: race_record}

    for path in sorted(RACE_CACHE_DIR.glob("*_all.json")):
        # filename: <activity_id>_<package_id>_all.json
        stem = path.stem                      # "<aid>_<pid>_all"
        try:
            aid_s, pid_s, _ = stem.rsplit("_", 2)
            aid, pid = int(aid_s), int(pid_s)
        except ValueError:
            continue                          # not one of ours

        try:
            rows = json.loads(path.read_text(encoding="utf-8")).get("rows") or []
        except (json.JSONDecodeError, OSError) as e:
            print(f"  warning: bad roster cache {path.name}: {e}")
            continue

        matched = []
        for row in rows:
            pname = (row.get("name") or "").strip()
            hay = pname if case_sensitive else pname.lower()
            ok = (hay == needle) if exact else (needle in hay)
            if ok:
                matched.append(row)
        if not matched:
            continue

        # Lazily resolve event + race names (one read per event).
        if aid not in meta_cache:
            meta_cache[aid] = _race_meta_lookup(aid)
        try:
            event_name = _get_event_name(aid, use_cache=True)[0]
        except RuntimeError:
            event_name = f"event {aid}"
        race_meta = meta_cache[aid].get(pid) or {}
        race_name = race_meta.get("name") or f"race {pid}"

        for row in matched:
            hits.append({
                "event_id": aid,
                "event": event_name,
                "package_id": pid,
                "race": race_name,
                "start": race_meta.get("qiang_at") or "",
                "player_id": row.get("id"),
                "name": row.get("name"),
                "bib": row.get("no"),
                "sex": _sex_label(row.get("sex")),
                "status": classify_status(row),
                "status_name": row.get("status_name"),
                "time_used": row.get("time_used"),
                "order": row.get("order"),
            })

    hits.sort(key=lambda h: (h["event"], h["race"], h["name"] or ""))
    hits.sort(key=lambda h: (
        h["start"] == "",          # False (0) sorts before True (1) → dated first
        h["start"],                # then chronological
    ))
    return hits


# TODO: cached results never expire - data fetched during or just after a race (runners still 'in race') is
#       served from cache on every later run unless --refresh is given. Same for player details and race lists.
#       Consider refetching when the cache predates the race finishing, or when any runner is still in race.
def fetch_all_players(activity_id, package_id, race_name, sex="", use_cache=True):
    """Paginate /api/web/players, printing an in-place progress line.

    Results are cached to RACE_CACHE_DIR keyed by
    (activity_id, package_id, sex); pass use_cache=False to force a refetch.
    """
    path = _race_cache_path(activity_id, package_id, sex)
    if use_cache and path.exists():
        try:
            cached = json.loads(path.read_text(encoding="utf-8"))
            print(f"  {race_name:<10} loaded {len(cached['rows'])} rows from cache")
            return cached["rows"], cached["api_total"]
        except (json.JSONDecodeError, KeyError, OSError) as e:
            print(f"\n  warning: bad list cache for {race_name}: {e} — refetching")
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass

    all_rows = []
    page = 1
    api_total = None
    last_page = None
    last_line_len = 0
    try:
        while True:
            params = {
                "activity_id": activity_id,
                "package_id": package_id,
                "page": page,
                "page_size": PAGE_SIZE,
                "sex": sex,
                "real_name": "",
                "age_range": "",
            }
            data = get_json("/api/web/players", params)
            rows = data.get("data") or []
            all_rows.extend(rows)
            api_total = data.get("total", api_total)
            last_page = data.get("last_page", last_page or 1)
            current_page = data.get("current_page", page)
            last_line_len = _progress(race_name, current_page, last_page, len(all_rows), api_total)
            if current_page >= last_page or not rows:
                break
            page += 1
            time.sleep(SLEEP_BETWEEN_PAGES)
    finally:
        # Clear the progress line even if we exit early due to an exception.
        sys.stdout.write("\r" + " " * last_line_len + "\r")
        sys.stdout.flush()

    RACE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"rows": all_rows, "api_total": api_total}, ensure_ascii=False),
        encoding="utf-8",
    )
    return all_rows, api_total


def fetch_player_detail(player_id, use_cache=True):
    """Fetch (and cache) the per-player detail payload.
    Parameter `use_cache` can be set to False to force a fresh fetch."""
    path = _player_cache_path(player_id)
    if use_cache and path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    data = get_json("/api/sport/admin/result", {"player_id": player_id})
    PLAYER_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


def fetch_all_details(rows, race_name, use_cache=True):
    """Fetch detail for every row, with a progress line."""
    details = {}
    total = len(rows)
    last_len = 0
    for i, row in enumerate(rows, 1):
        pid = row.get("id")
        if pid is None:
            continue
        cached = _player_cache_path(pid).exists()
        try:
            details[pid] = fetch_player_detail(pid, use_cache=use_cache)
        except Exception as e:
            print(f"\n  warning: detail fetch failed for player {pid}: {e}")
            # A corrupt cache file would otherwise fail on every re-run.
            try:
                _player_cache_path(pid).unlink(missing_ok=True)
            except OSError:
                pass
            cached = False
        last_len = _detail_progress(race_name, i, total)
        # only throttle when we actually hit network
        if not cached and i < total:
            time.sleep(SLEEP_BETWEEN_PAGES)
    sys.stdout.write("\r" + " " * last_len + "\r")
    sys.stdout.flush()
    return details


# TODO: ignores --refresh (always uses the cache when present), and if both endpoints fail it caches an empty
#       list, so later runs report no attachments even after the site recovers. Don't cache an empty result
#       caused by failures, and honour use_cache.
def fetch_attachments(game_id):
    """Union of both attachment endpoints. Each entry is tagged with
    `_source` ('h5' or 'track'). Returns [] if both endpoints fail."""
    ATTACHMENT_ENDPOINTS = [
        ("h5",    "/api/file/getFileListH5",       "game_id"),
        ("track", "/api/web/track_file/web_list",  "game_id"),
    ]
    path = ATTACHMENTS_CACHE_DIR / f"{game_id}.json"
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass

    merged = []
    seen_urls = set()
    for source, endpoint, param in ATTACHMENT_ENDPOINTS:
        try:
            data = get_json(endpoint, {param: game_id})
        except Exception as e:
            print(f"  note: {source} attachment list failed for {game_id}: {e}")
            continue
        for e in data or []:
            e = dict(e)                       # don't mutate the cached copy
            e["_source"] = source
            url = e.get("url") or ""
            if not url:
                continue
            if url in seen_urls:
                continue
            seen_urls.add(url)
            # Fill in type from URL if the payload didn't provide it.
            if not e.get("type"):
                e["type"] = _url_extension(url) or "bin"
            merged.append(e)

    ATTACHMENTS_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(merged, ensure_ascii=False), encoding="utf-8")
    return merged


def download_attachments(game_id, out_dir=None, types=None):
    entries = fetch_attachments(game_id)
    dest = pathlib.Path(out_dir) if out_dir else (CACHE_DIR / "attachments" / str(game_id))
    dest.mkdir(parents=True, exist_ok=True)
    if types:
        types = {t.lower().lstrip(".") for t in types}

    written = []
    for e in entries:
        ext = (e.get("type") or "").lower()
        # Each entry may carry several fetchable URLs; yield them all.
        urls = [("", e.get("url") or "")]
        if e.get("_source") == "track":
            urls += [
                ("_3d",   e.get("track_url") or ""),
                ("_3d_gd", e.get("track_url_gd") or ""),
                ("_json", e.get("json_url") or ""),
            ]
        base = (e.get("name") or pathlib.Path(urls[0][1]).stem or "file").strip()

        for suffix, url in urls:
            if not url:
                continue
            if suffix and ext == "":
                file_ext = _url_extension(url) or "bin"
            elif suffix:
                file_ext = _url_extension(url) or ext
            else:
                file_ext = ext
            if types and file_ext.lower() not in types and not suffix:
                # --attachment-types filters the *primary* file only;
                # viewer URLs are downloaded alongside their GPX
                continue
            filename = safe_filename(f"{base}{suffix}.{file_ext}")
            target = dest / filename
            if target.exists():
                written.append(target)
                continue
            print(f"  downloading {filename}")
            written.append(_download_file(url, target))
    return written


# ----------------------------------------------------------------------------
# Parsing / classification
# ----------------------------------------------------------------------------
# todo: Handle fractional seconds in time_used (e.g. 01:23:45.678)
# int("45.678") will throw ValueError and we'll return None
def parse_time_to_seconds(t):
    if not t or not isinstance(t, str):
        return None
    parts = t.strip().split(":")
    try:
        parts = [int(p) for p in parts]
    except ValueError:
        return None
    if len(parts) == 3:
        h, m, s = parts
    elif len(parts) == 2:
        h, m, s = 0, parts[0], parts[1]
    else:
        return None
    return h * 3600 + m * 60 + s


def classify_status(row):
    """
    Map one API row to a normalised outcome category.

    Returns one of: FINISH, TIMEOUT, WITHDREW, DSQ, DNS, OTHER.

    Treat Weibu's `status_name` as the authoritative outcome label, so it
    takes precedence over the numeric `status` code.
    Precedence:
      1. Exact match on one of the five known names (DSQ, 被关门, 已退赛, 待参赛, 已完赛).
      2. Fallback: a positive integer `order` means the athlete was ranked,
         i.e. classified as a finisher even if the label is unexpected.
      3. Anything else -> OTHER (surfaced in the console table so new labels
         are visible rather than silently dropped).

    The Vue filter in the original web app is roughly:
    
    if(val.order !== '') {
        if(val.status>251 && val.status<=300 || val.status ==200){
            return '0' + val.order   // show a rank → treated as a finisher
        } else {
            return val.status_name   // DSQ / 被关门 / 已退赛 / 待参赛
        }
    }

    this is not replicated below
    """

    name = (row.get("status_name") or "").strip()

    if name == "DSQ":     return "DSQ"
    if name == "被关门":   return "TIMEOUT"
    if name == "已退赛":   return "WITHDREW"
    if name == "待参赛":   return "DNS"
    if name == "DNS":     return "DNS"
    if name == "已完赛":   return "FINISH"
    if name == "正在比赛": return "IN RACE"
    if name == "比赛中":   return "IN RACE"

    # Fallback only if status_name is missing/unknown:
    # a positive rank means the athlete was classified.
    order = row.get("order")
    try:
        if order is not None and int(order) > 0:
            return "FINISH"
    except (TypeError, ValueError):
        pass

    return "OTHER"


def _mileage_key(r):
    """Sort key: numeric cp_mileage, unparseable/missing sorts last."""
    try:
        return float(r.get("cp_mileage"))
    except (TypeError, ValueError):
        return math.inf


def parse_checkpoints(detail, status=None):
    results = (detail or {}).get("results") or []
    ordered = sorted(results, key=_mileage_key)

    start_cp = next((r for r in ordered if _mileage_key(r) == 0), None)
    start_s = _wall_seconds(start_cp.get("when")) if start_cp else None

    # statuses for which a missing lap is genuinely "not reached"
    # and for which we must stop at the first gap
    stop_at_gap = status in ("TIMEOUT", "WITHDREW", "DSQ")
    keep_all    = status in ("FINISH", "IN RACE")
    # DNS: nothing reached (but the existing per-CP rule already yields False
    # for every non-START CP when time_used is empty)

    out = []
    day_offset = 0
    prev_wall = start_s
    gap_seen = False
    for i, r in enumerate(ordered, 1):
        name = r.get("cp_name") or ""
        is_start = _mileage_key(r) == 0
        lap_used = r.get("time_used")
        wall = _wall_seconds(r.get("when"))

        if wall is not None and prev_wall is not None and wall + day_offset < prev_wall:
            day_offset += 86400
        abs_wall = (wall + day_offset) if wall is not None else None
        cumulative = None if (is_start or abs_wall is None or start_s is None) \
                          else abs_wall - start_s

        has_lap = bool(lap_used) and lap_used not in ("--", "00:00:00")

        if is_start:
            reached = True
        elif gap_seen:
            reached = False
        elif stop_at_gap:
            # non-finisher: first empty lap closes the course for them
            if has_lap:
                reached = True
            else:
                reached = False
                gap_seen = True
        elif keep_all:
            reached = has_lap
        else:  # DNS or unknown status — be conservative
            reached = has_lap

        out.append({
            "cp_order": i,
            "cp_name": name,
            "cp_mileage": r.get("cp_mileage"),
            "when": r.get("when") if reached else "",
            "lap_used": lap_used if reached else "",
            "time_used": fmt_secs(cumulative) if (reached and cumulative is not None) else "",
            "time_seconds": cumulative if reached else None,
            "abs_wall_seconds": abs_wall if reached else None,
            "speed_allocation": r.get("speed_allocation") if reached else "",
            "status_label": r.get("status_label"),
            "closing_time": r.get("closing_time"),
            "is_warn": r.get("is_warn"),
            "reached": reached,
            "is_start": is_start,
        })
        if reached and wall is not None:
            prev_wall = abs_wall
    return out


def fmt_secs(s):
    if s is None:
        return ""
    s = int(round(s))
    return f"{s//3600:02d}:{(s%3600)//60:02d}:{s%60:02d}"


def percentile(sorted_vals, p):
    """Linear-interpolated percentile. Caller must pass a sorted list."""
    if not sorted_vals:
        return None
    k = (len(sorted_vals) - 1) * (p / 100.0)
    f = int(k)
    c = min(f + 1, len(sorted_vals) - 1)
    if f == c:
        return sorted_vals[f]
    return sorted_vals[f] + (sorted_vals[c] - sorted_vals[f]) * (k - f)


def _sex_label(v):
    return "M" if v == 1 else "F" if v == 2 else "?"


def safe_filename(name, max_len=120):
    """Make `name` safe as a filename on Windows, macOS, and Linux."""
    # Replace anything not alphanumeric / space / dash / underscore / dot.
    # \w with re.UNICODE keeps CJK characters (they're legal filenames).
    cleaned = re.sub(r"[^\w\-. ]", "_", name, flags=re.UNICODE)
    # Collapse runs of underscores/spaces, strip trailing dots/spaces (Windows).
    cleaned = re.sub(r"[_ ]{2,}", "_", cleaned).strip(" .")
    # Windows reserved device names.
    if cleaned.upper().split(".")[0] in {
        "CON", "PRN", "AUX", "NUL",
        *{f"COM{i}" for i in range(1, 10)},
        *{f"LPT{i}" for i in range(1, 10)},
    }:
        cleaned = "_" + cleaned
    if not cleaned:
        cleaned = "results"
    return cleaned[:max_len]


def excel_sheet_title(name):
    """Make a race name valid as an Excel sheet title."""
    title = name.replace('[', '(').replace(']', ')')   # keep the meaning of "[100K]"
    title = re.sub(r'[*?:/\\]', '-', title)             # other characters Excel forbids
    title = title.strip(" '")[:31].rstrip(" '")         # no edge apostrophes, 31-char limit
    return title or 'Race'                              # never empty


def _to_float_or_blank(v):
    """Coerce a numeric-looking string to float; return '' for empty/garbage."""
    if v is None or v == "":
        return ""
    try:
        return float(v)
    except (TypeError, ValueError):
        return ""
    

def rank_checkpoints(enriched):
    by_cp = {}
    for row in enriched:
        for cp in row.get("checkpoints", []):
            if cp["is_start"] or not cp["reached"]:
                continue
            if cp.get("time_seconds") is None:
                continue
            by_cp.setdefault(cp["cp_name"], []).append(cp)
    for cps in by_cp.values():
        _competition_rank(cps, key=lambda c: c["time_seconds"],    out="elapsed_rank")
        _competition_rank(cps, key=lambda c: c["abs_wall_seconds"], out="rank")


def _wall_seconds(when):
    """'HH:MM:SS.mmm' -> float seconds since midnight, or None."""
    if not when:
        return None
    try:
        h, m, s = when.split(":")
        return int(h) * 3600 + int(m) * 60 + float(s)
    except (ValueError, AttributeError):
        return None


def _competition_rank(items, key, out):
    """Standard competition ranking (1224) by `key(item)`, stored in item[out]."""
    ordered = sorted(items, key=key)
    prev_val = None
    prev_rank = 0
    for i, item in enumerate(ordered, 1):
        val = key(item)
        if prev_val is not None and val == prev_val:
            item[out] = prev_rank
        else:
            item[out] = i
            prev_val = val
            prev_rank = i


def summarize_race(event, name, rows, api_total, details=None):
    statuses = Counter()
    times_sec = []
    enriched = []
    details = details or {}

    # per-gender accumulators: total, finished, dns, in-race, and finish times
    sex_total = Counter()
    sex_fin = Counter()
    sex_dns = Counter()
    sex_inrace = Counter()
    sex_times = {}

    for row in rows:
        cat = classify_status(row)
        statuses[cat] += 1
        t = parse_time_to_seconds(row.get("time_used")) if cat == "FINISH" else None
        if t is not None:
            times_sec.append(t)

        sx = _sex_label(row.get("sex"))
        sex_total[sx] += 1
        if cat == "FINISH":
            sex_fin[sx] += 1
            if t is not None:
                sex_times.setdefault(sx, []).append(t)
        elif cat == "DNS":
            sex_dns[sx] += 1
        elif cat == "IN RACE":
            sex_inrace[sx] += 1

        cps = parse_checkpoints(details.get(row.get("id")), status=cat) if details else []
        enriched.append({
            "order": row.get("order"),
            "name": row.get("name"),
            "bib": row.get("no"),
            "status": cat,
            "status_code": row.get("status"),
            "status_name": row.get("status_name"),
            "cp_name": row.get("cp_name"),
            "time_used": row.get("time_used"),
            "time_seconds": t if t is not None else "",
            "sex": sx,
            "result_jing": row.get("result_jing"),
            "result_qiang": row.get("result_qiang"),
            "avg_peisu": row.get("avg_peisu"),
            "age_range": row.get("age_range"),
            "team": row.get("team"),
            "country": row.get("country"),
            "all_order": row.get("all_order"),
            "age_order": row.get("age_order"),
            "age_sex_order": row.get("age_sex_order"),
            "is_lack": row.get("is_lack"),
            "sign_id": row.get("sign_id"),
            "player_id": row.get("id"),
            "checkpoints": cps,
        })

    rank_checkpoints(enriched)

    dns = statuses.get("DNS", 0)
    starters = len(rows) - dns
    fin = statuses.get("FINISH", 0)
    non_fin_starters = starters - fin

    summary = {
        "event": event,
        "name": name,
        "rows_returned": len(rows),
        "api_total": api_total,
        "status": dict(statuses),
        "dns": dns,
        "starters": starters,
        "finished": fin,
        "non_finish_starters": non_fin_starters,
        "finish_rate_starters": (fin / starters) if starters else 0.0,
        "finish_rate_total": (fin / len(rows)) if rows else 0.0,
        "rows": enriched,
    }
    # TODO: checkpoints are matched by cp_name here, in export_xlsx() and in rank_checkpoints(). On a loop course
    #       that passes the same named checkpoint twice the passes would be merged and ranked together - key by
    #       position (cp_order) or cp_id instead. Not yet seen in real data.
    cp_columns = []
    for r in enriched:
        cps = [c for c in r["checkpoints"] if not c["is_start"]]
        if cps:
            cp_columns = [c["cp_name"] for c in cps]
            break
    summary["cp_columns"] = cp_columns

    by_sex = {}
    for sx in sorted(sex_total):
        s_start = sex_total[sx] - sex_dns[sx]
        s_fin = sex_fin[sx]
        st = sorted(sex_times.get(sx, []))
        by_sex[sx] = {
            "total": sex_total[sx],
            "dns": sex_dns[sx],
            "starters": s_start,
            "in_race": sex_inrace[sx],
            "finished": s_fin,
            "finish_rate": (s_fin / s_start) if s_start else 0.0,
            "fastest": fmt_secs(st[0]) if st else "",
            "median": fmt_secs(statistics.median(st)) if st else "",
            "slowest": fmt_secs(st[-1]) if st else "",
        }
    summary["by_sex"] = by_sex

    if times_sec:
        times_sec.sort()
        summary["finish_time_stats"] = {
            "n": len(times_sec),
            "fastest": fmt_secs(times_sec[0]),
            "slowest": fmt_secs(times_sec[-1]),
            "median": fmt_secs(statistics.median(times_sec)),
            "mean": fmt_secs(statistics.mean(times_sec)),
            "q1": fmt_secs(percentile(times_sec, 25)),
            "q3": fmt_secs(percentile(times_sec, 75)),
        }
        summary["hourly_histogram"] = hourly_histogram(times_sec)

    return summary


def hourly_histogram(times_sec):
    buckets = Counter(t // 3600 for t in times_sec)
    return [{"band": f"{h:02d}:00-{h+1:02d}:00", "count": buckets[h]}
            for h in sorted(buckets)]


# ----------------------------------------------------------------------------
# Console output
# ----------------------------------------------------------------------------
def print_race(summary):
    title = f"{summary['event']} - {summary['name']}"
    print(f"\n{title}")
    print(f"{'=' * _line_width(title)}")

    # Divide-by-zero guard: an empty race still renders a (zeroed) table
    # instead of raising. `or 1` keeps every percentage at 0.0%.
    total = summary["rows_returned"] or 1

    headers = ["Status", "Count", "% of total"]
    s = summary["status"]
    status_rows = [
        ("DNS (待参赛)",        s.get("DNS", 0)),
        ("In race",            s.get("IN RACE", 0)),
        ("Timed out (被关门)", s.get("TIMEOUT", 0)),
        ("Withdrew (已退赛)",  s.get("WITHDREW", 0)),
        ("Disqualified",       s.get("DSQ", 0)),
        ("Other",              s.get("OTHER", 0)),
        ("Finished",           s.get("FINISH", 0)),
    ]
    print(render_table(
        [[label, n, f"{100*n/total:.1f}%"] for label, n in status_rows],
        headers,
        colalign=("left", "right", "right"),
        maxwidths=[22, None, None],
    ))
    print(f"\n  Total: {summary['rows_returned']}")
    print(f"  Starters (total - DNS): {summary['starters']}")
    print(f"  Finishers (of starters): {100*summary['finish_rate_starters']:>5.1f}%")

    if summary.get("by_sex"):
        print()
        headers = ["Sex", "Total", "DNS", "Start", "In race", "Fin", "Fin%", "Fastest", "Median", "Slowest"]
        rows = []
        for sx, d in summary["by_sex"].items():
            rows.append([
                sx, d["total"], d["dns"], d["starters"], d["in_race"],
                d["finished"], f"{100*d['finish_rate']:.1f}%",
                d["fastest"], d["median"], d["slowest"],
            ])
        print(render_table(
            rows, headers,
            colalign=("center",) + ("right",) * (len(headers) - 1),
        ))

    if "finish_time_stats" in summary:
        print(f"\n {summary['event']} - {summary['name']} finish-times:\n")
        dnf = s.get("TIMEOUT", 0) + s.get("WITHDREW", 0) + s.get("DSQ", 0)
        dns = s.get("DNS", 0)
        inrace = s.get("IN RACE", 0)
        peak = max(
            [b["count"] for b in summary["hourly_histogram"]] + [dns, dnf, inrace]
        ) or 1
        for b in summary["hourly_histogram"]:
            bar = "█" * max(1, round(b["count"] / peak * 40))
            print(f"    {b['band']:<14}{b['count']:>4}  {bar}")
        print()
        for label, count in (("DNS", dns), ("DNF", dnf), ("In race", inrace)):
            bar = "█" * max(1, round(count / peak * 40)) if count else ""
            print(f"    {label:<14}{count:>4}  {bar}")
        ts = summary["finish_time_stats"]
        print(f"\n  n={ts['n']}  fastest={ts['fastest']}  "
              f"Q1={ts['q1']}  median={ts['median']}  mean={ts['mean']}  "
              f"Q3={ts['q3']}  slowest={ts['slowest']}")


def print_combined(all_summaries):
    if not all_summaries:
        return

    title = all_summaries[0].get("event", "Combined Summary")
    print(f"\n{title}")
    print(f"{'='*_line_width(title)}")
    headers = ["Race", "Start time", "Fastest", "Q1", "Median", "Q3", "Slowest"]
    rows = []
    for r in all_summaries:
        s = r["status"]
        ft = r.get("finish_time_stats", {})
        rows.append([
            r["name"],
            r["qiang_at"],
            ft.get("fastest", ""),
            ft.get("q1", ""),
            ft.get("median", ""),
            ft.get("q3", ""),
            ft.get("slowest", ""),
        ])
    print(render_table(
        rows, headers,
        colalign=("left",) + ("center",) * (len(headers) - 1),
        maxwidths=[24] + [19] + [8] * (len(headers) - 2),
    ))
    print()
    headers = ["Race", "Total", "DNS", "Start", "In race", "TO", "WD", "DSQ", "Other", "Finish", "Fin% *"]
    rows = []
    for r in all_summaries:
        s = r["status"]
        rows.append([
            r["name"],
            r["rows_returned"],
            r["dns"],
            r["starters"],
            s.get("IN RACE", 0),
            s.get("TIMEOUT", 0),
            s.get("WITHDREW", 0),
            s.get("DSQ", 0),
            s.get("OTHER", 0),
            r["finished"],
            f"{100*r['finish_rate_starters']:.1f}%",
        ])
    print(render_table(
        rows, headers,
        colalign=("left",) + ("right",) * (len(headers) - 1),
        maxwidths=[24] + [5] * (len(headers) - 1),
    ))
    print("\n * Fin% = finishers / (total - DNS)")

def export_xlsx(all_summaries, event_name):
    if not all_summaries:
        print("\nNo races could be summarised - no spreadsheet written.")
        return

    wb = Workbook()
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDDDDD")

    # ---- Summary sheet ----
    ws = wb.active
    ws.title = "Summary"
    ws.append([event_name])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    cols = [
        ("Race", 14),               # A
        ("Start (gun)", 18),        # B
        ("Total", 8),               # C
        ("DNS", 8),                 # D
        ("Starters", 8),            # E
        ("Timed out", 8),           # F
        ("Withdrew", 8),            # G
        ("DSQ", 8),                 # H
        ("Finished", 8),            # I
        ("Finish% (starters)", 18), # J
        ("Fastest", 10),            # K
        ("Q1", 10),                 # L
        ("Median", 10),             # M
        ("Mean", 10),               # N
        ("Q3", 10),                 # O
        ("Slowest", 10),            # P
    ]
    ws.append([h for h, _ in cols])
    for c in range(1, len(cols) + 1):
        ws.cell(row=3, column=c).font = header_font
        ws.cell(row=3, column=c).fill = header_fill
    for r in all_summaries:
        s = r["status"]
        ft = r.get("finish_time_stats", {})
        row = ws.max_row + 1
        ws.append([
            r["name"],                          # A
            r.get("qiang_at", ""),              # B
            r["rows_returned"],                 # C
            r["dns"],                           # D
            r["starters"],                      # E
            s.get("TIMEOUT", 0),                # F
            s.get("WITHDREW", 0),               # G
            s.get("DSQ", 0),                    # H
            r["finished"],                      # I
            f"=IF(E{row}=0,0,I{row}/E{row})",   # J: Finished / Starters
            ft.get("fastest", ""),              # K
            ft.get("q1", ""),                   # L
            ft.get("median", ""),               # M
            ft.get("mean", ""),                 # N
            ft.get("q3", ""),                   # O
            ft.get("slowest", ""),              # P
        ])
        ws.cell(row=row, column=10).number_format = "0.0%"
    for i, (_, w) in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- Gender sheet (one row per race per sex) ----
    ws = wb.create_sheet(title="By gender")
    cols = [
        ("Race", 14),               # A
        ("Sex", 6),                 # B
        ("Total", 8),               # C
        ("DNS", 8),                 # D
        ("Starters", 9),            # E
        ("Finished", 9),            # F
        ("Finish% (starters)", 18), # G
        ("Fastest", 10),            # H
        ("Median", 10),             # I
        ("Slowest", 10),            # J
    ]
    ws.append([h for h, _ in cols])
    for c in range(1, len(cols) + 1):
        ws.cell(row=1, column=c).font = header_font
        ws.cell(row=1, column=c).fill = header_fill
    for r in all_summaries:
        for sx, d in r.get("by_sex", {}).items():
            row = ws.max_row + 1
            ws.append([
                r["name"],
                sx,
                d["total"],
                d["dns"],
                d["starters"],
                d["finished"],
                f"=IF(E{row}=0,0,F{row}/E{row})",
                d["fastest"],
                d["median"],
                d["slowest"],
            ])
            ws.cell(row=row, column=7).number_format = "0.0%"
    for i, (_, w) in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # ---- One sheet per race ----
    cols = [
        ("order", 7),
        ("name", 22),
        ("bib", 8),
        ("sex", 5),
        ("status", 10),
        ("status_code", 11),
        ("status_name", 18),
        ("result_jing", 10),
        ("result_qiang", 10),
        ("time_used", 10),
        ("time_seconds", 12),
        ("avg_peisu", 9),
        ("cp_name", 22),
        ("age_range", 10),
        ("team", 18),
        ("country", 14),
        ("all_order", 9),
        ("age_order", 9),
        ("age_sex_order", 11),
        ("is_lack", 8),
        ("sign_id", 11),
        ("player_id", 11),
    ]
    keys = [k for k, _ in cols]
    for r in all_summaries:
        ws = wb.create_sheet(title=excel_sheet_title(r["name"]))
        cp_cols = r.get("cp_columns", [])
        header = (
            [h for h, _ in cols]
            + cp_cols
            + [f"rank {h}" for h in cp_cols]
            + [f"elapsed rank {h}" for h in cp_cols]
        )
        ws.append(header)
        for c in range(1, len(header) + 1):
            ws.cell(row=1, column=c).font = header_font
            ws.cell(row=1, column=c).fill = header_fill

        for row in r["rows"]:
            values = []
            for k in keys:
                v = row.get(k, "")
                if k == "avg_peisu":
                    v = "" if row.get("status") == "DNS" else _to_float_or_blank(v)
                values.append(v)

            # CP tail: cumulative time_used in CP order, '' for unreached.
            by_name = {c["cp_name"]: c for c in row.get("checkpoints", [])}
            for cp_name in cp_cols:
                c = by_name.get(cp_name)
                values.append(
                    "" if not c or c.get("is_start") or not c["reached"]
                    else c["time_used"]
                )
            # Arrival-order rank (start wall-clock + cumulative time_used).
            for cp_name in cp_cols:
                c = by_name.get(cp_name)
                values.append(c.get("rank", "") if c and c["reached"] else "")
            # Elapsed-time rank (cumulative time_used alone).
            for cp_name in cp_cols:
                c = by_name.get(cp_name)
                values.append(c.get("elapsed_rank", "") if c and c["reached"] else "")

            ws.append(values)

        widths = (
            [w for _, w in cols]
            + [10] * len(cp_cols)   # time_used per CP
            + [7] * len(cp_cols)    # rank CPx
            + [7] * len(cp_cols)    # elapsed rank CPx
        )
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "C2"
        peisu_col = keys.index("avg_peisu") + 1
        for cell_row in range(2, ws.max_row + 1):
            ws.cell(row=cell_row, column=peisu_col).number_format = "0.00"

    # ---- Long-format checkpoint sheet ----
    if FETCH_CHECKPOINTS:
        ws = wb.create_sheet(title="Checkpoints")
        cols = [
            ("race", 12),
            ("player_id", 11),
            ("bib", 8),
            ("name", 20),
            ("sex", 5),
            ("status", 10),
            ("cp_order", 8),
            ("cp_name", 22),
            ("cp_mileage", 10),
            ("when", 14),
            ("time_used", 10),
            ("time_seconds", 12),
            ("speed_allocation", 14),
            ("status_label", 10),
            ("closing_time", 12),
            ("is_warn", 8),
            ("reached", 8),
            ("is_start", 8),
            ("rank", 7),
        ]
        ws.append([h for h, _ in cols])
        for c in range(1, len(cols) + 1):
            ws.cell(row=1, column=c).font = header_font
            ws.cell(row=1, column=c).fill = header_fill

        for r in all_summaries:
            for row in r["rows"]:
                for cp in row.get("checkpoints", []):
                    ws.append([
                        r["name"],
                        row.get("player_id"),
                        row.get("bib"),
                        row.get("name"),
                        row.get("sex"),
                        row.get("status"),
                        cp["cp_order"],
                        cp["cp_name"],
                        cp["cp_mileage"],
                        cp["when"],
                        "" if cp["is_start"] else cp["time_used"],
                        "" if cp["is_start"] else cp["time_seconds"],
                        "" if cp["is_start"] else cp["speed_allocation"],
                        cp["status_label"],
                        cp["closing_time"],
                        cp["is_warn"],
                        cp["reached"],
                        cp["is_start"],
                        cp.get("rank", ""),   # only set for reached, non-start checkpoints
                    ])
        for i, (_, w) in enumerate(cols, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    filename = safe_filename(event_name) + ".xlsx"
    wb.save(filename)
    print(f"\nWrote {filename}")

# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main(argv=None):

    parser = argparse.ArgumentParser(
        description="Scrape weibu race results and build a report.",
    )
    parser.add_argument(
        "--event-id", type=_event_id_spec, action="extend", default=None,
        metavar="N|LO-HI",
        help="weibu event_id, or an inclusive range like 6000-6500; repeatable",
    )
    parser.add_argument(
        "--all-events", action="store_true",
        help="list every known event (from the registry) and exit",
    )
    parser.add_argument(
        "--cache-dir", default="cache",
        help="directory for cached API responses (default: cache)",
    )
    parser.add_argument(
        "--clear-cache", action="store_true",
        help="delete cached API responses before running (keeps the list of known event ids)",
    )
    parser.add_argument(
        "--no-details", action="store_true",
        help="skip the per-player detail calls (no checkpoint columns)",
    )
    parser.add_argument(
        "--attachments", action="store_true",
        help="download attachments (GPX/KML) for --event-id and exit",
    )
    parser.add_argument(
        "--attachment-types", default=None,
        help="comma-separated extensions to download, e.g. 'gpx' (default: all)",
    )
    parser.add_argument(
        "--attachment-dir", default=None,
        help="output directory (default: cache/attachments/{event_id}/)",
        )
    parser.add_argument(
        "--refresh", action="store_true",
        help="ignore list/event/detail caches and refetch from the API",
    )
    parser.add_argument(
        "--find-player", metavar="NAME", default=None,
        help="search cached race rosters for athletes whose name contains NAME, and exit",
    )
    parser.add_argument(
        "--exact-name", action="store_true",
        help="with --find-player, require the whole name to match",
    )
    parser.add_argument(
        "--case-sensitive", action="store_true",
        help="with --find-player, match case-sensitively",
    )
    args = parser.parse_args(argv)

    configure(
        fetch_details=not args.no_details,
        cache_dir=args.cache_dir,
    )
    if args.clear_cache:
        clear_cache()
    use_cache = not args.refresh
    
    if args.all_events:
        ids = sorted(set(_load_known_events()) | set(args.event_id or []))
        for eid, name in enumerate_events(ids, use_cache=use_cache, skip_errors=True):
            print(f"  {eid}\t{name}")
        return

    if args.find_player is not None:
        hits = find_player(
            args.find_player,
            case_sensitive=args.case_sensitive,
            exact=args.exact_name,
        )
        if not hits:
            print(f"No cached athletes match {args.find_player!r}.")
        else:
            print(f"{len(hits)} appearance(s) of {args.find_player!r}:\n")
            headers = ["Name", "Event", "Race", "Start", "bib", "Status", "Time", "eid", "pid"]
            rows = [[
                h["name"] or "",
                h["event"] or "",
                h["race"] or "",
                h["start"] or "",
                h["bib"] or "",
                h["status"],
                h["time_used"] or "",
                h["event_id"],
                h["player_id"],
            ] for h in hits]
            print(render_table(
                rows, headers,
                colalign=('left',) * len(headers),
                maxwidths=[None] * len(headers),
            ))
        return

    
    if not args.event_id:
        parser.error("--event-id is required unless --all-events or --find-player is given")

    resolved = enumerate_events(args.event_id, use_cache=use_cache, skip_errors=True)
    if not resolved:
        print(f"None of {args.event_id} resolved to a real event.")
        return
    # TODO: several ids or a range only process the first event that resolves - the rest are checked and then
    #       silently ignored. Process each event in turn, or say that only the first is used.
    activity_id, event = resolved[0]
    print(f"Event: {event}\n")

    if args.attachments:
        types = (
            {t.strip().lower() for t in args.attachment_types.split(",") if t.strip()}
            if args.attachment_types else None
        )
        paths = download_attachments(activity_id, out_dir=args.attachment_dir, types=types)
        print(f"\n{len(paths)} file(s) downloaded.")
        return

    races = get_races(activity_id, use_cache=use_cache)
    print("Race categories:")
    for r in races:
        start = r.get("qiang_at") or "-"
        print(f"  {r['name']} (id: {r['package_id']}) start: {start}")

    all_summaries = []
    for race in races:
        print()
        try:
            rows, api_total = fetch_all_players(activity_id, race["package_id"], race["name"], sex="", use_cache=use_cache)
            details = fetch_all_details(rows, race["name"], use_cache=use_cache) if FETCH_CHECKPOINTS else {}
            summary = summarize_race(event, race["name"], rows, api_total, details)
            summary["qiang_at"] = race.get("qiang_at", "")
            summary["ready_qiang_at"] = race.get("ready_qiang_at", "")
            all_summaries.append(summary)
            print_race(summary)
        except Exception as e:
            print(f"Error occurred while processing race {race['name']}: {e}")

    print_combined(all_summaries)

    export_xlsx(all_summaries, event)


if __name__ == "__main__":
    main()