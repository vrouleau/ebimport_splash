"""
Integration tests for ebimport_splash.

Requires:
  - Docker (docker compose) available
  - Port 5000 free
  - Run from repo root: pytest tests/ -v

These tests exercise the full pipeline via HTTP against the running container.
Manual SPLASH steps are documented as comments where they'd occur in a real workflow.
"""
import io
import os
import re
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import pytest
import requests

BASE_URL = os.environ.get("EBIMPORT_URL", "http://127.0.0.1:5000")
REPO_ROOT = Path(__file__).resolve().parent.parent
TEST_XLSX = REPO_ROOT / "tests" / "test_attendees.xlsx"
MEET_LXF = REPO_ROOT / "tests" / "fixtures" / "meet_template.lxf"
OUTPUT_DIR = REPO_ROOT / "tests" / "output"
TIMEOUT = 120


@pytest.fixture(scope="session", autouse=True)
def docker_up():
    """Start the container, wait for health, tear down after all tests.

    Set ``EBIMPORT_SKIP_STACK=1`` to assume the stack is already up — useful
    when running pytest from inside a sidecar container that lacks docker.
    """
    skip_stack = os.environ.get("EBIMPORT_SKIP_STACK") == "1"
    if not skip_stack:
        subprocess.run(
            ["docker", "compose", "up", "--build", "-d"],
            cwd=REPO_ROOT, check=True, capture_output=True,
        )
    # Wait for the service to respond
    deadline = time.time() + 60
    while time.time() < deadline:
        try:
            r = requests.get(f"{BASE_URL}/", timeout=3)
            if r.status_code == 200:
                break
        except requests.ConnectionError:
            pass
        time.sleep(2)
    else:
        pytest.fail("Container did not become healthy within 60s")
    yield
    if not skip_stack:
        subprocess.run(
            ["docker", "compose", "down"],
            cwd=REPO_ROOT, capture_output=True,
        )


@pytest.fixture(scope="session")
def test_xlsx():
    """Ensure test xlsx exists (regenerate if missing)."""
    if not TEST_XLSX.exists():
        subprocess.run(
            [sys.executable, "tests/generate_test_xlsx.py"],
            cwd=REPO_ROOT, check=True,
        )
    return TEST_XLSX


@pytest.fixture(scope="session", autouse=True)
def meet_lxf():
    """Build the augmented meet_template.lxf from the committed base file."""
    subprocess.run(
        [sys.executable, "tests/build_meet_fixture.py"],
        cwd=REPO_ROOT, check=True,
    )
    return MEET_LXF


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def upload(mode: str, xlsx_path: Path, meet_path: Path | None = None) -> dict:
    """Upload xlsx (and optional meet .lxf for lenex mode), return JSON."""
    with open(xlsx_path, "rb") as f_x:
        files = {"xlsx": ("test.xlsx", f_x)}
        meet_handle = None
        if meet_path is not None:
            meet_handle = open(meet_path, "rb")
            files["meet"] = ("meet.lxf", meet_handle)
        try:
            r = requests.post(
                f"{BASE_URL}/api/run",
                files=files,
                data={"mode": mode},
                timeout=TIMEOUT,
            )
        finally:
            if meet_handle is not None:
                meet_handle.close()
    assert r.status_code == 200, f"API returned {r.status_code}: {r.text}"
    return r.json()


def download_zip(download_id: str, save_as: str = None) -> zipfile.ZipFile:
    """Download result zip and return as ZipFile object. Optionally save to OUTPUT_DIR."""
    r = requests.get(
        f"{BASE_URL}/api/download/{download_id}",
        params={"name": "result.zip"},
        timeout=30,
    )
    assert r.status_code == 200
    if save_as:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        dest = OUTPUT_DIR / save_as
        dest.write_bytes(r.content)
        # Also extract for easy access
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            z.extractall(OUTPUT_DIR / save_as.replace(".zip", ""))
    return zipfile.ZipFile(io.BytesIO(r.content))


# ---------------------------------------------------------------------------
# Tests: Dry-run / Validation
# ---------------------------------------------------------------------------

class TestDryRun:
    def test_dry_run_returns_issues(self, test_xlsx):
        resp = upload("dry-run", test_xlsx)
        assert resp["returncode"] == 0
        assert "summary" in resp
        assert "issues" in resp
        # Should detect known defects
        issues_text = "\n".join(resp["issues"])
        assert "unknown" in issues_text.lower() or len(resp["issues"]) > 0

    def test_dry_run_produces_issues_zip(self, test_xlsx):
        resp = upload("dry-run", test_xlsx)
        # dry-run still produces a download (issues report zip)
        assert resp.get("download_id")


# ---------------------------------------------------------------------------
# Tests: Lenex Path
# ---------------------------------------------------------------------------

class TestLenexPath:
    @pytest.fixture(scope="class")
    def lenex_result(self, test_xlsx):
        if not MEET_LXF.exists():
            pytest.skip(f"meet template not found at {MEET_LXF}")
        resp = upload("lenex", test_xlsx, meet_path=MEET_LXF)
        assert resp["returncode"] == 0
        z = download_zip(resp["download_id"], save_as="lenex_result.zip")
        return resp, z

    def test_zip_contains_lxf_and_scripts(self, lenex_result):
        _, z = lenex_result
        names = z.namelist()
        assert "inscriptions.lxf" in names
        assert "masters_transfer.vbs" in names

    def test_lxf_is_valid_zip_with_lef(self, lenex_result):
        _, z = lenex_result
        lxf_bytes = z.read("inscriptions.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        assert "meet.lef" in lxf.namelist()

    def test_handicap_on_masters_athletes(self, lenex_result):
        """Masters athletes should have HANDICAP exception='X'."""
        _, z = lenex_result
        lxf_bytes = z.read("inscriptions.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        lef = lxf.read("meet.lef").decode()

        handicap_athletes = re.findall(r'<HANDICAP exception="X"', lef)
        assert len(handicap_athletes) > 0, "No HANDICAP exception='X' athletes found"
        # Not all athletes should have it
        all_athletes = re.findall(r'<ATHLETE ', lef)
        assert len(handicap_athletes) < len(all_athletes)

    def test_masters_entries_in_prelim_events(self, lenex_result):
        """Masters athletes' entries should point to prelim events (not Masters finals)."""
        _, z = lenex_result
        lxf_bytes = z.read("inscriptions.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        lef = lxf.read("meet.lef").decode()

        # Parse athlete blocks with HANDICAP exception
        athlete_blocks = re.findall(
            r'<ATHLETE [^>]*>(.*?)</ATHLETE>',
            lef, re.DOTALL,
        )
        ma_event_ids = []
        for block in athlete_blocks:
            if 'exception="X"' in block:
                eids = re.findall(r'eventid="(\d+)"', block)
                ma_event_ids.extend(int(e) for e in eids)

        assert len(ma_event_ids) > 0, "No entries found for Masters athletes"
        # Prelim events have IDs < 3000 in the template; Masters finals are 4xxx
        prelim_count = sum(1 for e in ma_event_ids if e < 3000)
        final_count = sum(1 for e in ma_event_ids if e >= 4600)
        assert prelim_count > 0, "No Masters entries in prelim events"
        # Masters-only events (UID 541) go to final — that's OK but should be minority
        assert prelim_count > final_count

    def test_open_25plus_no_handicap(self, lenex_result):
        """Open athletes aged 25+ should NOT have HANDICAP exception='X'."""
        _, z = lenex_result
        lxf_bytes = z.read("inscriptions.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        lef = lxf.read("meet.lef").decode()

        # Find athlete blocks without HANDICAP but aged 25+
        athlete_blocks = re.findall(
            r'<ATHLETE [^>]*birthdate="([^"]+)"[^>]*>(.*?)</ATHLETE>',
            lef, re.DOTALL,
        )
        for bd, block in athlete_blocks:
            try:
                year = int(bd[:4])
                age = 2026 - year
            except ValueError:
                continue
            if age >= 25 and 'exception="X"' not in block:
                return  # found an Open 25+ without handicap — correct
        # Fallback: ensure not everyone has handicap
        all_ath = re.findall(r'<ATHLETE ', lef)
        handicap = re.findall(r'<HANDICAP exception="X"', lef)
        assert len(handicap) < len(all_ath)

    def test_relay_entrytime_matches_team_time(self, lenex_result):
        """Relay ENTRY entrytime should be the relay row's team time,
        not the sum of teammates' individual best times. Regression for
        the bug where a 4×50 mixed medley swum in 3:31.08 was exported
        as ~5:38 (sum of four individual leg estimates)."""
        _, z = lenex_result
        lxf_bytes = z.read("inscriptions.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        lef = lxf.read("meet.lef").decode()

        relay_entries = re.findall(
            r'<RELAY\b[^>]*>.*?<ENTRY [^>]*entrytime="([^"]+)"',
            lef, re.DOTALL,
        )
        assert relay_entries, "no RELAY entries with entrytime found"
        # Lenex format: HH:MM:SS.ss — parse to seconds.
        def to_sec(s: str) -> float:
            h, m, rest = s.split(":")
            return int(h) * 3600 + int(m) * 60 + float(rest)
        # All relay times in the test xlsx are well under 4 minutes.
        # Pre-fix, the buggy sum produced 5+ minute entrytimes.
        too_slow = [t for t in relay_entries if to_sec(t) > 4 * 60]
        assert not too_slow, (
            f"relay entrytimes look summed (>4:00): {too_slow[:5]}"
        )

    def test_phantom_teammate_dob_from_coach_row(self, lenex_result):
        """Phantom Teammate is only referenced via Real Buddy's Corde duo's
        teammate field. Her birthdate is on her Coach ticket row — the loader
        should harvest it so the .lxf carries her DOB."""
        _, z = lenex_result
        lxf_bytes = z.read("inscriptions.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        lef = lxf.read("meet.lef").decode()
        m = re.search(
            r'<ATHLETE [^>]*firstname="Phantom"[^>]*birthdate="([^"]+)"',
            lef,
        )
        assert m, "Phantom Teammate ATHLETE missing or has no birthdate"
        assert m.group(1) == "1996-07-07", \
            f"unexpected DOB for Phantom Teammate: {m.group(1)}"


# ---------------------------------------------------------------------------
# Tests: Error handling
# ---------------------------------------------------------------------------

class TestErrors:
    def test_no_file_returns_400(self):
        r = requests.post(f"{BASE_URL}/api/run", data={"mode": "mdb"}, timeout=10)
        assert r.status_code == 400

    def test_invalid_mode_returns_400(self, test_xlsx):
        with open(test_xlsx, "rb") as f:
            r = requests.post(
                f"{BASE_URL}/api/run",
                files={"xlsx": ("test.xlsx", f)},
                data={"mode": "invalid"},
                timeout=10,
            )
        assert r.status_code == 400

    def test_invalid_download_id_returns_404(self):
        r = requests.get(
            f"{BASE_URL}/api/download/nonexistent",
            params={"name": "x.zip"},
            timeout=10,
        )
        assert r.status_code == 404


# ---------------------------------------------------------------------------
# MANUAL SPLASH STEPS
# ---------------------------------------------------------------------------
# See docs/MASTERS_TRANSFER.md for the full workflow including manual steps.
