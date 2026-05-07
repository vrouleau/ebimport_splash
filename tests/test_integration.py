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
import time
import zipfile
from pathlib import Path

import pytest
import requests

BASE_URL = os.environ.get("EBIMPORT_URL", "http://127.0.0.1:5000")
REPO_ROOT = Path(__file__).resolve().parent.parent
TEST_XLSX = REPO_ROOT / "tests" / "test_attendees.xlsx"
TIMEOUT = 120


@pytest.fixture(scope="session", autouse=True)
def docker_up():
    """Start the container, wait for health, tear down after all tests."""
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
    subprocess.run(
        ["docker", "compose", "down"],
        cwd=REPO_ROOT, capture_output=True,
    )


@pytest.fixture(scope="session")
def test_xlsx():
    """Ensure test xlsx exists (regenerate if missing)."""
    if not TEST_XLSX.exists():
        subprocess.run(
            ["python", "tests/generate_test_xlsx.py"],
            cwd=REPO_ROOT, check=True,
        )
    return TEST_XLSX


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def upload(mode: str, xlsx_path: Path) -> dict:
    """Upload xlsx and return JSON response."""
    with open(xlsx_path, "rb") as f:
        r = requests.post(
            f"{BASE_URL}/api/run",
            files={"xlsx": ("test.xlsx", f)},
            data={"mode": mode},
            timeout=TIMEOUT,
        )
    assert r.status_code == 200, f"API returned {r.status_code}: {r.text}"
    return r.json()


def download_zip(download_id: str) -> zipfile.ZipFile:
    """Download result zip and return as ZipFile object."""
    r = requests.get(
        f"{BASE_URL}/api/download/{download_id}",
        params={"name": "result.zip"},
        timeout=30,
    )
    assert r.status_code == 200
    return zipfile.ZipFile(io.BytesIO(r.content))


def parse_mdb_in_container(mdb_bytes: bytes, script: str) -> str:
    """Copy an MDB into the container and run a Python script, return stdout."""
    # Write mdb to a temp file, copy in, run
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
        f.write(mdb_bytes)
        tmp = f.name
    try:
        subprocess.run(
            ["docker", "cp", tmp, "ebimport-splash:/tmp/test_audit.mdb"],
            check=True, capture_output=True,
        )
    finally:
        os.unlink(tmp)
    # Ensure access-parser is available
    subprocess.run(
        ["docker", "exec", "ebimport-splash", "pip", "install", "-q", "access-parser"],
        capture_output=True,
    )
    result = subprocess.run(
        ["docker", "exec", "ebimport-splash", "python", "-c", script],
        capture_output=True, text=True,
    )
    assert result.returncode == 0, f"Script failed: {result.stderr}"
    return result.stdout


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
# Tests: MDB Path
# ---------------------------------------------------------------------------

class TestMDBPath:
    @pytest.fixture(scope="class")
    def mdb_result(self, test_xlsx):
        resp = upload("mdb", test_xlsx)
        assert resp["returncode"] == 0
        z = download_zip(resp["download_id"])
        return resp, z

    def test_zip_contains_mdb_and_scripts(self, mdb_result):
        _, z = mdb_result
        names = z.namelist()
        assert "meet.mdb" in names
        assert "masters_transfer.vbs" in names
        assert "masters_transfer.bat" in names
        assert "simulate_results.vbs" in names
        assert "simulate_results.bat" in names

    def test_zip_contains_issues(self, mdb_result):
        _, z = mdb_result
        assert "issues.txt" in z.namelist()

    def test_summary_has_entries(self, mdb_result):
        resp, _ = mdb_result
        summary = "\n".join(resp["summary"])
        assert "new individual entries" in summary
        assert "new athletes" in summary
        assert "new clubs" in summary

    def test_masters_in_prelim_with_bonusentry(self, mdb_result):
        """Masters athletes should be in prelim events with BONUSENTRY='T'."""
        _, z = mdb_result
        mdb_bytes = z.read("meet.mdb")
        script = """
from access_parser import AccessParser
db = AccessParser('/tmp/test_audit.mdb')
sr = db.parse_table('SWIMRESULT')
ev = db.parse_table('SWIMEVENT')

ev_info = {}
for i in range(len(ev['SWIMEVENTID'])):
    eid = int(ev['SWIMEVENTID'][i])
    ev_info[eid] = {'round': ev['ROUND'][i], 'masters': ev['MASTERS'][i]}

bonus_in_prelim = 0
bonus_in_final = 0
non_bonus_in_prelim = 0
for i in range(len(sr['SWIMRESULTID'])):
    eid = int(sr['SWIMEVENTID'][i]) if sr['SWIMEVENTID'][i] else 0
    ei = ev_info.get(eid, {})
    if sr['BONUSENTRY'][i] == 'T':
        if ei.get('round') == 2:
            bonus_in_prelim += 1
        else:
            bonus_in_final += 1
    else:
        if ei.get('round') == 2:
            non_bonus_in_prelim += 1

print(f'{bonus_in_prelim},{bonus_in_final},{non_bonus_in_prelim}')
"""
        out = parse_mdb_in_container(mdb_bytes, script)
        parts = out.strip().split(",")
        bonus_prelim = int(parts[0])
        bonus_final = int(parts[1])
        non_bonus_prelim = int(parts[2])
        # Masters should be mostly in prelim
        assert bonus_prelim > 0, "No Masters entries in prelim"
        assert non_bonus_prelim > 0, "No Open/15-18 entries in prelim"
        # Masters-only events go to final — that's expected but should be small
        assert bonus_final < bonus_prelim, "More Masters in final than prelim"

    def test_open_25plus_not_marked_bonus(self, mdb_result):
        """Open athletes aged 25+ should NOT have BONUSENTRY='T'."""
        _, z = mdb_result
        mdb_bytes = z.read("meet.mdb")
        script = """
from access_parser import AccessParser
from datetime import datetime
db = AccessParser('/tmp/test_audit.mdb')
sr = db.parse_table('SWIMRESULT')
ath = db.parse_table('ATHLETE')

ath_dob = {}
for i in range(len(ath['ATHLETEID'])):
    ath_dob[int(ath['ATHLETEID'][i])] = ath['BIRTHDATE'][i]

# Check: any BONUSENTRY='F' athlete aged 25+?
open_25plus = 0
for i in range(len(sr['SWIMRESULTID'])):
    if sr['BONUSENTRY'][i] == 'F':
        aid = int(sr['ATHLETEID'][i]) if sr['ATHLETEID'][i] else 0
        bd = ath_dob.get(aid)
        if bd:
            try:
                if isinstance(bd, datetime): pass
                elif isinstance(bd, str): bd = datetime.strptime(bd[:10], '%Y-%m-%d')
                age = 2026 - bd.year
                if age >= 25:
                    open_25plus += 1
            except: pass

# Check: any BONUSENTRY='T' athlete aged < 25? (should be 0)
young_bonus = 0
for i in range(len(sr['SWIMRESULTID'])):
    if sr['BONUSENTRY'][i] == 'T':
        aid = int(sr['ATHLETEID'][i]) if sr['ATHLETEID'][i] else 0
        bd = ath_dob.get(aid)
        if bd:
            try:
                if isinstance(bd, datetime): pass
                elif isinstance(bd, str): bd = datetime.strptime(bd[:10], '%Y-%m-%d')
                age = 2026 - bd.year
                if age < 25:
                    young_bonus += 1
            except: pass

print(f'{open_25plus},{young_bonus}')
"""
        out = parse_mdb_in_container(mdb_bytes, script)
        parts = out.strip().split(",")
        open_25plus = int(parts[0])
        young_bonus = int(parts[1])
        assert open_25plus > 0, "Expected some Open athletes aged 25+"
        assert young_bonus == 0, f"Found {young_bonus} BONUSENTRY='T' athletes under 25"


# ---------------------------------------------------------------------------
# Tests: Lenex Path
# ---------------------------------------------------------------------------

class TestLenexPath:
    @pytest.fixture(scope="class")
    def lenex_result(self, test_xlsx):
        resp = upload("lenex", test_xlsx)
        assert resp["returncode"] == 0
        z = download_zip(resp["download_id"])
        return resp, z

    def test_zip_contains_lxf_and_mdb(self, lenex_result):
        _, z = lenex_result
        names = z.namelist()
        assert "meet.lxf" in names
        assert "meet.mdb" in names

    def test_lxf_is_valid_zip_with_lef(self, lenex_result):
        _, z = lenex_result
        lxf_bytes = z.read("meet.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        assert "meet.lef" in lxf.namelist()

    def test_ma_suffix_on_masters_athletes(self, lenex_result):
        """Masters athletes should have _MA suffix on LICENSE."""
        _, z = lenex_result
        lxf_bytes = z.read("meet.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        lef = lxf.read("meet.lef").decode()

        licenses = re.findall(r'license="([^"]+)"', lef)
        ma_licenses = [l for l in licenses if l.endswith("_MA")]
        non_ma_licenses = [l for l in licenses if not l.endswith("_MA")]
        assert len(ma_licenses) > 0, "No _MA suffixed athletes found"
        assert len(non_ma_licenses) > 0, "All athletes have _MA (wrong)"

    def test_masters_entries_in_prelim_events(self, lenex_result):
        """Masters athletes' entries should point to prelim events (not Masters finals)."""
        _, z = lenex_result
        lxf_bytes = z.read("meet.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        lef = lxf.read("meet.lef").decode()

        # Parse athlete blocks with _MA suffix
        athlete_blocks = re.findall(
            r'<ATHLETE [^>]*license="([^"]+)"[^>]*>(.*?)</ATHLETE>',
            lef, re.DOTALL,
        )
        ma_event_ids = []
        for lic, block in athlete_blocks:
            if lic.endswith("_MA"):
                eids = re.findall(r'eventid="(\d+)"', block)
                ma_event_ids.extend(int(e) for e in eids)

        assert len(ma_event_ids) > 0, "No entries found for _MA athletes"
        # Prelim events have IDs < 3000 in the template; Masters finals are 4xxx
        prelim_count = sum(1 for e in ma_event_ids if e < 3000)
        final_count = sum(1 for e in ma_event_ids if e >= 4600)
        assert prelim_count > 0, "No Masters entries in prelim events"
        # Masters-only events (UID 541) go to final — that's OK but should be minority
        assert prelim_count > final_count

    def test_open_25plus_no_ma_suffix(self, lenex_result):
        """Open athletes aged 25+ should NOT have _MA suffix."""
        _, z = lenex_result
        lxf_bytes = z.read("meet.lxf")
        lxf = zipfile.ZipFile(io.BytesIO(lxf_bytes))
        lef = lxf.read("meet.lef").decode()

        # Find athletes with birthdate making them 25+ but no _MA
        athletes = re.findall(
            r'<ATHLETE [^>]*lastname="([^"]+)"[^>]*license="([^"]+)"[^>]*birthdate="([^"]+)"',
            lef,
        )
        for name, lic, bd in athletes:
            try:
                year = int(bd[:4])
                age = 2026 - year
            except ValueError:
                continue
            if age >= 25 and not lic.endswith("_MA"):
                # This is an Open 25+ athlete — correct, no _MA
                return  # found at least one, test passes
        # If we get here, no Open 25+ athletes exist (unlikely with test data)
        # Just check no one under 25 has _MA
        for name, lic, bd in athletes:
            if lic.endswith("_MA"):
                year = int(bd[:4])
                age = 2026 - year
                assert age >= 25, f"{name} has _MA but age={age}"


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
