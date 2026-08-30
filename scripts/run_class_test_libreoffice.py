#!/usr/bin/env python3
"""
Run the VBA class-module support probe against LibreOffice.

Starts a headless soffice listening on a UNO socket, runs uno_class_probe.py
under an interpreter that can import `uno`, and prints a support matrix.

Exit code is 0 when the probe ran to completion -- individual VBA features
failing is a *result*, not a script error. Use --strict to fail the process
when any test does not PASS.
"""

import argparse
import json
import os
import shutil
import socket
import subprocess
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
DEFAULT_PORT = 2002

GROUP_LABELS = {
    "basic": "Core class features",
    "lifecycle": "Class_Initialize / Class_Terminate",
    "predeclared": "VB_PredeclaredId (default instance)",
    "implements": "Implements (interfaces)",
    "events": "Public Event / RaiseEvent / WithEvents",
}


def find_libreoffice() -> str | None:
    """Locate the soffice binary on Windows, macOS, or Linux."""
    candidates = {
        "win32": [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ],
        "darwin": ["/Applications/LibreOffice.app/Contents/MacOS/soffice"],
    }.get(sys.platform, ["/usr/bin/soffice", "/usr/bin/libreoffice"])

    for path in candidates:
        if Path(path).exists():
            return path

    return shutil.which("soffice") or shutil.which("libreoffice")


def _can_import_uno(python_exe: str, extra_env: dict[str, str]) -> bool:
    """Actually try `import uno` -- guessing from paths is not reliable enough."""
    env = os.environ.copy()
    env.update(extra_env)
    tag = f"{python_exe}{' +env' if extra_env else ''}"
    try:
        done = subprocess.run(
            [python_exe, "-c", "import uno"],
            capture_output=True,
            text=True,
            timeout=30,
            env=env,
        )
    except subprocess.TimeoutExpired:
        print(f"  uno check: {tag} -> timed out")
        return False
    except OSError as exc:
        print(f"  uno check: {tag} -> {type(exc).__name__}")
        return False

    if done.returncode == 0:
        print(f"  uno check: {tag} -> OK")
        return True

    reason = (done.stderr or "").strip().splitlines()
    print(f"  uno check: {tag} -> {reason[-1][:120] if reason else 'failed'}")
    return False


def find_uno_python(soffice: str) -> tuple[str, dict[str, str]] | tuple[None, None]:
    """
    Find an interpreter that can `import uno`, plus the env it needs.

    LibreOffice ships its own Python on Windows and macOS. On Linux the distro
    package (python3-uno) wires uno into the *system* interpreter -- which
    matters in CI, where actions/setup-python puts a different python3 first on
    PATH that has no uno. So every candidate is verified by running it.
    """
    program_dir = Path(soffice).resolve().parent

    # LibreOffice's own bundled interpreters need no extra environment.
    candidates: list[tuple[str, dict[str, str]]] = [
        (str(p), {})
        for p in (
            program_dir / "python.exe",
            program_dir / "python",
            program_dir.parent / "Resources" / "python",
        )
        if p.exists()
    ]

    # Debian/Ubuntu layout: uno.py sits beside soffice and the URE needs
    # bootstrapping when imported from an interpreter outside LibreOffice.
    linux_env: dict[str, str] = {}
    for lib_dir in ("/usr/lib/libreoffice/program", str(program_dir)):
        if Path(lib_dir, "uno.py").exists():
            linux_env = {
                "PYTHONPATH": lib_dir,
                "URE_BOOTSTRAP": f"file://{lib_dir}/fundamentalrc",
            }
            break

    for exe in ("/usr/bin/python3", shutil.which("python3"), sys.executable):
        if not exe:
            continue
        candidates.append((exe, linux_env))
        if linux_env:
            candidates.append((exe, {}))

    for exe, env in candidates:
        if _can_import_uno(exe, env):
            return exe, env

    return None, None


def wait_for_port(port: int, timeout: float) -> bool:
    """Poll the UNO socket until soffice accepts connections."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.5)
            if sock.connect_ex(("127.0.0.1", port)) == 0:
                return True
        time.sleep(0.5)
    return False


def start_soffice(soffice: str, port: int, profile_dir: Path) -> subprocess.Popen:
    """Launch headless soffice with an isolated user profile."""
    profile_dir.mkdir(parents=True, exist_ok=True)
    # Must be a real file URL. Naive f"file://{path}" yields "file://C:\..." on
    # Windows, which soffice rejects, and it then never opens the UNO socket.
    profile_url = profile_dir.resolve().as_uri()
    cmd = [
        soffice,
        f"-env:UserInstallation={profile_url}",
        "--headless",
        "--invisible",
        "--nologo",
        "--nofirststartwizard",
        "--norestore",
        "--nodefault",
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ServiceManager",
    ]
    print(f"Starting: {' '.join(cmd)}")
    return subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def print_matrix(data: dict) -> tuple[int, int]:
    """Render the probe output as a grouped support matrix. Returns (passed, total)."""
    results = data.get("results", [])
    if not results:
        print("\nNo test results were produced.")
        return 0, 0

    symbols = {
        "PASS": "PASS",
        "FAIL": "FAIL",
        "ERROR": "ERR ",
        "NOT_FOUND": "N/A ",
        "NO_RESULT": "----",
        "UNKNOWN": "??? ",
    }
    passed = 0

    for group, label in GROUP_LABELS.items():
        rows = [r for r in results if r["group"] == group]
        if not rows:
            continue
        print(f"\n  {label}")
        for row in rows:
            status = row["status"]
            if status == "PASS":
                passed += 1
            detail = row["detail"]
            if len(detail) > 150:
                detail = detail[:150] + "..."
            print(f"    [{symbols.get(status, status)}] {row['function']}")
            print(f"           {detail}")

    return passed, len(results)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workbook",
        default=str(PROJECT_ROOT / "output" / "class_test_workbook.xlsm"),
        help="workbook to probe (default: output/class_test_workbook.xlsm)",
    )
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument(
        "--startup-timeout", type=float, default=90.0,
        help="seconds to wait for soffice to accept connections",
    )
    parser.add_argument(
        "--probe-timeout", type=float, default=180.0,
        help="seconds to wait for the UNO probe to finish",
    )
    parser.add_argument(
        "--strict", action="store_true",
        help="exit non-zero unless every test PASSes",
    )
    args = parser.parse_args()

    # CI runs this through a shell that is not a tty, so Python block-buffers
    # stdout. A step that then hits the job timeout gets killed with its whole
    # log still in the buffer -- which is indistinguishable from hanging before
    # the first print. Line buffering keeps the progress log truthful.
    sys.stdout.reconfigure(line_buffering=True)

    workbook = Path(args.workbook)
    output_dir = workbook.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("LibreOffice VBA class-module support probe")
    print("=" * 60)
    print(f"Platform: {sys.platform}")

    if not workbook.exists():
        print(f"Workbook not found: {workbook}")
        print("Run scripts/create_class_test_excel.py first.")
        return 1

    soffice = find_libreoffice()
    if not soffice:
        print("LibreOffice not found. Install it, or use: docker compose run --rm class-test")
        return 1
    print(f"soffice: {soffice}")
    try:
        version = subprocess.run(
            [soffice, "--version"], capture_output=True, text=True, timeout=20
        ).stdout.strip()
        print(f"version: {version or '(no output)'}")
    except subprocess.TimeoutExpired:
        # soffice.exe --version is known to hang on Windows in some setups.
        print("version: (--version timed out; continuing)")
    except (OSError, subprocess.SubprocessError):
        print("version: (unavailable)")

    print("Looking for an interpreter that can import uno...")
    python_exe, extra_env = find_uno_python(soffice)
    if python_exe is None:
        print("No interpreter on this machine can 'import uno'.")
        print("Linux: apt install python3-uno. Windows/macOS: use LibreOffice's bundled python.")
        return 1
    print(f"UNO python: {python_exe} {extra_env or ''}")

    probe = Path(__file__).parent / "uno_class_probe.py"
    result_path = output_dir / "class_test_result.json"
    profile_dir = output_dir / ".lo_profile"

    proc = start_soffice(soffice, args.port, profile_dir)
    try:
        print(f"Waiting up to {args.startup_timeout}s for the UNO socket...")
        if not wait_for_port(args.port, args.startup_timeout):
            print(f"soffice did not accept connections within {args.startup_timeout}s")
            # Never plain read() here: soffice does not exit on its own, so the
            # read would block until the CI job timeout kills the whole step.
            proc.terminate()
            try:
                _, stderr = proc.communicate(timeout=15)
            except subprocess.TimeoutExpired:
                proc.kill()
                _, stderr = proc.communicate(timeout=15)
            if stderr:
                print(f"soffice stderr: {stderr.decode(errors='replace')[:800]}")
            return 1
        print(f"soffice listening on port {args.port}")

        env = os.environ.copy()
        env.update(extra_env)

        print(f"Running probe (timeout {args.probe_timeout}s)...")
        try:
            completed = subprocess.run(
                [python_exe, str(probe), str(workbook.absolute()),
                 str(result_path.absolute()), str(args.port)],
                capture_output=True,
                text=True,
                timeout=args.probe_timeout,
                env=env,
            )
        except subprocess.TimeoutExpired:
            print(f"Probe timed out after {args.probe_timeout}s")
            return 1

        if completed.stderr.strip():
            print(f"probe stderr:\n{completed.stderr[:1000]}")

        if not result_path.exists():
            print("Probe produced no result file.")
            print(completed.stdout[:2000])
            return 1
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()
        print("soffice terminated")

    data = json.loads(result_path.read_text(encoding="utf-8"))

    print("\n" + "-" * 60)
    print("Environment")
    print("-" * 60)
    for line in data.get("report", []):
        print(f"  {line}")

    print("\n" + "-" * 60)
    print("VBA class-module support matrix")
    print("-" * 60)
    passed, total = print_matrix(data)

    print("\n" + "=" * 60)
    print(f"Result: {passed}/{total} tests PASSED")
    print(f"Full output: {result_path}")
    print("=" * 60)

    if args.strict and passed != total:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
