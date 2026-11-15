"""Run one DCR regression suite in a fresh, manually offline Windows session."""

import argparse
from contextlib import ExitStack, contextmanager
from pathlib import Path
import os
import secrets
import socket
import subprocess
import sys
import tempfile
import time
from types import ModuleType, SimpleNamespace
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parent
PHASE1 = "test_security_phase1_controls"
SUITES = (
    PHASE1,
    "test_finding_a1_cross_workspace",
    "test_finding_a2_project_ownership",
    "test_finding_b_record_ownership",
    "test_four_findings_regression",
    "test_five_findings_regression",
)
HOST = "127.0.0.1"
INITDB_ONLY = "initdb-only"
_prepared = False
_pg_bin = None
_run_dir = None
_allowed_port = None


def _blocked(*args, **kwargs):
    raise RuntimeError("External service access is disabled in DCR security tests")


_WINDOWS_OFFLINE_PROBE = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $interfaces = @(Get-NetIPInterface -IncludeAllCompartments)
    $adapters = @(Get-NetAdapter -IncludeHidden)
    $routes = @(Get-NetRoute -IncludeAllCompartments)
    if (!$adapters.Count -or !$routes.Count -or !($interfaces | Where-Object {
        $_.InterfaceIndex -eq 1 -and $_.AddressFamily -eq 'IPv4' -and $_.ConnectionState -eq 'Connected'
    })) { throw 'Incomplete network inventory or connected IPv4 loopback unavailable' }

    $miniports = @()
    $blocked = @(
        foreach ($adapter in $adapters) {
            if ($adapter.Status -eq 'Up' -and $adapter.HardwareInterface -is [bool] -and
                !$adapter.HardwareInterface -and $adapter.ifIndex -gt 1 -and
                $adapter.ComponentID -in 'ms_ndiswanip','ms_ndiswanipv6','ms_ndiswanbh') {
                $miniports += $adapter.ifIndex
            } elseif ($adapter.Status -notin 'Disconnected','Disabled','Not Present') {
                "Adapter '$($adapter.Name)' ($($adapter.InterfaceDescription), ifIndex $($adapter.ifIndex)): status '$($adapter.Status)'"
            }
        }
        foreach ($interface in $interfaces) {
            if ($interface.InterfaceIndex -ne 1 -and $interface.ConnectionState -ne 'Disconnected') {
                "IP interface '$($interface.InterfaceAlias)' ($($interface.AddressFamily), ifIndex $($interface.InterfaceIndex), compartment $($interface.CompartmentId)): state '$($interface.ConnectionState)'"
            }
        }
        foreach ($route in $routes) {
            if (!$route.DestinationPrefix -or $route.DestinationPrefix -in '0.0.0.0/0','::/0' -or
                $route.InterfaceIndex -in $miniports) {
                "Route '$($route.DestinationPrefix)' via '$($route.InterfaceAlias)' (ifIndex $($route.InterfaceIndex), compartment $($route.CompartmentId), next hop $($route.NextHop))"
            }
        }
    )
    if ($blocked.Count) { $blocked | Write-Output; exit 2 }
    Write-Output 'Offline preflight passed'
    exit 0
} catch {
    Write-Output "Network inventory failed: $($_.Exception.Message)"
    exit 3
}
"""


def _require_offline_windows():
    if os.name != "nt":
        raise RuntimeError("This runner requires the Windows offline preflight")
    powershell = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    result = subprocess.run(
        [str(powershell), "-NoProfile", "-NonInteractive", "-Command", _WINDOWS_OFFLINE_PROBE],
        env={"SystemRoot": os.environ["SystemRoot"], "WINDIR": os.environ["SystemRoot"]},
        capture_output=True, timeout=20,
    )
    output = result.stdout.decode("utf-8", errors="replace").strip()
    errors = result.stderr.decode("utf-8", errors="replace").strip()
    if result.returncode or errors or output != "Offline preflight passed":
        details = "\n".join(filter(None, (output, errors))) or f"PowerShell exit {result.returncode}"
        raise RuntimeError(f"Offline Windows boundary unverified or external connectivity present\n{details}")


def _guard_sockets(event, args):
    if event in {"socket.gethostbyname", "socket.gethostbyaddr", "socket.getnameinfo"}:
        _blocked()
    if event == "socket.getaddrinfo" and args[0] != HOST:
        _blocked()
    if event in {"socket.connect", "socket.sendto", "socket.sendmsg"}:
        if _allowed_port is None or args[1] != (HOST, _allowed_port):
            _blocked()
    if event == "socket.bind" and args[1] != (HOST, 0):
        _blocked()


def _install_sdk_stubs():
    """Expose only the declarations the AI blueprint builds at import time."""
    google = ModuleType("google")
    google.__path__ = []
    genai = ModuleType("google.genai")
    types = ModuleType("google.genai.types")
    types.Type = SimpleNamespace(**{name: name for name in
                                    ("OBJECT", "STRING", "BOOLEAN", "INTEGER", "NUMBER", "ARRAY")})
    for name in ("FunctionDeclaration", "Schema", "Tool", "Content", "Part", "GenerateContentConfig"):
        setattr(types, name, SimpleNamespace)
    genai.Client = _blocked
    genai.types = types
    google.genai = genai
    drive_service = ModuleType("drive_service")
    for name in ("get_drive_service", "get_drive_write_service", "process_drive_folder", "register_webhook"):
        setattr(drive_service, name, _blocked)
    sys.modules.update({"google": google, "google.genai": genai, "google.genai.types": types,
                        "drive_service": drive_service})


def _postgres_environment():
    windows = Path(os.environ["SystemRoot"]).resolve(strict=True)
    system32 = windows / "System32"
    cmd = system32 / "cmd.exe"
    if not cmd.is_file():
        raise RuntimeError("Windows command processor unavailable")
    return {
        "SystemRoot": str(windows), "WINDIR": str(windows),
        "COMSPEC": str(cmd), "PATH": os.pathsep.join((str(_pg_bin), str(system32))),
        "TEMP": str(_run_dir), "TMP": str(_run_dir),
    }


def _initialize_cluster(pgdata, password, log, process_options):
    _require_offline_windows()
    if pgdata.exists():
        raise RuntimeError("Refusing to initialize an existing PostgreSQL data directory")
    pwfile = pgdata.with_name("init-password")
    with pwfile.open("x", encoding="ascii") as stream:
        stream.write(password + "\n")
    try:
        result = subprocess.run(
            [str(_pg_bin / "initdb.exe"), "-D", str(pgdata),
             "--auth=scram-sha-256", "--username=dcr_test", f"--pwfile={pwfile}",
             "--encoding=UTF8", "--no-locale"],
            stdout=log, stderr=subprocess.STDOUT, timeout=90, **process_options,
        )
    finally:
        pwfile.unlink()
    if result.returncode:
        raise RuntimeError(f"initdb exited with {result.returncode & 0xFFFFFFFF:#010x}. See {log.name}")
    if not (pgdata / "PG_VERSION").is_file():
        raise RuntimeError(f"initdb reported success without creating PG_VERSION. See {log.name}")


@contextmanager
def _running_postgres(pgdata, port, log, process_options):
    import psutil

    pidfile = pgdata / "postmaster.pid"
    pg_ctl = str(_pg_bin / "pg_ctl.exe")
    executable = _pg_bin / "postgres.exe"
    server_log = _run_dir / "postgres-server.log"
    if pidfile.exists():
        raise RuntimeError("Refusing to start PostgreSQL with an existing postmaster.pid")
    started = time.time()
    owned = None

    def identify_server():
        try:
            lines = pidfile.read_text(encoding="utf-8").splitlines()
            pid = int(lines[0])
            if (pid <= 0 or Path(lines[1]).resolve() != pgdata.resolve()
                    or int(lines[3]) != port):
                raise RuntimeError(f"Unexpected PostgreSQL PID file identity: {pidfile}")
            process = psutil.Process(pid)
            args = process.cmdline()
            if (process.create_time() < started - 2 or Path(process.exe()).resolve() != executable
                    or "-D" not in args or Path(args[args.index("-D") + 1]).resolve() != pgdata.resolve()
                    or not process.is_running()):
                raise RuntimeError(f"Cannot verify PostgreSQL process {pid} for {pgdata}")
            return process
        except (FileNotFoundError, psutil.NoSuchProcess):
            return None
        except (IndexError, ValueError) as exc:
            raise RuntimeError(f"Incomplete or invalid PostgreSQL PID or command line: {pidfile}") from exc

    def control(*args, timeout):
        result = subprocess.run(
            [pg_ctl, *args], stdout=log, stderr=subprocess.STDOUT, timeout=timeout, **process_options,
        )
        if result.returncode:
            raise RuntimeError(f"pg_ctl {args[0]} exited with {result.returncode & 0xFFFFFFFF:#010x}. "
                               f"See {log.name} and {server_log}")

    try:
        control("start", "-D", str(pgdata), "-p", str(executable), "-l", str(server_log),
                "-o", f"-h {HOST} -p {port} -c ssl=off", "-w", "-t", "30", timeout=40)
        owned = identify_server()
        if owned is None:
            raise RuntimeError(f"PostgreSQL started without a verifiable server PID. See {server_log}")
        listeners = {(c.laddr.ip, c.laddr.port) for c in owned.net_connections(kind="tcp")
                     if c.status == psutil.CONN_LISTEN}
        if listeners != {(HOST, port)}:
            raise RuntimeError(f"Unexpected listeners for PostgreSQL PID {owned.pid}: {listeners}")
        yield
    finally:
        # pg_ctl can time out while its server is still starting. Look for
        # that server's PID file before deciding whether it is safe to stop.
        current = identify_server()
        if owned is None and current is None:
            deadline = time.monotonic() + 5
            while current is None and time.monotonic() < deadline:
                time.sleep(0.1)
                current = identify_server()
            if current is None:
                log.write(b"No live server PID verified after startup failure; no stop signal sent.\n")
        if current is not None:
            if owned is not None and current != owned:
                raise RuntimeError(f"PostgreSQL PID identity changed; refusing to stop {current.pid}")
            control("stop", "-D", str(pgdata), "-m", "fast", "-w", "-t", "10", timeout=15)
            current.wait(timeout=10)
        elif owned is not None and owned.is_running():
            raise RuntimeError(f"PostgreSQL PID file is missing; refusing to stop PID {owned.pid}")


@contextmanager
def disposable_database():
    """Start a new authenticated cluster and prove ownership before DCR writes."""
    global _allowed_port
    if not _prepared or _pg_bin is None or "db" in sys.modules or "app" in sys.modules:
        raise RuntimeError("Use the dedicated runner with a fresh interpreter and PostgreSQL binaries")
    _require_offline_windows()
    import psycopg2
    from psycopg2.pool import ThreadedConnectionPool

    with socket.socket() as probe:
        probe.bind((HOST, 0))
        port = probe.getsockname()[1]
    pgdata = _run_dir / "pgdata"
    password = secrets.token_hex(24)
    process_options = {"env": _postgres_environment(), "cwd": _run_dir}
    with (_run_dir / "postgres.log").open("xb") as log:
        _initialize_cluster(pgdata, password, log, process_options)
        _require_offline_windows()
        with _running_postgres(pgdata, port, log, process_options):
            try:
                _allowed_port = port
                dsn = dict(hostaddr=HOST, host=HOST, port=port, dbname="postgres", user="dcr_test",
                           password=password, sslmode="disable", connect_timeout=3)
                try:
                    wrong = psycopg2.connect(**dict(dsn, password=secrets.token_hex(24)))
                except psycopg2.OperationalError as exc:
                    if "password authentication failed" not in str(exc):
                        raise RuntimeError("Could not verify PostgreSQL password rejection") from None
                else:
                    wrong.close()
                    raise RuntimeError("PostgreSQL unexpectedly accepted an incorrect password")

                with ExitStack() as cleanup:
                    pool = ThreadedConnectionPool(1, 16, **dsn)
                    cleanup.callback(pool.closeall)
                    conn = pool.getconn()
                    try:
                        with conn, conn.cursor() as cursor:
                            cursor.execute(
                                "SELECT current_setting('data_directory'), host(inet_server_addr()), "
                                "inet_server_port(), current_database(), current_user, "
                                "(SELECT count(*) FROM information_schema.tables WHERE table_schema='public')"
                            )
                            directory, addr, actual_port, database, role, tables = cursor.fetchone()
                        if (Path(directory).resolve() != pgdata.resolve()
                                or (addr, actual_port, database, role, tables) != (HOST, port, "postgres", "dcr_test", 0)):
                            raise RuntimeError("Connected database is not the verified fresh test cluster")
                    finally:
                        pool.putconn(conn)

                    import db
                    import drive_upload
                    cleanup.enter_context(patch.object(db, "get_pool", return_value=pool))
                    cleanup.enter_context(patch.object(db, "enqueue_notification", side_effect=_blocked))
                    cleanup.enter_context(patch.object(drive_upload, "get_upload_credentials", side_effect=_blocked))
                    cleanup.enter_context(patch.object(drive_upload, "access_token", side_effect=_blocked))
                    db.init()
                    yield db
            finally:
                _allowed_port = None


def main():
    global _prepared, _pg_bin, _run_dir
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("suite", choices=(*SUITES, INITDB_ONLY))
    parser.add_argument("--pg-bin", type=Path, help="Directory containing reviewed PostgreSQL binaries")
    args = parser.parse_args()
    if not (sys.flags.isolated and sys.flags.no_site and sys.dont_write_bytecode):
        parser.error("launch the private interpreter with -I -S -B")
    venv = Path(sys.executable).parent.parent.resolve()
    config = venv / "pyvenv.cfg"
    if not config.is_file() or "include-system-site-packages = false" not in config.read_text().lower():
        parser.error("use a private venv that excludes system packages")
    protected = (
        Path("D:/DCR").resolve(),
        (Path(os.environ["LOCALAPPDATA"]) / "DCR_DEV_ISOLATED").resolve(),
    )
    if any(ROOT.is_relative_to(path) for path in protected):
        parser.error("run from the integration worktree, outside the protected checkouts")
    if any(venv.is_relative_to(path) for path in protected):
        parser.error("use a private runtime outside the protected checkouts")
    if any(ROOT.glob(".env*")) or (ROOT / "service_account.json").exists():
        parser.error("credential-like file in the test checkout")
    _require_offline_windows()
    if args.suite != PHASE1:
        if args.pg_bin is None:
            parser.error("database suites require --pg-bin")
        _pg_bin = args.pg_bin.resolve(strict=True)
        if any(_pg_bin.is_relative_to(path) for path in protected):
            parser.error("use PostgreSQL binaries outside the protected checkouts")
        names = ("initdb.exe", "postgres.exe", "pg_ctl.exe")
        if not all((_pg_bin / name).is_file() and (_pg_bin / name).resolve().parent == _pg_bin
                   for name in names):
            parser.error("PostgreSQL binary directory is incomplete")
    temp_root = (Path(os.environ["LOCALAPPDATA"]) / "Temp").resolve(strict=True)
    if any(temp_root.is_relative_to(path) for path in protected):
        parser.error("temporary directory overlaps a protected checkout")
    _run_dir = Path(tempfile.mkdtemp(prefix=f"dcr-{args.suite}-", dir=temp_root))
    environ = {key: os.environ[key] for key in ("SystemRoot", "WINDIR") if key in os.environ}
    environ.update(TEMP=str(_run_dir), TMP=str(_run_dir), USERPROFILE=str(_run_dir),
                   LOCALAPPDATA=str(_run_dir), APPDATA=str(_run_dir), HOME=str(_run_dir),
                   SECRET_KEY=secrets.token_hex(32), MAGIC_SECRET=secrets.token_hex(32),
                   GEMINI_API_KEY="synthetic-test-sdk-stub")
    os.environ.clear()
    os.environ.update(environ)
    os.chdir(_run_dir)
    sys.path[:0] = [str(ROOT), str(venv / "Lib/site-packages")]
    sys.addaudithook(_guard_sockets)
    _install_sdk_stubs()
    _prepared = True
    # Keep __main__ and imports from the suites on the same runner state.
    sys.modules["security_test_support"] = sys.modules[__name__]
    print(f"Disposable test output: {_run_dir}", flush=True)
    if args.suite == INITDB_ONLY:
        with (_run_dir / "postgres.log").open("xb") as log:
            with tempfile.TemporaryDirectory(prefix="cluster-", dir=_run_dir) as cluster:
                _initialize_cluster(
                    Path(cluster) / "pgdata", secrets.token_hex(24), log,
                    {"env": _postgres_environment(), "cwd": _run_dir},
                )
        print("PostgreSQL initialization and disposable data cleanup passed. No server started.")
        return 0
    suite = unittest.defaultTestLoader.loadTestsFromName(args.suite)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    return 0 if result.wasSuccessful() and not result.skipped else 1


if __name__ == "__main__":
    raise SystemExit(main())
