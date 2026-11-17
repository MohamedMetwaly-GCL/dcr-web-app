# DCR security regressions

Run each suite in its own Python process using the test-only runner. The Phase 1 suite captures SQL and needs no PostgreSQL server. A1, A2, B, the four-findings suite and the five-findings suite require a new PostgreSQL cluster for each invocation. None import the Full Preview launcher.

## Requirements

- An approved private Windows venv outside the protected DCR checkouts. For target validation use Python 3.11.9 and Flask in the declared >=2.3,<3.0 range, with its resolved dependencies.
- Application imports need bcrypt, psycopg2-binary and requests. The four-findings export/import checks also require openpyxl, ReportLab, Pillow, arabic-reshaper and python-bidi. Install from the reviewed application dependency set in the approved environment. The runner supplies import-time Gemini declarations and blocks live SDK clients.
- The five database suites need psutil and a reviewed Windows PostgreSQL binary distribution containing initdb.exe, postgres.exe and pg_ctl.exe, with its original DLLs and share files. The pixeltable_pgserver Python wrapper is unnecessary. The database pool override uses plaintext loopback within this disposable cluster, so these suites do not validate deployment TLS.
- The last two suites invoke Node.js to verify rendered JavaScript. Install it in its standard Windows directory or ensure it is resolvable in the runner's sanitized process. A skipped Node test makes the runner exit nonzero.

## Run

Use an approved offline Windows session with loopback retained. Disconnect external adapters, VPNs and tunnels and keep them disconnected until the subprocesses exit. The runner checks interface and route state before database operations. Its Python socket audit hook is additional protection and does not enforce native egress for libpq, PostgreSQL, or Node. The network inventory is point-in-time. Leave testing pending if the offline boundary cannot be maintained. The runner does not alter host networking.

From the integration worktree, point $python and $pgBin at the approved runtime and reviewed PostgreSQL distribution:

    $python = '<approved-private-venv>\Scripts\python.exe'
    $pgBin = '<reviewed-PostgreSQL-distribution>\bin'
    $suites = @(
        'test_security_phase1_controls',
        'test_finding_a1_cross_workspace',
        'test_finding_a2_project_ownership',
        'test_finding_b_record_ownership',
        'test_four_findings_regression',
        'test_five_findings_regression'
    )
    foreach ($suite in $suites) {
        $pgArgs = @()
        if ($suite -ne 'test_security_phase1_controls') { $pgArgs = @('--pg-bin', $pgBin) }
        & $python -I -S -B .\security_test_support.py $suite @pgArgs
        if ($LASTEXITCODE -ne 0) { throw "Suite failed or was blocked: $suite" }
    }

The runner clears inherited credentials and database configuration and uses synthetic app keys. It imports the Flask app for its in-process test client, which leaves the server and polling entry point outside the run. It creates a unique system-temporary directory and starts PostgreSQL on a selected loopback port with a generated password and SCRAM. Before schema setup, it checks the owned executable, listener, data path, server address and port, database, role, and empty application schema, and checks rejection of an incorrect password. Cleanup stops the owned process and leaves only the new synthetic cluster and its log at the printed temporary path. The preserved preview cluster is outside the run.
