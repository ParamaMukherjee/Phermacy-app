# MediBill Pro v20 — automated validation report

Date: 2026-11-21

## Passed

- `python -m py_compile main.py license_client.py licensing_server/app.py licensing_server/owner_cli.py test_licensing.py`
- Existing v19 headless database/login-owner invariants, PDF generation/print-path smoke test, Excel export smoke tests, and checkout transaction/customer auto-save/stock deduction/PDF flow: PASS.
- New licensing API tests: device trial activation, status token issuance, payment reference intake, owner-only pending payment list, owner lifetime approval, AMC renewal, unauthorized owner rejection: PASS.
- Ed25519 API signature verification, desktop-client signature verification, and tamper rejection: PASS.

## Not tested / release blockers

- No Windows EXE was built in this Linux environment. Run the GitHub Actions workflow on Windows.
- No Windows GUI launch or physical printer test was performed.
- No hosted licensing service exists yet. Configure keys, persistent database, owner secret, HTTPS URL, and public key before building the customer release.
- No real UPI payment was made or verified. UTR submission is only a request; manual owner verification is required.
- Offline grace, device migration, hosting restart/persistence, and real expiry were not tested against a deployed service.

The source package is a deployment candidate, not a production-certified executable.
