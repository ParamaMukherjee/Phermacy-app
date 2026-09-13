MEDIBILL PRO — GET THE FINISHED WINDOWS EXE WITHOUT INSTALLING PYTHON

You were clear: neither you nor the shop owner should have to install Python.

This package is set up so Windows builds the EXE for you in GitHub's Windows
build machine. Your own PC does NOT need Python.

FASTEST METHOD
==============

1. Create a free GitHub account if you don't already have one.

2. Create a new repository, e.g.:
       medibill-pro

3. Upload ALL files/folders from this package into that repository.
   IMPORTANT: keep:
       .github/workflows/build-windows.yml

4. In GitHub, open:
       Actions
   then select:
       Build MediBill Windows EXE

5. Click:
       Run workflow

6. Wait for the green checkmark.

7. Open the completed workflow run. At the bottom, under Artifacts,
   download:
       MediBill-Windows

8. Extract it. You will have:
       MediBill.exe

9. Send ONLY MediBill.exe to the shop owner.

SHOP OWNER
==========

The shop owner:
    double-clicks MediBill.exe
    -> MediBill opens.

They do NOT need:
    - Python
    - pip
    - Node.js
    - a browser
    - a backend
    - a server
    - command prompt
    - separate frontend

DATA
====

SQLite is local. The app creates its database automatically under the user's
MediBillPro data folder.

DEMO LOGIN
==========

Admin:
  admin@medibillwb.in
  Admin@1234

Store Manager:
  store@medibillwb.in
  Demo@1234

Cashier:
  cashier@medibillwb.in
  Demo@1234

Accounts:
  accounts@medibillwb.in
  Demo@1234

SMS
===

Actual SMS delivery still needs a transport:
- an SMS provider over the internet, or
- a GSM/SIM modem.

This does NOT require a separate backend: the desktop application can make
the provider request itself. If no transport is configured, the local SMS
outbox retains the message.

IMPORTANT
=========

The Windows EXE must be built by a Windows runner. This environment cannot
compile a genuine Windows executable, so the included GitHub workflow performs
that build automatically without installing Python on your PC.
