# REFACTOR README

This branch (fix/refactor-core-2026) contains a core refactor to consolidate DB and printing state, remove shadowed builtins, and add a safe printing fallback.

What changed
- globals.bas: removed custom IsNumeric, made ShortGuid produce 6-digit IDs, improved session helpers
- monitor_service.bas: validation improvements
- printer_service.bas: now a facade; uses receipt_service for state and print_helper_client for I/O
- db_service.bas: added db_SetPrintAttemptStarted
- attendance_processor.bas: refactored to use db_service APIs
- print_helper_client.bas: fallback print (file -> notepad)
- receipt_service.bas: centralized print datetime update

How to test
- Open the Access project and run `tests.run_e2e_test` (or call `run_e2e_test` from the Immediate window).
- Alternatively follow the PR_READY.md steps.

Next steps
- Optional: implement Winsock socket printing in print_helper_client
- Clean up UI forms to remove business logic
- Add more automated tests
