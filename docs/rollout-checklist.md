# Rollout checklist

## Pre-pilot
- [ ] Confirm the account has Cortex and CoCo enabled, and that trial accounts have a payment method added.
- [ ] Run `sql/01` through `sql/05` as ACCOUNTADMIN. Confirm `validation/validate.sql` passes.
- [ ] Confirm the OAuth integration's `ALLOWED_ROLES_LIST` (if used) includes only `COCO_SHEETS_ANALYST`.
- [ ] Set per-user daily CoCo credit limits.
- [ ] Share the OAuth client ID through a channel your org approves. Do not share the client secret.
- [ ] Confirm the pilot user's sign-in role is `COCO_SHEETS_ANALYST`, not ACCOUNTADMIN (OAuth blocks ACCOUNTADMIN by default).
- [ ] Confirm the Sheets install has the Code.gs and appsscript.json copied from the browser, since the add-on files are not fetchable from every network.

## Pilot
- [ ] Grant `COCO_SHEETS_ANALYST` to pilot users only.
- [ ] Each pilot user signs in with their own credentials and confirms their row and masking results.
- [ ] Test a refresh on one table, and confirm the preview before accepting.
- [ ] Confirm citations point back to the expected query.
- [ ] Review `sql/06_usage_monitoring.sql` output after one week. If it returns 0 rows, confirm that CoCo requests have run and that `ACCOUNT_USAGE` latency has passed before troubleshooting.
- [ ] Run a test prompt in the Sheets pane, then confirm the usage queries show that request.

## Go-live decision
- [ ] Preview status reviewed against current Snowflake release notes.
- [ ] Excel-only features (CORTEX cell functions, refresh on open) are not required by any Sheets users.
- [ ] Data owners approve the semantic view and policy scope.
- [ ] Incident owner and rollback steps (drop the role grant, disable the integration) are documented.
