---
name: titan-tokos-dashboard-updater
description: Update and publish the Titan/Tokos advertising dashboard from the shared Google Sheets reports. Use when Codex needs to refresh the dashboard data, compare Monday report data, rebuild titan-tokos-dashboard.html, commit changes, and push the GitHub Pages site for a1reklamaru/titantokos.
---

# Titan/Tokos Dashboard Updater

Use the project script instead of hand-editing embedded dashboard data.

## Workflow

1. Run from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\update-dashboard.ps1
```

2. The script downloads:

- Titan CSV: Google Sheets document `1anw9pGO6y7emBZlbvgUJUycLtC54KMz5d-i3vXmnji8`, gid `0`
- Tokos CSV: same document, gid `1190677190`

3. The script parses weekly campaign blocks, rebuilds:

- `clients.<client>.weekly`
- `clients.<client>.monthly`
- `clients.<client>.campaignsWeekly`
- `clients.<client>.campaignsMonthly`

4. It replaces only the `const clients = ...` block in `titan-tokos-dashboard.html`.

5. If data changed, it commits as `Update dashboard data` and pushes to GitHub Pages.

## Validation

After running, check:

```powershell
git status --short --branch
git log --oneline -1
```

If the script reports no changes, do not create a commit.

If Google Sheets download fails, retry once. If it still fails, report the network/access error.
