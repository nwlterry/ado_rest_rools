# ado_rest_rools

Windows Forms PowerShell app for Azure DevOps Server REST APIs (branch policies, repos, collection/project operations).

Compatible with **Azure DevOps Server 2020 Update 2 through 25H2 (20260506.1)** and later.

(The repository name is spelled `ado_rest_rools`.)

## Layout

```
scripts/ado_rest_tools.ps1   # current tool (rewritten)
configs/                     # environment config files (*.json)
backups/                     # automatic backups created before updates (created at runtime)
archive/                     # .bak / .org snapshots
logs/                        # daily run logs
.vscode/
GROUP.md
README.md
```

## Features

- Multi-environment support via JSON config files
- Configurable or auto-detected REST API version (6.0 / 7.0 / 7.1 / 7.2)
- Automatic backup of policies and repository state before any modifying operation
- Form-based config selection and editor

## Run

```powershell
.\scripts\ado_rest_tools.ps1
```

1. Create or edit a config under `configs/` (or use the **Edit / New Config** button in the form).
2. Select the environment in the form and click **Load Config**.
3. Refresh projects and run queries or updates.

Do not commit a live PAT. Sample configs are provided with empty PAT fields.

---

See [GROUP.md](GROUP.md) for sibling repositories. Catalog: https://github.com/nwlterry/nwlterry
