# ado_rest_rools

Windows Forms PowerShell app for Azure DevOps Server REST APIs (branch policies, repos, collection/project operations).

(The repository name is spelled `ado_rest_rools`.)

## Layout

```
scripts/ado_rest_tools.ps1   # current tool
archive/                     # .bak / .org snapshots
logs/                        # daily run logs
.vscode/
GROUP.md
README.md
```

## Run

```powershell
.\scripts\ado_rest_tools.ps1
```

Set in the script before use: `$ADOServerFQDN`, `$collection` (default `DevOpsCollection`), `$MyPat` (PAT as HTTP Basic `:token`). Do not commit a live PAT.

---

See [GROUP.md](GROUP.md) for sibling repositories. Catalog: https://github.com/nwlterry/nwlterry
