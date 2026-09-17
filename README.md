# ado_rest_rools

Windows Forms PowerShell app for Azure DevOps Server REST APIs (branch policies, repos, and related collection/project operations).

(The repository name is spelled `ado_rest_rools`.)

## Files

| File | Purpose |
| --- | --- |
| `ado_rest_tools.ps1` | Current WinForms tool. |
| `ado_rest_tools.ps1.bak` | Previous snapshot. |
| `ado_rest_tools.ps1.org` | Earlier snapshot. |
| `ADO_REST_Tools_Result_YYYY-MM-DD.log` | Daily run logs written next to the script. |

## Run

```powershell
.\ado_rest_tools.ps1
```

Set these in the script before use:

- `$ADOServerFQDN`
- `$collection` (default `DevOpsCollection`)
- `$MyPat` (PAT; used as HTTP Basic `:token`)

The UI lets you pick a project and repo. Policy type IDs used in the script include:

- Minimum number of reviewers: `fa4e907d-c16b-4a4c-9dfa-4906e5d171dd`
- Required reviewers: `fd2167ab-b0be-447a-8ec8-39368250530e`

Do not commit a live PAT. Keep `$MyPat` empty in git and fill it locally.
