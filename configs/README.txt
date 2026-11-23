# Sample configuration files for Azure DevOps Server REST Tool
# Place files in a "configs" folder next to the script.
# Copy this template and fill in values. Never commit real PATs.

# Example: configs/prod.json
{
  "Name": "Production",
  "ServerFQDN": "ado.contoso.com",
  "Collection": "DevOpsCollection",
  "PAT": "",
  "ApiVersion": "auto",
  "UseHttps": true,
  "Notes": "ApiVersion can be 'auto', '6.0', '7.0', '7.1', or '7.2'"
}

# Example: configs/dev.json
{
  "Name": "Development",
  "ServerFQDN": "ado-dev.contoso.com",
  "Collection": "DefaultCollection",
  "PAT": "",
  "ApiVersion": "7.2",
  "UseHttps": true,
  "Notes": "25H2 / Azure DevOps Server (2025+) typically uses 7.2"
}
