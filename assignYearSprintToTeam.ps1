[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$Organization,  # e.g. https://dev.azure.com/your-org
    [Parameter(Mandatory = $true)][string]$Project,       # e.g. Azure Boards Rollout 2

    # Assign to one team (cleanest). You can wrap this to run for all teams if needed.
    [Parameter(Mandatory = $true)][string]$TeamName,

    # Year iteration node to target (e.g., 2027)
    [Parameter(Mandatory = $true)][int]$YearOfIteration,

    # Optional: only assign sprints whose startDate >= FromDate
    [Parameter(Mandatory = $false)][datetime]$FromDate = (Get-Date),

    # Optional: prevent duplicate POSTs by skipping already assigned iterations
    [Parameter(Mandatory = $false)][bool]$SkipIfAlreadyAssigned = $true
)

Write-Host "`nValues provided to the script:"
Write-Host "Organization: $Organization"
Write-Host "Project: $Project"
Write-Host "TeamName: $TeamName"
Write-Host "YearOfIteration: $YearOfIteration"
Write-Host "FromDate: $FromDate"
Write-Host "SkipIfAlreadyAssigned: $SkipIfAlreadyAssigned"
Write-Host "AZURE_DEVOPS_EXT_PAT is set: $([bool]$env:AZURE_DEVOPS_EXT_PAT)"
Write-Host "PAT length: $($env:AZURE_DEVOPS_EXT_PAT.Length)`n"

# Normalize base URL
$Organization = $Organization.TrimEnd('/')

# PAT auth
if (-not $env:AZURE_DEVOPS_EXT_PAT) {
    throw "Missing AZURE_DEVOPS_EXT_PAT. Set it as a secret pipeline variable and pass via env."
}

$pat = $env:AZURE_DEVOPS_EXT_PAT
$base64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))

$headers = @{
    Authorization = "Basic $base64"
    Accept        = "application/json;api-version=7.1"
    "Content-Type"= "application/json"
}

function Invoke-AdoRest {
    param(
        [Parameter(Mandatory)][ValidateSet("GET","POST","PATCH","PUT","DELETE")] [string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()] $Body
    )

    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 50
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $json "
    } else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }
}

$projectEsc = [uri]::EscapeDataString($Project)
$teamEsc    = [uri]::EscapeDataString($TeamName)
$yearName   = $YearOfIteration.ToString()

# -----------------------------
# 1) Load iteration tree and find the year node
# Classification Nodes (Iterations) supports $depth for children. [1](https://learn.microsoft.com/en-us/rest/api/azure/devops/wit/classification-nodes/get?view=azure-devops-rest-7.1)
# -----------------------------
$treeUri = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations?`$depth=4"
$tree = Invoke-AdoRest -Method GET -Uri $treeUri

$yearNode = @($tree.children) | Where-Object { $_.name -eq $yearName } | Select-Object -First 1
if (-not $yearNode) {
    throw "Year iteration '$yearName' not found under Project Iterations. Ensure the year node exists first."
}

# Child nodes under the year are sprints (iteration nodes) and include 'identifier'. [1](https://learn.microsoft.com/en-us/rest/api/azure/devops/wit/classification-nodes/get?view=azure-devops-rest-7.1)
$sprints = @($yearNode.children)
if ($sprints.Count -eq 0) {
    Write-Host "No sprints found under year '$yearName'. Nothing to assign."
    return
}

# Optional filter by FromDate using sprint attributes.startDate (if present)
$fromUtc = $FromDate.ToUniversalTime()
$sprintsToAssign = @()

foreach ($s in $sprints) {
    if ($s.attributes -and $s.attributes.startDate) {
        $sd = [datetime]$s.attributes.startDate
        if ($sd.ToUniversalTime() -ge $fromUtc) { $sprintsToAssign += $s }
    } else {
        # If no date attributes exist, include it
        $sprintsToAssign += $s
    }
}

Write-Host "Sprints found under '$yearName': $($sprints.Count)"
Write-Host "Sprints to assign (after FromDate): $($sprintsToAssign.Count)"

# -----------------------------
# 2) (Optional) Get already assigned iterations for the team to avoid duplicates
# Iterations - List returns the team's assigned iterations. [3](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/iterations/list?view=azure-devops-rest-7.1)
# -----------------------------
$alreadyAssigned = @()
if ($SkipIfAlreadyAssigned) {
    $listUri = "$Organization/$projectEsc/$teamEsc/_apis/work/teamsettings/iterations?api-version=7.1"
    $listResp = Invoke-AdoRest -Method GET -Uri $listUri
    $alreadyAssigned = @($listResp.values | Select-Object -ExpandProperty id)
    Write-Host "Already assigned to team '$TeamName': $($alreadyAssigned.Count)"
}

# -----------------------------
# 3) Assign each sprint to the team
# Iterations - Post Team Iteration: POST teamsettings/iterations with body { id: <guid> }. [2](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/iterations/post-team-iteration?view=azure-devops-rest-7.1)
# -----------------------------
$assignUri = "$Organization/$projectEsc/$teamEsc/_apis/work/teamsettings/iterations?api-version=7.1"

foreach ($s in $sprintsToAssign) {
    $iterId = $s.identifier
    $iterName = $s.name

    if ($SkipIfAlreadyAssigned -and ($alreadyAssigned -contains $iterId)) {
        Write-Host "Skipping (already assigned): $iterName"
        continue
    }

    if ($PSCmdlet.ShouldProcess($TeamName, "Assign sprint '$iterName'")) {
        try {
            Invoke-AdoRest -Method POST -Uri $assignUri -Body @{ id = $iterId } | Out-Null
            Write-Host "Assigned: $iterName"
        }
        catch {
            Write-Host "Warning: Failed to assign '$iterName' to '$TeamName'. Error: $($_.Exception.Message)"
        }
    }
}

Write-Host "`nDone (assign sprints under year)."
