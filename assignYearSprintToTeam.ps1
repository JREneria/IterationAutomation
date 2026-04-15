[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$Organization,  # e.g. https://dev.azure.com/your-org
    [Parameter(Mandatory = $true)][string]$Project,       # e.g. Azure Boards Rollout 2

    # Optional: if empty, assign to ALL teams
    [Parameter(Mandatory = $false)][string]$TeamName,

    # Year iteration node to target (e.g., 2027). If 0, defaults to current year.
    [Parameter(Mandatory = $false)][int]$YearOfIteration = 0,

    # Optional: prevent duplicate POSTs by skipping already assigned iterations
    [Parameter(Mandatory = $false)][bool]$SkipIfAlreadyAssigned = $true
)

Write-Host "`nValues provided to the script:"
Write-Host "Organization: $Organization"
Write-Host "Project: $Project"
Write-Host "TeamName: $TeamName"
Write-Host "YearOfIteration: $YearOfIteration"
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
    Accept        = "application/json"
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
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $json -ContentType "application/json"
    } else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }
}

function Resolve-YearAndFromDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$YearOfIteration = 0
    )

    $now = Get-Date
    $currentYear = $now.Year
    $year = if ($YearOfIteration -eq 0) { $currentYear } else { $YearOfIteration }

    if ($year -eq $currentYear) {
        $fromDate = $now.Date
        Write-Host "YearOfIteration = current year ($year). Using FromDate = today: $fromDate"
    }
    elseif ($year -gt $currentYear) {
        $fromDate = (Get-Date -Year $year -Month 1 -Day 1).Date
        Write-Host "YearOfIteration = future year ($year). Using FromDate = start of year: $fromDate"
    }
    else {
        # Choose your policy for past years:
        # Option A: allow past assignment from Jan 1 of that year
        $fromDate = (Get-Date -Year $year -Month 1 -Day 1).Date
        Write-Host "YearOfIteration = past year ($year). Using FromDate = start of year: $fromDate"

        # Option B (safer): block past years
        # throw "YearOfIteration ($year) is in the past. Refusing to assign historical sprints."
    }

    return [pscustomobject]@{
        Year        = $year
        YearName    = $year.ToString()
        FromDate    = $fromDate
        FromDateUtc = $fromDate.ToUniversalTime()
    }
}

function Get-ProjectId {
    param([string]$Org, [string]$ProjectName)

    $projectEsc = [uri]::EscapeDataString($ProjectName)
    $uri = "$Org/_apis/projects/$projectEsc?api-version=7.1"
    $p = Invoke-AdoRest -Method GET -Uri $uri
    return $p.id
}

function Get-TeamList {
    param(
        [string]$Org,
        [string]$ProjectName,
        [string]$TeamNameOrEmpty
    )

    if ($TeamNameOrEmpty -and $TeamNameOrEmpty.Trim().Length -gt 0) {
        return @($TeamNameOrEmpty.Trim())
    }
    
    else if ($TeamNameOrEmpty -eq "auto") {

    Write-Host "No TeamName provided. Resolving ALL teams in project..."

    $projectId = Get-ProjectId -Org $Org -ProjectName $ProjectName
    if (-not $projectId) { throw "Failed to resolve projectId for '$ProjectName'." }

    $teamsUri = "$Org/_apis/projects/$projectId/teams?api-version=7.1"
    $teamsResp = Invoke-AdoRest -Method GET -Uri $teamsUri

    $names = @($teamsResp.value | Select-Object -ExpandProperty name)
    if ($names.Count -eq 0) { throw "No teams found in project '$ProjectName'." }

    return $names
    }
}

# Resolve year + implied FromDate
$resolved = Resolve-YearAndFromDate -YearOfIteration $YearOfIteration
$yearName = $resolved.YearName
$fromUtc  = $resolved.FromDateUtc

$projectEsc = [uri]::EscapeDataString($Project)

# 1) Load iteration tree and find the year node (Classification Nodes - Get w/ $depth) [1](https://learn.microsoft.com/en-us/rest/api/azure/devops/wit/classification-nodes/get?view=azure-devops-rest-7.1)
$treeUri = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations?`$depth=4&api-version=7.1"
$tree = Invoke-AdoRest -Method GET -Uri $treeUri

$yearNode = @($tree.children) | Where-Object { $_.name -eq $yearName } | Select-Object -First 1
if (-not $yearNode) {
    throw "Year iteration '$yearName' not found under Project Iterations. Ensure the year node exists first."
}

# Child nodes under the year are sprints and include 'identifier'. [1](https://learn.microsoft.com/en-us/rest/api/azure/devops/wit/classification-nodes/get?view=azure-devops-rest-7.1)
$sprints = @($yearNode.children)
if ($sprints.Count -eq 0) {
    Write-Host "No sprints found under year '$yearName'. Nothing to assign."
    return
}

# Filter by FromDate (implied)
$sprintsToAssign = @()
foreach ($s in $sprints) {
    if ($s.attributes -and $s.attributes.startDate) {
        $sd = [datetime]$s.attributes.startDate
        if ($sd.ToUniversalTime() -ge $fromUtc) { $sprintsToAssign += $s }
    } else {
        $sprintsToAssign += $s
    }
}

Write-Host "Sprints found under '$yearName': $($sprints.Count)"
Write-Host "Sprints to assign (after implied FromDate): $($sprintsToAssign.Count)"

# 2) Assign each sprint to one team OR all teams if none provided
$teamList = Get-TeamList -Org $Organization -ProjectName $Project -TeamNameOrEmpty $TeamName

foreach ($team in $teamList) {

    $teamEsc = [uri]::EscapeDataString($team)

    # Work Iterations - Post Team Iteration: add iteration to the team [2](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/iterations/post-team-iteration?view=azure-devops-rest-7.1)
    $assignUri = "$Organization/$projectEsc/$teamEsc/_apis/work/teamsettings/iterations?api-version=7.1"

    # Optional: load assigned iterations for THIS team (idempotency) [3](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/iterations/list?view=azure-devops-rest-7.1)
    $alreadyAssigned = @()
    if ($SkipIfAlreadyAssigned) {
        $listUri = "$Organization/$projectEsc/$teamEsc/_apis/work/teamsettings/iterations?api-version=7.1"
        $listResp = Invoke-AdoRest -Method GET -Uri $listUri
        $alreadyAssigned = @($listResp.values | Select-Object -ExpandProperty id)
        Write-Host "Already assigned to team '$team': $($alreadyAssigned.Count)"
    }

    foreach ($s in $sprintsToAssign) {
        $iterId   = $s.identifier
        $iterName = $s.name

        if ($SkipIfAlreadyAssigned -and ($alreadyAssigned -contains $iterId)) {
            Write-Host "Skipping (already assigned) [$team]: $iterName"
            continue
        }

        if ($PSCmdlet.ShouldProcess($team, "Assign sprint '$iterName'")) {
            try {
                Invoke-AdoRest -Method POST -Uri $assignUri -Body @{ id = $iterId } | Out-Null
                Write-Host "Assigned [$team]: $iterName"
            }
            catch {
                Write-Host "Warning: Failed to assign '$iterName' to '$team'. Error: $($_.Exception.Message)"
            }
        }
    }
}

Write-Host "`nDone (assign sprints under year)."
