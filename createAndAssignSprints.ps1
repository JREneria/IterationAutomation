
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$Organization,  # e.g. https://dev.azure.com/your-org
    [Parameter(Mandatory = $true)][string]$Project,       # e.g. Azure Boards Rollout 2

    [Parameter(Mandatory = $false)][string]$Teams,
    [Parameter(Mandatory = $true)][int]$YearOfIteration,
    [Parameter(Mandatory = $true)][datetime]$StartDate,
    [Parameter(Mandatory = $true)][int]$NumberOfSprints,

    [Parameter(Mandatory = $false)][int]$SprintLengthDays = 5,
    [Parameter(Mandatory = $false)][int]$GapDays = 2
)
$Organization = $Organization.TrimEnd('/')
Write-Host "`nValues provided to the script:"
Write-Host "Organization: $Organization"
Write-Host "Project: $Project"
Write-Host "Teams: $Teams"
Write-Host "YearOfIteration: $YearOfIteration"
Write-Host "StartDate: $StartDate"
Write-Host "NumberOfSprints: $NumberOfSprints"
Write-Host "SprintLengthDays: $SprintLengthDays"
Write-Host "GapDays: $GapDays"
Write-Host "AZURE_DEVOPS_EXT_PAT is set: $([bool]$env:AZURE_DEVOPS_EXT_PAT)`n"

# --- PAT / Auth ---
if (-not $env:AZURE_DEVOPS_EXT_PAT) {
    throw "Missing AZURE_DEVOPS_EXT_PAT. Set it as a secret pipeline variable and pass via env."
}

$pat = $env:AZURE_DEVOPS_EXT_PAT
$base64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))
$headers = @{
    Authorization = "Basic $base64"
    Accept = "application/json;api-version=7.1"
    "Content-Type" = "application/json"
}

function Invoke-AdoRest {
    param(
        [Parameter(Mandatory)] [ValidateSet("GET","POST","PATCH","DELETE")] [string]$Method,
        [Parameter(Mandatory)] [string]$Uri,
        [Parameter()] $Body
    )
    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 20
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $json -ContentType "application/json"
    } else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }
}

function Add-IterationToTeams {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Organization,

        [Parameter(Mandatory = $true)]
        [string]$ProjectEsc,   # already URL-encoded project name

        [Parameter(Mandatory = $true)]
        [string[]]$TeamList,

        [Parameter(Mandatory = $true)]
        [string]$IterationId,  # GUID (createdSprint.identifier)

        [Parameter(Mandatory = $true)]
        [string]$SprintName    # for logging
    )

    foreach ($team in $TeamList) {
        $teamEsc = [uri]::EscapeDataString($team)

        # IMPORTANT: use ${} when a variable is followed by ? in interpolated strings
        $assignUri = "$Organization/$ProjectEsc/$teamEsc/_apis/work/teamsettings/iterations"

        try {
            $assignBody = @{ id = $IterationId }
            Invoke-AdoRest -Method POST -Uri $assignUri -Body $assignBody | Out-Null
            Write-Host "Assigned: $SprintName to Team: $team"
        }
        catch {
            Write-Host "Warning: Could not assign sprint '$SprintName' to team '$team'. Error: $($_.Exception.Message)"
        }
    }
}

function Resolve-ProjectId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Organization,
        [Parameter(Mandatory=$true)][string]$ProjectName
    )

    $Organization = $Organization.TrimEnd('/')

    $listUri = "$Organization/_apis/projects?`$top=1000&api-version=7.1"
    Write-Host "GET $listUri"

    try {
        $resp = Invoke-AdoRest -Method GET -Uri $listUri
    }
    catch {
        Write-Host "❌ Projects list call failed."
        Write-Host "Exception: $($_.Exception.Message)"

        # Try to extract HTTP status + response body (super useful)
        if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream()) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $body = $reader.ReadToEnd()
                Write-Host "---- Response Body (first 500 chars) ----"
                Write-Host ($body.Substring(0, [Math]::Min(500, $body.Length)))
                Write-Host "----------------------------------------"
            } catch {}
        }
        throw
    }

    if (-not $resp.value) {
        # If resp is a string/HTML, it will have Length and no value
        $typeName = $resp.GetType().FullName
        Write-Host "⚠ Unexpected response type: $typeName"
        if ($resp -is [string]) {
            Write-Host "---- Response (first 200 chars) ----"
            Write-Host ($resp.Substring(0, [Math]::Min(5000, $resp.Length)))
            Write-Host "------------------------------------"
        }
        throw "Could not list projects from org '$Organization'. Likely auth/permission issue."
    }

    $match = $resp.value | Where-Object { $_.name -ieq $ProjectName } | Select-Object -First 1
    if (-not $match) {
        $names = ($resp.value | Select-Object -ExpandProperty name | Sort-Object) -join ", "
        throw "Project '$ProjectName' not found. Available projects: $names"
    }

    return $match.id
}

# --- Normalize / Encode path segments for REST URLs ---
# REST URLs require URL-encoded project/team names if they contain spaces. [3](https://learn.microsoft.com/en-us/azure/devops/cli/troubleshooting?view=azure-devops)
$projectEsc = [uri]::EscapeDataString($Project)


# -------------------------
# 1) RESOLVE TEAMS
# -------------------------
$teamList = @()

if ($Teams -and $Teams.Trim().Length -gt 0) {
    $teamList = $Teams -split "," |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne "" }

    Write-Host "Using explicitly provided teams:"
    $teamList | ForEach-Object { Write-Host " - $_" }
}
else {
    Write-Host "No teams specified. Resolving ALL teams in the project..."

    # Get project ID
    # $projectInfoUri = "${Organization}/_apis/projects/${projectEsc}?api-version=7.1"
    # $projectInfo = Invoke-AdoRest -Method GET -Uri $projectInfoUri
    $projectId = Resolve-ProjectId -Organization $Organization -ProjectName $Project
    Write-Host "Resolved Project ID: $projectId"

    if (-not $projectId) {
        throw "Failed to resolve project ID for project '$Project'"
    }

    # Get all teams in project
    $teamsUri = "$Organization/_apis/projects/$projectId/teams"
    $teamsResponse = Invoke-AdoRest -Method GET -Uri $teamsUri

    $teamList = $teamsResponse.value | Select-Object -ExpandProperty name

    Write-Host "Discovered teams:"
    $teamList | ForEach-Object { Write-Host " - $_" }
}

if ($teamList.Count -eq 0) {
    throw "No teams resolved. Aborting."
}

# =========================
# 1) Ensure annual iteration exists (REST)
# =========================
Write-Host "`n=== Ensuring annual iteration '$YearOfIteration' exists (REST) ==="
# Get Iterations root and its children (top-level nodes like 2026, 2027, etc.) 
$getIterationsUri = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations?`$depth=2"
$iterationsTree = Invoke-AdoRest -Method GET -Uri $getIterationsUri

$yearName = "$YearOfIteration"
$yearNode = $null
if ($iterationsTree.children) {
    $yearNode = @($iterationsTree.children) | Where-Object { $_.name -eq $yearName } | Select-Object -First 1
}

if (-not $yearNode) {
    $yearStart  = Get-Date -Year $YearOfIteration -Month 1 -Day 1
    $yearFinish = Get-Date -Year $YearOfIteration -Month 12 -Day 31

    Write-Host "Annual iteration missing. Creating '$YearOfIteration' at top-level under Iterations root..."
    if ($PSCmdlet.ShouldProcess("$Project", "Create annual iteration $YearOfIteration")) {

        # Create year node directly under Iterations root (no CLI --path, no \Iteration). 
        $createYearUri = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations"
        $yearBody = @{
            name = $yearName
            attributes = @{
                startDate  = $yearStart.ToString("o")
                finishDate = $yearFinish.ToString("o")
            }
        }

        $yearNode = Invoke-AdoRest -Method POST -Uri $createYearUri -Body $yearBody
        Write-Host "Created annual iteration: $($yearNode.name)"
    }
} else {
    Write-Host "Annual iteration '$YearOfIteration' already exists."
}

# =========================
# 2) Load existing sprints under the year (idempotency)
# =========================
Write-Host "`n=== Loading existing sprints under '$YearOfIteration' ==="

# Re-fetch deeper so we can see sprints under the year
$getIterationsUriDeep = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations?`$depth=3"
$iterationsTreeDeep = Invoke-AdoRest -Method GET -Uri $getIterationsUriDeep

$yearNodeDeep = $null
if ($iterationsTreeDeep.children) {
    $yearNodeDeep = @($iterationsTreeDeep.children) | Where-Object { $_.name -eq $yearName } | Select-Object -First 1
}


# Build a lookup: name -> identifier (GUID)
$existingSprintByName = @{}

if ($yearNodeDeep -and $yearNodeDeep.children) {
    foreach ($child in $yearNodeDeep.children) {
        # child.name is sprint name, child.identifier is GUID
        if ($child.name -and $child.identifier) {
            $existingSprintByName[$child.name] = $child.identifier
        }
    }
}


# =========================
# 3) Create sprints under the year + assign to teams (REST)
# =========================
$startDateIteration = $StartDate

for ($i = 1; $i -le $NumberOfSprints; $i++) {

    $finishDateIteration = $startDateIteration.AddDays($SprintLengthDays - 1)

    # ISO week number (stable, Monday-based)
    $weekNumber = [System.Globalization.ISOWeek]::GetWeekOfYear($startDateIteration)

    $sprintName = "Week $weekNumber - " +
        $startDateIteration.ToString("MM.dd.yyyy") + " - " +
        $finishDateIteration.ToString("MM.dd.yyyy")

    if ($existingSprintByName.ContainsKey($sprintName)) {
        $existingId = $existingSprintByName[$sprintName]
        Write-Host "Sprint exists: $sprintName. Assigning to teams..."
        Add-IterationToTeams -Organization $Organization -ProjectEsc $projectEsc -TeamList $teamList -IterationId $existingId -SprintName $sprintName
        
        # advance to next sprint window
        $startDateIteration = $finishDateIteration.AddDays($GapDays + 1)
        continue
    }

    Write-Host "`nCreating sprint: $sprintName"
    if ($PSCmdlet.ShouldProcess("$Project", "Create sprint $sprintName under $YearOfIteration")) {

        # Create sprint as a child node under the year: POST .../Iterations/{year} 
        $createSprintUri = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations/${YearOfIteration}"
        $sprintBody = @{
            name = $sprintName
            attributes = @{
                startDate  = $startDateIteration.ToString("o")
                finishDate = $finishDateIteration.ToString("o")
            }
        }

        $createdSprint = Invoke-AdoRest -Method POST -Uri $createSprintUri -Body $sprintBody

        # Assign sprint to each team via Work Iterations API (POST Team Iteration). 
        Add-IterationToTeams -Organization $Organization -ProjectEsc $projectEsc -TeamList $teamList -IterationId $createdSprint.identifier -SprintName $sprintName

        # Update our idempotency set so re-runs are clean within the same run
        $existingSprintByName[$sprintName] = $createdSprint.identifier
    }

    # Move to next sprint window
    $startDateIteration = $finishDateIteration.AddDays($GapDays + 1)
}

Write-Host "`nDone."
