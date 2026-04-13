[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$Organization,  # e.g. https://dev.azure.com/your-org
    [Parameter(Mandatory = $true)][string]$Project,       # e.g. Azure Boards Rollout 2

    [Parameter(Mandatory = $true)][int]$YearOfIteration,
    [Parameter(Mandatory = $true)][datetime]$StartDate,
    [Parameter(Mandatory = $true)][int]$NumberOfSprints,

    [Parameter(Mandatory = $false)][int]$SprintLengthDays = 5,
    [Parameter(Mandatory = $false)][int]$GapDays = 2
)

Write-Host "`nValues provided to the script:"
Write-Host "Organization: $Organization"
Write-Host "Project: $Project"
Write-Host "YearOfIteration: $YearOfIteration"
Write-Host "StartDate: $StartDate"
Write-Host "NumberOfSprints: $NumberOfSprints"
Write-Host "SprintLengthDays: $SprintLengthDays"
Write-Host "GapDays: $GapDays"
Write-Host "AZURE_DEVOPS_EXT_PAT is set: $([bool]$env:AZURE_DEVOPS_EXT_PAT)`n"

# Normalize org URL once
$Organization = $Organization.TrimEnd('/')

# --- PAT / Auth ---
if (-not $env:AZURE_DEVOPS_EXT_PAT) {
    throw "Missing AZURE_DEVOPS_EXT_PAT. Set it as a secret pipeline variable and pass via env."
}

$pat = $env:AZURE_DEVOPS_EXT_PAT
$base64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))

# Supply api-version via Accept header (supported), and JSON content type for POST bodies. [2](https://johnnyreilly.com/list-pipelines-with-azure-devops-api)
$headers = @{
    Authorization = "Basic $base64"
    Accept        = "application/json;api-version=7.1"
    "Content-Type"= "application/json"
}

function Get-SprintWindowsToEndOfYear {
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][int]$SprintLengthDays,
        [Parameter(Mandatory)][int]$GapDays,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    if ($SprintLengthDays -lt 1) { throw "SprintLengthDays must be >= 1" }
    if ($GapDays -lt 0) { throw "GapDays must be >= 0" }

    $windows = @()
    $start = $StartDate.Date

    while ($start -le $EndDate.Date) {
        $finish = $start.AddDays($SprintLengthDays - 1)

        # Policy A: clamp finish date to EndDate
        if ($finish -gt $EndDate.Date) {
            $finish = $EndDate.Date
        }

        $windows += [pscustomobject]@{
            Start  = $start
            Finish = $finish
        }

        # Next sprint start = finish + gap + 1
        $start = $finish.AddDays($GapDays + 1)
    }

    return $windows
}

function Invoke-AdoRest {
    param(
        [Parameter(Mandatory)][ValidateSet("GET","POST","PATCH","DELETE")] [string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()] $Body
    )

    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 20
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $json -ContentType "application/json"
    } else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }
}

# Encode project for URL
$projectEsc = [uri]::EscapeDataString($Project)
$yearName   = $YearOfIteration.ToString()

# =========================
# 1) Ensure annual iteration exists (REST)
# =========================
Write-Host "`n=== Ensuring annual iteration '$yearName' exists (REST) ==="

# Classification Nodes (Iterations) GET by depth for tree inspection. [1](https://learn.microsoft.com/en-us/rest/api/azure/devops/release/releases/list?view=azure-devops-rest-7.1)
$getIterationsUri = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations?`$depth=2"
$iterationsTree = Invoke-AdoRest -Method GET -Uri $getIterationsUri

$yearNode = $null
if ($iterationsTree.children) {
    $yearNode = @($iterationsTree.children) | Where-Object { $_.name -eq $yearName } | Select-Object -First 1
}

if (-not $yearNode) {
    $yearStart  = Get-Date -Year $YearOfIteration -Month 1 -Day 1
    $yearFinish = Get-Date -Year $YearOfIteration -Month 12 -Day 31

    Write-Host "Annual iteration missing. Creating '$yearName'..."

    # Classification Nodes (Iterations) POST to create under root. [1](https://learn.microsoft.com/en-us/rest/api/azure/devops/release/releases/list?view=azure-devops-rest-7.1)
    $createYearUri = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations"
    $yearBody = @{
        name = $yearName
        attributes = @{
            startDate  = $yearStart.ToString("o")
            finishDate = $yearFinish.ToString("o")
        }
    }

    if ($PSCmdlet.ShouldProcess($Project, "Create annual iteration $yearName")) {
        $yearNode = Invoke-AdoRest -Method POST -Uri $createYearUri -Body $yearBody
        Write-Host "Created annual iteration: $($yearNode.name)"
    }
} else {
    Write-Host "Annual iteration '$yearName' already exists."
}

# =========================
# 2) Load existing sprints under the year (idempotency)
# =========================
Write-Host "`n=== Loading existing sprints under '$yearName' ==="

$getIterationsUriDeep = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations?`$depth=3"
$iterationsTreeDeep = Invoke-AdoRest -Method GET -Uri $getIterationsUriDeep

$yearNodeDeep = $null
if ($iterationsTreeDeep.children) {
    $yearNodeDeep = @($iterationsTreeDeep.children) | Where-Object { $_.name -eq $yearName } | Select-Object -First 1
}

# Map sprintName -> identifier
$existingSprintByName = @{}
if ($yearNodeDeep -and $yearNodeDeep.children) {
    foreach ($child in $yearNodeDeep.children) {
        if ($child.name -and $child.identifier) {
            $existingSprintByName[$child.name] = $child.identifier
        }
    }
}

# =========================
# 3) Create sprints under the year (REST) - NO ASSIGNMENT
# =========================
Write-Host "`n=== Creating sprints ==="
# Decide how many sprints we will create
$yearEnd = Get-Date -Year $YearOfIteration -Month 12 -Day 31

if ($NumberOfSprints -gt 0) {
    # Manual mode: create exactly N sprints
    $sprintWindows = @()
    $startDateIteration = $StartDate.Date

    for ($i=1; $i -le $NumberOfSprints; $i++) {
        $finishDateIteration = $startDateIteration.AddDays($SprintLengthDays - 1)
        $sprintWindows += [pscustomobject]@{ Start = $startDateIteration; Finish = $finishDateIteration }
        $startDateIteration = $finishDateIteration.AddDays($GapDays + 1)
    }
}
else {
    # Auto mode: create until end of year
    $sprintWindows = Get-SprintWindowsToEndOfYear `
        -StartDate $StartDate `
        -SprintLengthDays $SprintLengthDays `
        -GapDays $GapDays `
        -EndDate $yearEnd

    Write-Host "Auto-calculated sprint count: $($sprintWindows.Count) (from $($StartDate.Date) to $yearEnd)"
}

foreach ($w in $sprintWindows) {

    $startDateIteration = $w.Start
    $finishDateIteration = $w.Finish

    $weekNumber = [System.Globalization.ISOWeek]::GetWeekOfYear($startDateIteration)

    $sprintName = "Week $weekNumber - " +
        $startDateIteration.ToString("MM.dd.yyyy") + " - " +
        $finishDateIteration.ToString("MM.dd.yyyy")

    if ($existingSprintByName.ContainsKey($sprintName)) {
        Write-Host "Sprint exists: $sprintName. Skipping create."
        continue
    }

    Write-Host "Creating sprint: $sprintName"

    $createSprintUri = "$Organization/$projectEsc/_apis/wit/classificationnodes/Iterations/$yearName"
    $sprintBody = @{
        name = $sprintName
        attributes = @{
            startDate  = $startDateIteration.ToString("o")
            finishDate = $finishDateIteration.ToString("o")
        }
    }

    $createdSprint = Invoke-AdoRest -Method POST -Uri $createSprintUri -Body $sprintBody
    Write-Host "Created sprint: $($createdSprint.name)"
    $existingSprintByName[$sprintName] = $createdSprint.identifier
}

Write-Host "`nDone (create only)."
