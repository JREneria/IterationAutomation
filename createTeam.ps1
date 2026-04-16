[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$Organization,   # e.g. https://dev.azure.com/your-org

    [Parameter(Mandatory = $true)]
    [string]$Project,        # Project name

    [Parameter(Mandatory = $true)]
    [string]$TeamName,       # New team name

    [Parameter(Mandatory = $true)]
    [string]$ClientsJson,    # JSON array e.g. ["ClientA","ClientB"]
    
    [Parameter(Mandatory = $false)]
    [string]$RolesJson = '["Developers","Business Analysts", "Benefits System Administrators", "Project Managers"]',

    # Safety switches
    [Parameter(Mandatory = $false)]
    [bool]$DryRun = $false,

    [Parameter(Mandatory = $false)]
    [bool]$SkipTeamFieldValues = $false,

    [Parameter(Mandatory = $false)]
    [bool]$SkipTeamMembershipGroups = $false,

    [Parameter(Mandatory = $false)]
    [bool]$SkipIterationAssignment = $false
)

# ----------------------------
# Logging: inputs
# ----------------------------
Write-Host "`n=== Bootstrap Team Script ==="
Write-Host "Organization: $Organization"
Write-Host "Project: $Project"
Write-Host "TeamName: $TeamName"
Write-Host "YearOfIteration: $YearOfIteration"
Write-Host "DryRun: $DryRun"
Write-Host "SkipTeamFieldValues: $SkipTeamFieldValues"
Write-Host "SkipTeamMembershipGroups: $SkipTeamMembershipGroups"
Write-Host "SkipIterationAssignment: $SkipIterationAssignment"

# Normalize org URL
$Organization = $Organization.TrimEnd('/')

# ----------------------------
# Auth token
# ----------------------------
if (-not $env:AZURE_DEVOPS_EXT_PAT) {
    throw "Missing AZURE_DEVOPS_EXT_PAT. In Azure Pipelines, map env: AZURE_DEVOPS_EXT_PAT: $(System.AccessToken) or a PAT."
}
$token = $env:AZURE_DEVOPS_EXT_PAT
Write-Host "Token provided: $([bool]$token) (length=$($token.Length))"

# Detect token type:
# - System.AccessToken is typically a JWT (starts with eyJ...)
# - PAT is opaque; use Basic with ":PAT"
$authHeaderValue = $null
if ($token -match '^eyJ') {
    $authHeaderValue = "Bearer $token"
    Write-Host "Auth mode: Bearer (System.AccessToken/JWT)"
} else {
    $base64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$token"))
    $authHeaderValue = "Basic $base64"
    Write-Host "Auth mode: Basic (PAT)"
}

$headers = @{
    Authorization = $authHeaderValue
    Accept        = "application/json;api-version=7.1"
    "Content-Type"= "application/json"
}

# ----------------------------
# Helpers: HTTP error parsing (PS7 + Windows PS)
# ----------------------------
function Get-HttpStatusCode {
    param([object]$ErrorRecord)
    $resp = $ErrorRecord.Exception.Response
    if ($resp -is [System.Net.Http.HttpResponseMessage]) {
        return [int]$resp.StatusCode
    }
    if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
        return [int]$ErrorRecord.Exception.Response.StatusCode
    }
    return $null
}

function Get-HttpErrorBody {
    param([object]$ErrorRecord)

    $resp = $ErrorRecord.Exception.Response
    if ($resp -and $resp -is [System.Net.Http.HttpResponseMessage]) {
        try { return $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult() } catch { return $null }
    }
    if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.GetResponseStream) {
        try {
            $reader = New-Object System.IO.StreamReader($ErrorRecord.Exception.Response.GetResponseStream())
            return $reader.ReadToEnd()
        } catch { return $null }
    }
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return $ErrorRecord.ErrorDetails.Message
    }
    return $null
}

# ----------------------------
# Helper: REST caller (api-version in URL per guidance) [13](https://learn.microsoft.com/en-us/azure/devops/integrate/how-to/call-rest-api?view=azure-devops)[14](https://learn.microsoft.com/en-us/rest/api/azure/devops/?view=azure-devops-rest-7.2)
# ----------------------------
function Invoke-AdoRest {
    param(
        [Parameter(Mandatory)][ValidateSet("GET","POST","PATCH","PUT","DELETE")] [string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()] $Body
    )
    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 50
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $json
    } else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }
}

# ----------------------------
# Parse clients list
# ----------------------------
try {
    $clients = @($ClientsJson | ConvertFrom-Json)
} catch {
    throw "ClientsJson must be a JSON array, e.g. ['ClientA','ClientB']. Received: $ClientsJson"
}
$clients = $clients | ForEach-Object { "$_".Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique
if ($clients.Count -lt 1) { throw "At least 1 client is required." }

Write-Host "Clients resolved ($($clients.Count)): $($clients -join ', ')"

# ----------------------------
# Parse roles list
# ----------------------------
$roles = @($RolesJson | ConvertFrom-Json) | ForEach-Object { "$_".Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique
if ($roles.Count -lt 1) { throw "At least 1 role is required." }
Write-Host "Roles resolved ($($roles.Count)): $($roles -join ', ')"

# ----------------------------
# Resolve org short name (for vssps.dev.azure.com Graph endpoints)
# ----------------------------
function Get-OrgNameFromUrl {
    param([string]$OrgUrl)
    # Handles https://dev.azure.com/{org}
    if ($OrgUrl -match '^https://dev\.azure\.com/([^/]+)') {
        return $Matches[1]
    }
    throw "Organization URL must be in form https://dev.azure.com/{org}. Got: $OrgUrl"
}
$orgName = Get-OrgNameFromUrl -OrgUrl $Organization

# ----------------------------
# (1) Resolve ProjectId (idempotent lookup)
# Uses Projects - List and filters by name. ://learn.microsoft.com/en-us/rest/api/azure/devops/core/projects/list?view=azure-devops-rest-7.1)
# ----------------------------
function Get-ProjectIdByName {
    param([string]$Org, [string]$ProjectName)

    Write-Host "[Get-ProjectIdByName] Resolving projectId for '$ProjectName'..."

    # If it's already a GUID, accept it
    if ($ProjectName -match '^[0-9a-fA-F-]{36}$') { return $ProjectName }

    $skip = 0
    $top  = 100
    while ($true) {
        $uri = "$Org/_apis/projects?`$top=$top&`$skip=$skip&api-version=7.1"
        $resp = Invoke-AdoRest -Method GET -Uri $uri
        $match = @($resp.value) | Where-Object { $_.name -eq $ProjectName } | Select-Object -First 1
        if ($match) { return $match.id }

        if (-not $resp.value -or $resp.value.Count -lt $top) { break }
        $skip += $top
    }

    return $null
}

$projectId = Get-ProjectIdByName -Org $Organization -ProjectName $Project
if (-not $projectId) {
    throw "Project '$Project' not found or not accessible to this identity."
}
Write-Host "ProjectId: $projectId"

$projectEsc = [uri]::EscapeDataString($Project)
$teamEsc    = [uri]::EscapeDataString($TeamName)

# ----------------------------
# (b) Ensure team exists (idempotent)
# Teams - Get Teams (list) [2](https://learn.microsoft.com/en-us/rest/api/azure/devops/core/teams/get-teams?view=azure-devops-rest-7.1)
# Teams - Create (POST) [1](https://learn.microsoft.com/en-us/rest/api/azure/devops/core/teams/create?view=azure-devops-rest-7.1)
# ----------------------------
function Ensure-Team {
    param([string]$Org, [string]$ProjectId, [string]$TeamName, [bool]$DryRun)

    Write-Host "[Ensure-Team] Ensuring team exists: '$TeamName'..."

    $listUri = "$Org/_apis/projects/$ProjectId/teams"
    $teams = Invoke-AdoRest -Method GET -Uri $listUri
    $existing = @($teams.value) | Where-Object { $_.name -eq $TeamName } | Select-Object -First 1
    if ($existing) {
        Write-Host "[Ensure-Team] Team already exists. Reusing: '$TeamName' (id=$($existing.id))"
        return $existing
    }

    if ($DryRun -or -not $PSCmdlet.ShouldProcess($TeamName, "Create Team")) {
        Write-Host "[Ensure-Team] DryRun/WhatIf: would create team '$TeamName' (skipping)."
        return $null
    }

    $createUri = "$Org/_apis/projects/$ProjectId/teams"
    $body = @{ name = $TeamName }

    try {
        $created = Invoke-AdoRest -Method POST -Uri $createUri -Body $body
        Write-Host "[Ensure-Team] Team created: '$TeamName' (id=$($created.id))"
        return $created
    } catch {
        Write-Warning "[Ensure-Team] Create failed; re-checking if team now exists (race condition)..."
        $teams2 = Invoke-AdoRest -Method GET -Uri $listUri
        $existing2 = @($teams2.value) | Where-Object { $_.name -eq $TeamName } | Select-Object -First 1
        if ($existing2) {
            Write-Host "[Ensure-Team] Team now exists. Reusing: '$TeamName' (id=$($existing2.id))"
            return $existing2
        }
        throw
    }
}

$teamObj = Ensure-Team -Org $Organization -ProjectId $projectId -TeamName $TeamName -DryRun $DryRun
# Even in DryRun, we can proceed with best-effort for other steps, but membership/iterations need a team path.
function Test-AreaExists {
    param(
        [string]$Org,
        [string]$ProjectEsc,
        [string[]]$FullPathSegments   # e.g. @('TeamName') or @('TeamName','ClientA')
    )

    $encodedPath = ($FullPathSegments | ForEach-Object { [uri]::EscapeDataString($_) }) -join "/"
    $uri = "$Org/$ProjectEsc/_apis/wit/classificationnodes/Areas/$encodedPath"

    try {
        Invoke-AdoRest -Method GET -Uri $uri | Out-Null
        return $true
    } catch {
        $code = $_.Exception.Response.StatusCode.value__
        if ($code -eq 404) { return $false }
        throw
    }
}
# ----------------------------
# (c) Ensure Areas: Project\TeamName\Client (idempotent)
# Classification Nodes - Create Or Update supports Areas and path nesting. [3](https://www.reddit.com/r/azuredevops/comments/ii7rn0/difference_between_azure_artifacts_and_pipeline/)[4](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/iterations/post-team-iteration?view=azure-devops-rest-7.1)
# ----------------------------
function Ensure-AreaNode {
    param(
        [string]$Org,
        [string]$ProjectEsc,
        [string[]]$ParentSegments,   # @() for root; @('TeamName') for under team
        [string]$Name,
        [bool]$DryRun
    )

    # Build full path (from Areas root)
    $fullPath = @()
    if ($ParentSegments) { $fullPath += $ParentSegments }
    $fullPath += $Name

    Write-Host "[Ensure-AreaNode] Target AreaPath: $($fullPath -join '\')"

    #  Idempotency: if it already exists, do nothing
    if (Test-AreaExists -Org $Org -ProjectEsc $ProjectEsc -FullPathSegments $fullPath) {
        Write-Host "[Ensure-AreaNode] Exists already. Skipping."
        return $null
    }

    # Build parent path for CREATE call
    $parentPath = ""
    if ($ParentSegments -and $ParentSegments.Count -gt 0) {
        $encoded = $ParentSegments | ForEach-Object { [uri]::EscapeDataString($_) }
        $parentPath = "/" + ($encoded -join "/")
    }

    $uri = "$Org/$ProjectEsc/_apis/wit/classificationnodes/Areas$parentPath"
    Write-Host "[Ensure-AreaNode] Creating under parent '$($ParentSegments -join '\')' name='$Name'"
    Write-Host "[Ensure-AreaNode] POST $uri"

    if ($DryRun -or -not $PSCmdlet.ShouldProcess("$($fullPath -join '\')", "Create Area Node")) {
        Write-Host "[Ensure-AreaNode] DryRun/WhatIf: skipping create."
        return $null
    }

    $body = @{ name = $Name }
    return Invoke-AdoRest -Method POST -Uri $uri -Body $body
}
# Ensure team root area node under project Areas root
$null = Ensure-AreaNode -Org $Organization -ProjectEsc $projectEsc -ParentSegments @() -Name $TeamName -DryRun $DryRun

# Ensure each client under TeamName
foreach ($c in $clients) {
    $null = Ensure-AreaNode -Org $Organization -ProjectEsc $projectEsc -ParentSegments @($TeamName) -Name $c -DryRun $DryRun
}

# ----------------------------
# (c) Configure Team Field Values (System.AreaPath)
# Teamfieldvalues - Update (PATCH) [5](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/teamfieldvalues/update?view=azure-devops-rest-7.1)[6](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/Teamfieldvalues/Get?view=azure-devops-rest-7.1)
# ----------------------------
function Update-TeamFieldValues {
    param(
        [string]$Org,
        [string]$ProjectEsc,
        [string]$TeamEsc,
        [string]$ProjectName,
        [string]$TeamName,
        [string[]]$Clients,
        [bool]$DryRun
    )

    Write-Host "[Update-TeamFieldValues] Setting team area paths (default + allowed values)..."

    $uri = "$Org/$ProjectEsc/$TeamEsc/_apis/work/teamsettings/teamfieldvalues?api-version=7.1"

    # ADO expects AreaPath values like "Project\Area\SubArea" (no leading backslash). [5](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/teamfieldvalues/update?view=azure-devops-rest-7.1)[6](https://learn.microsoft.com/en-us/rest/api/azure/devops/work/Teamfieldvalues/Get?view=azure-devops-rest-7.1)
    $teamRoot = "$ProjectName\$TeamName"

    $values = @()
    $values += @{ value = $teamRoot; includeChildren = $true }
    foreach ($c in $Clients) {
        $values += @{ value = "$ProjectName\$TeamName\$c"; includeChildren = $true }
    }

    $body = @{
        defaultValue = $teamRoot
        values       = $values
    }

    if ($DryRun -or -not $PSCmdlet.ShouldProcess($TeamName, "PATCH TeamFieldValues (Area Paths)")) {
        Write-Host "[Update-TeamFieldValues] DryRun/WhatIf: would PATCH $uri"
        Write-Host ($body | ConvertTo-Json -Depth 10)
        return
    }

    Invoke-AdoRest -Method PATCH -Uri $uri -Body $body | Out-Null
    Write-Host "[Update-TeamFieldValues] Team field values updated."
}

if (-not $SkipTeamFieldValues) {
    Update-TeamFieldValues -Org $Organization -ProjectEsc $projectEsc -TeamEsc $teamEsc `
        -ProjectName $Project -TeamName $TeamName -Clients $clients -DryRun $DryRun
} else {
    Write-Host "[Update-TeamFieldValues] Skipped by flag."
}

# ----------------------------
# (d) Add client groups to Team membership group
# Graph Descriptors - Get (projectId -> scope descriptor) [9](https://learn.microsoft.com/en-us/rest/api/azure/devops/wit/classification-nodes/get?view=azure-devops-rest-7.1)
# Graph Groups - List (scoped) [7](https://multishoring.com/blog/azure-logic-apps-vs-power-automate/)
# Graph Memberships - Add (PUTyoutube.com/watch?v=jqXss_jArtM)[16](https://learn.microsoft.com/en-us/azure/devops/pipelines/artifacts/pipeline-artifacts?view=azure-devops)
# ----------------------------
function Get-ProjectScopeDescriptor {
    param([string]$OrgName, [string]$ProjectId)

    Write-Host "[Graph] Resolving project scope descriptor..."
    $uri = "https://vssps.dev.azure.com/$OrgName/_apis/graph/descriptors/$ProjectId"
    $resp = Invoke-AdoRest -Method GET -Uri $uri
    return $resp.value
}

function Get-GraphGroupsInScope {
    param([string]$OrgName, [string]$ScopeDescriptor)

    Write-Host "[Graph] Listing groups in project scope (paged)..."
    $all = @()
    $continuation = $null

    do {
        $ctPart = if ($continuation) { "&continuationToken=$([uri]::EscapeDataString($continuation))" } else { "" }
        $uri = "https://vssps.dev.azure.com/$OrgName/_apis/graph/groups?scopeDescriptor=$([uri]::EscapeDataString($ScopeDescriptor))$ctPart&api-version=7.1-preview.1"

        # Use Invoke-WebRequest to access X-MS-ContinuationToken header. [7](https://multishoring.com/blog/azure-logic-apps-vs-power-automate/)
        $resp = Invoke-WebRequest -Method GET -Uri $uri -Headers $headers
        $json = $resp.Content | ConvertFrom-Json
        $all += @($json.value)

        $continuation = $resp.Headers.'X-MS-ContinuationToken'
    } while ($continuation)

    Write-Host "[Graph] Total groups retrieved: $($all.Count)"
    return $all
}

function Find-TeamGroupDescriptor {
    param([object[]]$Groups, [string]$TeamName)

    Write-Host "[Graph] Finding team membership group descriptor for '$TeamName'..."
    $g = $Groups | Where-Object { $_.displayName -eq $TeamName } | Select-Object -First 1
    if ($g) { return $g.descriptor }

    $g = $Groups | Where-Object { $_.principalName -like "*$TeamName*" } | Select-Object -First 1
    if ($g) { return $g.descriptor }

    return $null
}

function Find-ClientRoleGroupDescriptor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object[]] $Groups,
        [Parameter(Mandatory)] [string] $ClientName,
        [Parameter(Mandatory)] [string] $RoleName
    )

    $principal = "$ClientName_$RoleName"
    Write-Host "[Graph] Looking for group principalName: '$principal'"

    $g = $Groups | Where-Object { $_.principalName -eq $principal } | Select-Object -First 1
    if ($g) { return $g.descriptor }
    return $null
}

function Add-GraphMembershipIdempotent {
    param([string]$OrgName, [string]$SubjectDescriptor, [string]$ContainerDescriptor, [bool]$DryRun)

    $uri = "https://vssps.dev.azure.com/$OrgName/_apis/graph/memberships/$SubjectDescriptor/$ContainerDescriptor?api-version=7.1-preview.1"
    Write-Host "[Graph] Adding membership (subject -> team group)..."

    if ($DryRun -or -not $PSCmdlet.ShouldProcess($ContainerDescriptor, "PUT Graph Membership")) {
        Write-Host "[Graph] DryRun/WhatIf: would PUT $uri"
        return
    }

    try {
        # Memberships - Add is PUT. [8](https://www.youtube.com/watch?v=jqXss_jArtM)[16](https://learn.microsoft.com/en-us/azure/devops/pipelines/artifacts/pipeline-artifacts?view=azure-devops)
        Invoke-AdoRest -Method PUT -Uri $uri | Out-Null
        Write-Host "[Graph] Membership added/ensured."
    } catch {
        $code = Get-HttpStatusCode $_
        $body = Get-HttpErrorBody $_

        # Treat "already exists" as success (idempotent)
        if ($code -eq 409 -or ($body -match 'already exists')) {
            Write-Host "[Graph] Membership already exists (idempotent)."
            return
        }
        Write-Warning "[Graph] Membership add failed (HTTP $code). Body: $body"
        throw
    }
}

if (-not $SkipTeamMembershipGroups) {
    try {
        $scopeDesc = Get-ProjectScopeDescriptor -OrgName $orgName -ProjectId $projectId
        $graphGroups = Get-GraphGroupsInScope -OrgName $orgName -ScopeDescriptor $scopeDesc

        $teamGroupDesc = Find-TeamGroupDescriptor -Groups $graphGroups -TeamName $TeamName
        if (-not $teamGroupDesc) {
            Write-Warning "Team membership group not found for '$TeamName'. Skipping membership updates."
        } else {
            foreach ($c in $clients) {
                foreach ($r in $roles) {
                        $roleDesc = Find-ClientRoleGroupDescriptor -Groups $graphGroups -ClientName $c -RoleName $r
                        if (-not $roleDesc) {
                            Write-Warning "Group not found: '$c $r'. Skipping."
                            continue
                        }
                
                        Add-GraphMembershipIdempotent `
                            -OrgName $orgName `
                            -SubjectDescriptor $roleDesc `
                            -ContainerDescriptor $teamGroupDesc `
                            -DryRun $DryRun
                
                        Write-Host "Ensured: $c $r is in Team '$TeamName'"
                    }
            }
        }
    } catch {
        Write-Warning "Team membership group configuration encountered an error: $($_.Exception.Message)"
        # Decide if you want this to fail. For rollout bootstrap, many teams prefer continuing.
        # throw
    }
} else {
    Write-Host "[Graph] Skipped by flag."
}

# ----------------------------
# (e) Assign iterations to team (current date forward)
# - Classification Nodes - Get Iterations tree [12](https://developercommunity.visualstudio.com/t/Graph-APIs-not-working-for-Azure-DevOps/10975063?viewtype=all&stateGroup=active&ftype=problem)
# - Iterations - List [11](https://oshamrai.wordpress.com/2025/03/30/azure-devops-rest-api-python-8-manage-areas-and-iterations-in-team-projects/)
# - Iterations - Post Team Iteration [10](https://github.com/MicrosoftDocs/azure-devops-docs/blob/main/docs/integrate/get-started/rest/samples.md)
# ----------------------------
function Resolve-YearAndFromDate {
    param([int]$YearOfIteration)

    $now = Get-Date
    $currentYear = $now.Year
    $year = if ($YearOfIteration -eq 0) { $currentYear } else { $YearOfIteration }

    $fromDate = if ($year -eq $currentYear) { $now.Date } else { (Get-Date -Year $year -Month 1 -Day 1).Date }

    return [pscustomobject]@{
        Year     = $year
        YearName = $year.ToString()
        FromUtc  = $fromDate.ToUniversalTime()
        FromDate = $fromDate
    }
}

function Get-IterationTree {
    param([string]$Org, [string]$ProjectEsc)

    # Classification Nodes - Get with $depth. [12](https://developercommunity.visualstudio.com/t/Graph-APIs-not-working-for-Azure-DevOps/10975063?viewtype=all&stateGroup=active&ftype=problem)
    $uri = "$Org/$ProjectEsc/_apis/wit/classificationnodes/Iterations?`$depth=4"
    return Invoke-AdoRest -Method GET -Uri $uri
}

function Get-TeamIterationIds {
    param([string]$Org, [string]$ProjectEsc, [string]$TeamEsc)

    $uri = "$Org/$ProjectEsc/$TeamEsc/_apis/work/teamsettings/iterations"
    $resp = Invoke-AdoRest -Method GET -Uri $uri

    # Some responses use 'values' (docs) and some use 'value' in practice; handle both safely. [11](https://oshamrai.wordpress.com/2025/03/30/azure-devops-rest-api-python-8-manage-areas-and-iterations-in-team-projects/)
    $items = @()
    if ($resp.PSObject.Properties.Name -contains 'values') { $items = @($resp.values) }
    elseif ($resp.PSObject.Properties.Name -contains 'value') { $items = @($resp.value) }

    return @($items | ForEach-Object { $_.id })
}

function Add-TeamIteration {
    param([string]$Org, [string]$ProjectEsc, [string]$TeamEsc, [string]$IterationId, [bool]$DryRun)

    # Post Team Iteration expects { id: <uuid> }. [10](https://github.com/MicrosoftDocs/azure-devops-docs/blob/main/docs/integrate/get-started/rest/samples.md)
    $uri = "$Org/$ProjectEsc/$TeamEsc/_apis/work/teamsettings/iterations"
    if ($DryRun -or -not $PSCmdlet.ShouldProcess($TeamEsc, "POST Team Iteration $IterationId")) {
        Write-Host "[Iterations] DryRun/WhatIf: would POST $uri with id=$IterationId"
        return
    }
    Invoke-AdoRest -Method POST -Uri $uri -Body @{ id = $IterationId } | Out-Null
}

if (-not $SkipIterationAssignment) {
    $yr = Resolve-YearAndFromDate -YearOfIteration $YearOfIteration
    Write-Host "[Iterations] Target year: $($yr.YearName). FromDate: $($yr.FromDate)"

    try {
        $tree = Get-IterationTree -Org $Organization -ProjectEsc $projectEsc
        $yearNode = @($tree.children) | Where-Object { $_.name -eq $yr.YearName } | Select-Object -First 1

        if (-not $yearNode) {
            Write-Warning "Year iteration '$($yr.YearName)' not found under project iterations. Skipping iteration assignment (no failure)."
        } else {
            $sprints = @($yearNode.children)
            if ($sprints.Count -eq 0) {
                Write-Host "[Iterations] No sprints found under year '$($yr.YearName)'. Nothing to assign."
            } else {
                # Filter sprints from current date forward
                $toAssign = @()
                foreach ($s in $sprints) {
                    if ($s.attributes -and $s.attributes.startDate) {
                        $sd = [datetime]$s.attributes.startDate
                        if ($sd.ToUniversalTime() -ge $yr.FromUtc) { $toAssign += $s }
                    } else {
                        $toAssign += $s
                    }
                }

                Write-Host "[Iterations] Sprints under year: $($sprints.Count). After filter: $($toAssign.Count)"

                $assigned = @()
                try {
                    $assigned = Get-TeamIterationIds -Org $Organization -ProjectEsc $projectEsc -TeamEsc $teamEsc
                } catch {
                    Write-Warning "[Iterations] Could not list already-assigned iterations; will attempt to assign anyway."
                }
                $assignedNorm = @($assigned | ForEach-Object { $_.ToString().ToLowerInvariant() })

                foreach ($s in $toAssign) {
                    $iterId = $s.identifier  # GUID from classification node [12](https://developercommunity.visualstudio.com/t/Graph-APIs-not-working-for-Azure-DevOps/10975063?viewtype=all&stateGroup=active&ftype=problem)
                    $iterName = $s.name
                    $iterNorm = $iterId.ToString().ToLowerInvariant()

                    if ($assignedNorm -contains $iterNorm) {
                        Write-Host "[Iterations] Skipping already assigned: $iterName"
                        continue
                    }

                    try {
                        Add-TeamIteration -Org $Organization -ProjectEsc $projectEsc -TeamEsc $teamEsc -IterationId $iterId -DryRun $DryRun
                        Write-Host "✅ Assigned iteration: $iterName"
                    } catch {
                        Write-Warning "Failed assigning iteration '$iterName'. Error: $($_.Exception.Message)"
                    }
                }
            }
        }
    } catch {
        Write-Warning "[Iterations] Iteration assignment encountered an error: $($_.Exception.Message)"
        # Choose whether to fail; for rollout bootstrap, often continue.
        # throw
    }
} else {
    Write-Host "[Iterations] Skipped by flag."
}

Write-Host "`n✅ Done (bootstrap-team.ps1)."
