[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$Organization,   # https://dev.azure.com/your-org

    [Parameter(Mandatory = $true)]
    [string]$Project,

    [Parameter(Mandatory = $true)]
    [string]$TeamName,

    [Parameter(Mandatory = $true)]
    [string]$ClientsJson,

    [Parameter(Mandatory = $false)]
    [string]$RolesJson = '["Developers","Business Analysts","Benefits System Administrators","Project Managers"]',

    [Parameter(Mandatory = $false)]
    [bool]$DryRun = $false,

    [Parameter(Mandatory = $false)]
    [bool]$SkipTeamFieldValues = $false,

    [Parameter(Mandatory = $false)]
    [bool]$SkipTeamMembershipGroups = $false,

    [Parameter(Mandatory = $false)]
    [bool]$SkipIterationAssignment = $false
)

# =========================
# FUNCTIONS
# =========================
function Get-OAuthToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$Scope
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = $Scope
        grant_type    = "client_credentials"
    }

    (Invoke-RestMethod -Method POST -Uri $tokenUri -ContentType "application/x-www-form-urlencoded" -Body $body).access_token
}

function Invoke-AdoRest {
    param(
        [Parameter(Mandatory)][ValidateSet("GET","POST","PATCH","PUT","DELETE")] [string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()] $Body
    )

    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 50
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $AdoHeaders -Body $json -ContentType "application/json"
    }
    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $AdoHeaders
}

function Invoke-MsGraph {
    param(
        [Parameter(Mandatory)][ValidateSet("GET","POST","PATCH","PUT","DELETE")] [string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()] $Body
    )

    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 50
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $GraphHeaders -Body $json -ContentType "application/json"
    }
    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $GraphHeaders
}

function Get-OrgNameFromUrl {
    param([string]$OrgUrl)
    if ($OrgUrl -match '^https://dev\.azure\.com/([^/]+)') { return $Matches[1] }
    throw "Organization must be https://dev.azure.com/{org}. Got: $OrgUrl"
}

function Get-ProjectIdByName {
    param([string]$Org, [string]$ProjectName)

    if ($ProjectName -match '^[0-9a-fA-F-]{36}$') { return $ProjectName }

    $skip = 0; $top = 100
    while ($true) {
        $uri = "$Org/_apis/projects?`$top=$top&`$skip=$skip&api-version=$ApiVersionCore"  
        $resp = Invoke-AdoRest -Method GET -Uri $uri
        $match = @($resp.value) | Where-Object { $_.name -eq $ProjectName } | Select-Object -First 1
        if ($match) { return $match.id }
        if (-not $resp.value -or $resp.value.Count -lt $top) { break }
        $skip += $top
    }
    return $null
}

function Create-Team {
    param([string]$Org, [string]$ProjectId, [string]$TeamName, [bool]$DryRun)

    $listUri   = "$Org/_apis/projects/$ProjectId/teams?api-version=$ApiVersionCore"         # Teams - Get Teams [4](https://www.powershellgallery.com/packages/ado.core/1.0.20/Content/functions%5Cget-adoclassificationnode.ps1)
    $createUri = "$Org/_apis/projects/$ProjectId/teams?api-version=$ApiVersionCore"         # Teams - Create [5](https://www.reddit.com/r/azuredevops/comments/ii7rn0/difference_between_azure_artifacts_and_pipeline/)

    $teams = Invoke-AdoRest -Method GET -Uri $listUri
    $existing = @($teams.value) | Where-Object { $_.name -eq $TeamName } | Select-Object -First 1
    if ($existing) {
        Write-Host "[Team] Exists: $TeamName (id=$($existing.id))"
        return $existing
    }

    if ($DryRun -or -not $PSCmdlet.ShouldProcess($TeamName, "Create Team")) {
        Write-Host "[Team] DryRun/WhatIf: would create team '$TeamName'"
        return $null
    }

    $created = Invoke-AdoRest -Method POST -Uri $createUri -Body @{ name = $TeamName }
    Write-Host "[Team] Created: $TeamName (id=$($created.id))"
    return $created
}

function Test-AreaExists {
    param([string]$Org, [string]$ProjectEsc, [string[]]$FullPathSegments)

    $encodedPath = ($FullPathSegments | ForEach-Object { [uri]::EscapeDataString($_) }) -join "/"
    $uri = "$Org/$ProjectEsc/_apis/wit/classificationnodes/Areas/${encodedPath}?api-version=${ApiVersionWit}"   # Classification Nodesithub.com/MicrosoftDocs/azure-devops-docs/blob/main/docs/integrate/get-started/rest/samples.md)[2](https://medium.com/@kanerika/power-automate-vs-logic-apps-2025-full-comparison-of-microsoft-automation-tools-f569b42f2cea)

    try {
        Invoke-AdoRest -Method GET -Uri $uri | Out-Null
        return $true
    } catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 404) { return $false }
        throw
    }
}

function Create-AreaNode {
    param([string]$Org, [string]$ProjectEsc, [string[]]$ParentSegments, [string]$Name, [bool]$DryRun)

    $fullPath = @()
    if ($ParentSegments) { $fullPath += $ParentSegments }
    $fullPath += $Name

    if (Test-AreaExists -Org $Org -ProjectEsc $ProjectEsc -FullPathSegments $fullPath) {
        Write-Host "[Area] Exists: $($fullPath -join '\')"
        return
    }

    $parentPath = ""
    if ($ParentSegments -and $ParentSegments.Count -gt 0) {
        $parentPath = "/" + (($ParentSegments | ForEach-Object { [uri]::EscapeDataString($_) }) -join "/")
    }

    $uri = "$Org/$ProjectEsc/_apis/wit/classificationnodes/Areas${parentPath}?api-version=${ApiVersionWit}"   
    if ($DryRun -or -not $PSCmdlet.ShouldProcess(($fullPath -join '\'), "Create Area Node")) {
        Write-Host "[Area] DryRun/WhatIf: would create $($fullPath -join '\')"
        return
    }

    Invoke-AdoRest -Method POST -Uri $uri -Body @{ name = $Name } | Out-Null
    Write-Host "[Area] Created: $($fullPath -join '\')"
}

function Update-TeamFieldValues {
    param([string]$Org, [string]$ProjectEsc, [string]$TeamEsc, [string]$ProjectName, [string]$TeamName, [string[]]$Clients, [bool]$DryRun)

    $uri = "$Org/$ProjectEsc/$TeamEsc/_apis/work/teamsettings/teamfieldvalues?api-version=$ApiVersionWork"   
    $teamRoot = "$ProjectName\$TeamName"

    $values = @(@{ value = $teamRoot; includeChildren = $true })
    foreach ($c in $Clients) { $values += @{ value = "$ProjectName\$TeamName\$c"; includeChildren = $true } }

    $body = @{ defaultValue = $teamRoot; values = $values }

    if ($DryRun -or -not $PSCmdlet.ShouldProcess($TeamName, "PATCH TeamFieldValues")) {
        Write-Host "[TeamFieldValues] DryRun/WhatIf: would PATCH $uri"
        return
    }

    Invoke-AdoRest -Method PATCH -Uri $uri -Body $body | Out-Null
    Write-Host "[TeamFieldValues] Updated"
}

function Get-ProjectScopeDescriptor {
    param([string]$OrgName, [string]$ProjectId)

    $uri = "https://vssps.dev.azure.com/$OrgName/_apis/graph/descriptors/${ProjectId}?api-version=7.1"  
    (Invoke-AdoRest -Method GET -Uri $uri).value
}

function Get-GraphGroupsInScope {
    param([string]$OrgName, [string]$ScopeDescriptor)

    $all = @()
    $continuation = $null
    do {
        $ctPart = if ($continuation) { "&continuationToken=$([uri]::EscapeDataString($continuation))" } else { "" }
        $uri = "https://vssps.dev.azure.com/$OrgName/_apis/graph/groups?scopeDescriptor=$([uri]::EscapeDataString($ScopeDescriptor))${ctPart}&api-version=${ApiVersionGraphPreview}"  
        $resp = Invoke-WebRequest -Method GET -Uri $uri -Headers $AdoHeaders
        $json = $resp.Content | ConvertFrom-Json
        $all += @($json.value)
        $continuation = $resp.Headers.'X-MS-ContinuationToken'
    } while ($continuation)

    Write-Host "[Graph] Groups in scope: $($all.Count)"
    return $all
}

function Find-TeamGroupDescriptor {
    param([object[]]$Groups, [string]$TeamName)

    $g = $Groups | Where-Object { $_.displayName -eq $TeamName } | Select-Object -First 1
    if ($g) { return $g.descriptor }

    $g = $Groups | Where-Object { $_.principalName -like "*$TeamName*" } | Select-Object -First 1
    if ($g) { return $g.descriptor }

    return $null
}

function Find-ClientRoleGroupDescriptor {
    param([object[]]$Groups, [string]$ClientName, [string]$RoleName)

    # Your current pattern is "ClientA Developers" (space-separated)
    $principal = "$ClientName $RoleName"
    $g = $Groups | Where-Object { $_.principalName -eq $principal } | Select-Object -First 1
    if ($g) { return $g.descriptor }
    return $null
}

function Find-AadGroupObjectIdByDisplayName {
    param(
        [Parameter(Mandatory)][string]$DisplayName
    )

    Write-Host "GraphAccessToken length: $($GraphAccessToken.Length)"
    if ([string]::IsNullOrWhiteSpace($GraphAccessToken)) {
      throw "GraphAccessToken is empty. Token acquisition failed or variable scope is wrong."
    }

    
    Invoke-MsGraph -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$top=1"
    $payload = Get-JwtPayload -Jwt $GraphAccessToken
    Write-Host "aud  : $($payload.aud)"
    Write-Host "tid  : $($payload.tid)"
    Write-Host "appid: $($payload.appid)"
    Write-Host "roles: $($payload.roles -join ', ')"


    # Escape single quotes for OData:  '  ->  ''
    $safe = $DisplayName.Replace("'", "''")

    # Build query in a dictionary (no manual string escaping)
    $query = [ordered]@{
        '$filter' = "displayName eq '$safe'"
        '$select' = 'id,displayName'
    }
    $ub = [System.UriBuilder]::new("https://graph.microsoft.com/v1.0/groups")
    $ub.Query = ($query.GetEnumerator() | ForEach-Object {
        # URL-encode each value
        "{0}={1}" -f [uri]::EscapeDataString($_.Key), [uri]::EscapeDataString($_.Value)
    }) -join '&'
    
    $uri = $ub.Uri.AbsoluteUri
    Write-Host "[MS Graph] GET $uri"
    $resp = Invoke-MsGraph -Method GET -Uri $uri

    if (-not $resp.value -or $resp.value.Count -eq 0) { return $null }
    return $resp.value[0].id
}

function Materialize-AadGroupInAdoGraph {
    param(
        [Parameter(Mandatory)][string]$OrgName,
        [Parameter(Mandatory)][string]$AadObjectId,
        [Parameter(Mandatory)][string]$DisplayName
    )

    $uri = "https://vssps.dev.azure.com/$OrgName/_apis/graph/groups?api-version=7.1-preview.1"
    $body = @{
        "@odata.type" = "#Microsoft.VisualStudio.Services.Graph.GraphGroupOriginIdCreationContext"
        originId      = $AadObjectId
        displayName   = $DisplayName
    }

    return Invoke-AdoRest -Method POST -Uri $uri -Body $body
}

function Add-GraphMembershipIdempotent {
    param([string]$OrgName, [string]$SubjectDescriptor, [string]$ContainerDescriptor, [bool]$DryRun)

    $uri = "https://vssps.dev.azure.com/$OrgName/_apis/graph/memberships/$SubjectDescriptor/$ContainerDescriptor?api-version=$ApiVersionGraphPreview"  # Memberships Add [13](https://www.youtube.com/watch?v=RXQ9jZmKzfE)[14](https://michelcarlo.com/2024/07/27/how-to-create-azure-devops-pull-requests-using-power-automate/)

    if ($DryRun -or -not $PSCmdlet.ShouldProcess($ContainerDescriptor, "PUT Graph Membership")) {
        Write-Host "[Graph] DryRun/WhatIf: would add membership"
        return
    }

    try {
        Invoke-AdoRest -Method PUT -Uri $uri | Out-Null
        Write-Host "[Graph] Membership ensured"
    } catch {
        $body = $_.ErrorDetails.Message
        if ($body -match 'already exists') { return }
        throw
    }
}

function Resolve-YearAndFromDate {
    param([int]$YearOfIteration)
    $now = Get-Date
    $year = if ($YearOfIteration -eq 0) { $now.Year } else { $YearOfIteration }
    $fromDate = if ($year -eq $now.Year) { $now.Date } else { (Get-Date -Year $year -Month 1 -Day 1).Date }
    [pscustomobject]@{ YearName=$year.ToString(); FromUtc=$fromDate.ToUniversalTime(); FromDate=$fromDate }
}

function Get-IterationTree {
    param([string]$Org, [string]$ProjectEsc)
    $uri = "$Org/$ProjectEsc/_apis/wit/classificationnodes/Iterations?`$depth=4&api-version=$ApiVersionWit"  # Classification Nodes Get [6](https://github.com/MicrosoftDocs/azure-devops-docs/blob/main/docs/integrate/get-started/rest/samples.md)[2](https://medium.com/@kanerika/power-automate-vs-logic-apps-2025-full-comparison-of-microsoft-automation-tools-f569b42f2cea)
    Invoke-AdoRest -Method GET -Uri $uri
}

function Get-TeamIterationIds {
    param([string]$Org, [string]$ProjectEsc, [string]$TeamEsc)
    $uri = "$Org/$ProjectEsc/$TeamEsc/_apis/work/teamsettings/iterations?api-version=$ApiVersionWork"  # Iterations List [15](https://learn.microsoft.com/en-us/rest/api/azure/devops/graph/?view=azure-devops-rest-7.1)
    $resp = Invoke-AdoRest -Method GET -Uri $uri
    $items = @()
    if ($resp.PSObject.Properties.Name -contains 'values') { $items = @($resp.values) }
    elseif ($resp.PSObject.Properties.Name -contains 'value') { $items = @($resp.value) }
    @($items | ForEach-Object { $_.id })
}

function Add-TeamIteration {
    param([string]$Org, [string]$ProjectEsc, [string]$TeamEsc, [string]$IterationId, [bool]$DryRun)
    $uri = "$Org/$ProjectEsc/$TeamEsc/_apis/work/teamsettings/iterations?api-version=$ApiVersionWork"  # Post Team Iteration [16](https://learn.microsoft.com/en-us/rest/api/azure/devops/graph/groups?view=azure-devops-rest-7.1)
    if ($DryRun -or -not $PSCmdlet.ShouldProcess($TeamEsc, "Assign Iteration")) { return }
    Invoke-AdoRest -Method POST -Uri $uri -Body @{ id = $IterationId } | Out-Null
}

function Get-JwtPayload {
    param([Parameter(Mandatory)][string]$Jwt)

    $parts = $Jwt.Split('.')
    if ($parts.Count -lt 2) { throw "Not a JWT token." }

    $payload = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
    }

    $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
    return $json | ConvertFrom-Json
}

# =========================
# VARIABLES
# =========================
$Organization = $Organization.TrimEnd('/')
$ApiVersionCore = "7.1"
$ApiVersionWit  = "7.1"
$ApiVersionWork = "7.1"
$ApiVersionGraphPreview = "7.1-preview.1"
$TenantId     = $env:AAD_TENANT_ID
$ClientId     = $env:AAD_CLIENT_ID
$ClientSecret = $env:AAD_CLIENT_SECRET

Write-Host "aud   : $($payload.aud)"
Write-Host "appid : $($payload.appid)"
Write-Host "roles : $($payload.roles -join ', ')"

Write-Host "`n=== Bootstrap Team Script ==="
Write-Host "Organization: $Organization"
Write-Host "Project: $Project"
Write-Host "TeamName: $TeamName"
Write-Host "DryRun: $DryRun"
Write-Host "SkipTeamFieldValues: $SkipTeamFieldValues"
Write-Host "SkipTeamMembershipGroups: $SkipTeamMembershipGroups"
Write-Host "SkipIterationAssignment: $SkipIterationAssignment"

if ([string]::IsNullOrWhiteSpace($TenantId))     { throw "Missing AAD_TENANT_ID env var." }
if ([string]::IsNullOrWhiteSpace($ClientId))     { throw "Missing AAD_CLIENT_ID env var." }
if ([string]::IsNullOrWhiteSpace($ClientSecret)) { throw "Missing AAD_CLIENT_SECRET env var." }

Write-Host "AAD vars OK: TenantId=$TenantId, ClientId=$ClientId, SecretLength=$($ClientSecret.Length)"

# Graph token only (AAD lookup)
$GraphAccessToken = Get-OAuthToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -Scope "https://graph.microsoft.com/.default"

$payload = Get-JwtPayload -Jwt $GraphAccessToken

Write-Host "aud   : $($payload.aud)"
Write-Host "appid : $($payload.appid)"
Write-Host "roles : $($payload.roles -join ', ')"

$GraphHeaders = @{
    Authorization = "Bearer $GraphAccessToken"
    Accept        = "application/json"
}
# --- PAT / Auth ---
if (-not $env:AZURE_DEVOPS_EXT_PAT) {
    throw "Missing AZURE_DEVOPS_EXT_PAT. Set it as a secret pipeline variable and pass via env."
}

$pat = $env:AZURE_DEVOPS_EXT_PAT
$base64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))

$AdoHeaders = @{
    Authorization = "Basic $base64"
    Accept        = "application/json"
    "Content-Type"= "application/json"
}

$orgName = Get-OrgNameFromUrl -OrgUrl $Organization

# =========================
# PARSE CLIENTS / ROLES
# =========================
try { $clients = @($ClientsJson | ConvertFrom-Json) } catch { throw "ClientsJson must be JSON array. Got: $ClientsJson" }
$clients = $clients | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Select-Object -Unique
if ($clients.Count -lt 1) { throw "At least 1 client is required." }

try { $roles = @($RolesJson | ConvertFrom-Json) } catch { throw "RolesJson must be JSON array. Got: $RolesJson" }
$roles = $roles | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Select-Object -Unique
if ($roles.Count -lt 1) { throw "At least 1 role is required." }

Write-Host "Clients: $($clients -join ', ')"
Write-Host "Roles: $($roles -join ', ')"

# =========================
# CREATE TEAM
# =========================
$projectId = Get-ProjectIdByName -Org $Organization -ProjectName $Project
if (-not $projectId) { throw "Project '$Project' not found / not accessible." }

$projectEsc = [uri]::EscapeDataString($Project)
$teamEsc    = [uri]::EscapeDataString($TeamName)

$teamObj = Create-Team -Org $Organization -ProjectId $projectId -TeamName $TeamName -DryRun $DryRun

# =========================
# CREATE ROOT AREA AND AREA PATH PER CLIENT
# =========================

Create-AreaNode -Org $Organization -ProjectEsc $projectEsc -ParentSegments @()        -Name $TeamName -DryRun $DryRun
foreach ($c in $clients) {
    Create-AreaNode -Org $Organization -ProjectEsc $projectEsc -ParentSegments @($TeamName) -Name $c -DryRun $DryRun
}

# =========================
# WORK: TEAM FIELD VALUES (AreaPath)
# =========================

if (-not $SkipTeamFieldValues) {
    Update-TeamFieldValues -Org $Organization -ProjectEsc $projectEsc -TeamEsc $teamEsc -ProjectName $Project -TeamName $TeamName -Clients $clients -DryRun $DryRun
}

# =========================
# GRAPH: TEAM MEMBERSHIP GROUPS (ADO Graph)
# =========================

if (-not $SkipTeamMembershipGroups) {
#    try {
        $scopeDesc = Get-ProjectScopeDescriptor -OrgName $orgName -ProjectId $projectId
        $graphGroups = Get-GraphGroupsInScope -OrgName $orgName -ScopeDescriptor $scopeDesc

        $teamGroupDesc = Find-TeamGroupDescriptor -Groups $graphGroups -TeamName $TeamName
        if (-not $teamGroupDesc) {
            Write-Warning "Team membership group not found for '$TeamName'."
        } else {
            foreach ($c in $clients) {
                foreach ($r in $roles) {
                    
                    $aadName = "$c $r"   # e.g. "AdvocateAurora Developers"
                    $oid = Find-AadGroupObjectIdByDisplayName -DisplayName $aadName
                    Write-Host $oid

                    if (-not $oid) {
                        Write-Warning "AAD group not found: '$aadName' (skipping)"
                        continue
                    }

                    $mat = Materialize-AadGroupInAdoGraph -OrgName $orgName -AadObjectId $oid -DisplayName $aadName
                    if (-not $mat -or -not $mat.descriptor) {
                        Write-Warning "Failed to materialize '$aadName' into ADO Graph (skipping)"
                        continue
                    }

                    Add-GraphMembership -OrgName $orgName -SubjectDescriptor $mat.descriptor -ContainerDescriptor $teamGroupDesc
                    Write-Host "Added '$aadName' to Team '$TeamName'"

                }
            }
        }
  #  }
  #  catch {
  #      Write-Warning "Team membership group config error: $($_.Exception.Message)"
  #  }
}

# =========================
# WORK: ITERATION ASSIGNMENT (current date forward) — kept as-is; ensure api-version on calls
# =========================

if (-not $SkipIterationAssignment) {
    # If you later re-add YearOfIteration param, plug it here.
    # For now, default current-year behavior:
    $yr = Resolve-YearAndFromDate -YearOfIteration 0
    try {
        $tree = Get-IterationTree -Org $Organization -ProjectEsc $projectEsc
        $yearNode = @($tree.children) | Where-Object { $_.name -eq $yr.YearName } | Select-Object -First 1
        if (-not $yearNode) {
            Write-Warning "Year iteration '$($yr.YearName)' not found. Skipping iteration assignment."
        } else {
            $sprints = @($yearNode.children)
            $toAssign = @()
            foreach ($s in $sprints) {
                if ($s.attributes -and $s.attributes.startDate) {
                    $sd = [datetime]$s.attributes.startDate
                    if ($sd.ToUniversalTime() -ge $yr.FromUtc) { $toAssign += $s }
                } else {
                    $toAssign += $s
                }
            }

            $assigned = @()
            try { $assigned = Get-TeamIterationIds -Org $Organization -ProjectEsc $projectEsc -TeamEsc $teamEsc } catch { }
            $assignedNorm = @($assigned | ForEach-Object { $_.ToString().ToLowerInvariant() })

            foreach ($s in $toAssign) {
                $iterId = $s.identifier
                $iterName = $s.name
                if ($assignedNorm -contains $iterId.ToString().ToLowerInvariant()) { continue }
                Add-TeamIteration -Org $Organization -ProjectEsc $projectEsc -TeamEsc $teamEsc -IterationId $iterId -DryRun $DryRun
                Write-Host "Assigned: $iterName"
            }
        }
    } catch {
        Write-Warning "Iteration assignment error: $($_.Exception.Message)"
    }
}

Write-Host "`n✅ Done (bootstrap-team.ps1)"
