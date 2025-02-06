<#
  PIM-PAM Script with Already-Active Check & Error Handling
  ---------------------------------------------------------
  - Checks if user is already active in the selected role (RoleAssignmentExists).
  - Catches "PendingRoleAssignmentRequest" if there's an existing pending request.
  - If either condition occurs, we display an error in red, wait 5 seconds in yellow, disconnect,
    and let the user press Enter to exit.
  - If activation succeeds, it shows success, then does the usual 5-second wait, disconnect, etc.
#>

###############################################################################
# 1) Utility: Flush-ConsoleInput
###############################################################################
function Flush-ConsoleInput {
    while ([Console]::KeyAvailable) {
        [Console]::ReadKey($true) | Out-Null
    }
}

###############################################################################
# 2) Basic Startup & Connect to Microsoft Graph
###############################################################################
Write-Host "`nPIM-PAM script is running..." -ForegroundColor Green
Start-Sleep -Seconds 2

# Ensure PowerShell window stays open when double-clicked
if (-not $Host.Name -match "ConsoleHost") {
    Write-Host "Forcing console mode..."
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -NoNewWindow
    exit
}

try {
    Connect-MgGraph -NoWelcome
    Write-Host "Connected to Microsoft Graph!" -ForegroundColor Cyan
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    Write-Host "Press Enter to exit..."
    Read-Host | Out-Null
    exit
}

###############################################################################
# 3) Get Current User & Eligible Roles
###############################################################################
$context     = Get-MgContext
$currentUser = (Get-MgUser -UserId $context.Account).Id

# Get all available roles for the user
$myRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty RoleDefinition -All -Filter "principalId eq '$currentUser'"

# Check if roles exist and ensure they have valid names
$validRoles = $myRoles | Where-Object { $_.RoleDefinition -and $_.RoleDefinition.DisplayName }

if ($validRoles.Count -eq 0) {
    Write-Host "You do not have any eligible roles for activation, or role data is missing."
    Write-Host "Press Enter to exit..."
    Read-Host | Out-Null
    exit
}

Write-Host "`nAvailable Roles for Activation:`n"
$index   = 1
$roleMap = @{}
foreach ($role in $validRoles) {
    Write-Host ("[{0}] {1}" -f $index, $role.RoleDefinition.DisplayName)
    $roleMap[$index] = $role
    $index++
}

###############################################################################
# 4) Prompt User for Role Selection
###############################################################################
do {
    $selection = Read-Host "`nEnter the number corresponding to the role you want to activate"
    if ([string]::IsNullOrWhiteSpace($selection) -or -not ($selection -match '^\d+$') -or [int]$selection -le 0 -or [int]$selection -gt $roleMap.Count) {
        Write-Host "ERROR: Invalid selection. Please enter a valid number from the list." -ForegroundColor Red
        $selection = $null
    }
} while (-not $selection)

Flush-ConsoleInput

$myRole = $roleMap[[int]$selection]

###############################################################################
# 5) Prompt for Activation Duration
###############################################################################
do {
    $durationInput = Read-Host "`nEnter duration for activation (e.g., 1H for 1 hour, 30M for 30 minutes, 2H30M for 2 hours 30 minutes)"
    if ([string]::IsNullOrWhiteSpace($durationInput) -or $durationInput -notmatch '^\d+[HM]?$') {
        Write-Host "ERROR: Invalid duration format. Use '1H' for 1 hour or '30M' for 30 minutes." -ForegroundColor Red
        $durationInput = $null
    }
} while (-not $durationInput)

# Convert ISO 8601 duration (PTxHxM) format
$duration = $durationInput.ToUpper() -replace '(\d+)H', 'PT${1}H' -replace '(\d+)M', '${1}M'
if ($duration -match '^\d+M$') { 
    $duration = "PT$duration" 
}

###############################################################################
# 6) Prompt for Justification
###############################################################################
do {
    $justification = Read-Host "`nEnter the reason for activation"
    if ([string]::IsNullOrWhiteSpace($justification)) {
        Write-Host "ERROR: Justification cannot be empty." -ForegroundColor Red
        $justification = $null
    }
} while (-not $justification)

Flush-ConsoleInput

###############################################################################
# 7) Check if Role Already Active
###############################################################################
try {
    $activeRoleAssignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($myRole.PrincipalId)' and roleDefinitionId eq '$($myRole.RoleDefinitionId)'"
    if ($activeRoleAssignments) {
        Write-Host "`nERROR: [RoleAssignmentExists] : The Role assignment already exists (already active)." -ForegroundColor Red
        
        Write-Host "`nYou will be disconnected from Microsoft Graph in 5 seconds..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
        Disconnect-MgGraph | Out-Null
        Write-Host "`nYou have been disconnected from Microsoft Graph!" -ForegroundColor Red
        
        Write-Host "`nPress Enter to exit..." -ForegroundColor Cyan
        Read-Host | Out-Null
        exit
    }
}
catch {
    Write-Host "Warning: Could not check if user is already active. Continuing..."
}

###############################################################################
# 8) Prepare the Activation Request
###############################################################################
# Handle DirectoryScopeId: Use "/" for tenant-wide roles
$directoryScopeId = if ($myRole.DirectoryScopeId -eq $null -or $myRole.DirectoryScopeId -eq "") { "/" } else { $myRole.DirectoryScopeId }

$params = @{
    Action            = "selfActivate"
    PrincipalId       = $myRole.PrincipalId
    RoleDefinitionId  = $myRole.RoleDefinitionId
    DirectoryScopeId  = $directoryScopeId
    Justification     = $justification
    ScheduleInfo      = @{
        StartDateTime = Get-Date
        Expiration    = @{
            Type     = "AfterDuration"
            Duration = $duration
        }
    }
}

###############################################################################
# 9) Attempt Activation & Catch Known Errors
###############################################################################
[bool]$activationSuccess = $false
try {
    New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params | Out-Null
    $activationSuccess = $true
}
catch {
    $errorMsg = $_.Exception.Message

    if ($errorMsg -match "RoleAssignmentExists") {
        Write-Host "`nERROR: [RoleAssignmentExists] : The Role assignment already exists." -ForegroundColor Red
    }
    elseif ($errorMsg -match "PendingRoleAssignmentRequest") {
        Write-Host "`nERROR: [PendingRoleAssignmentRequest] : There is already an existing pending role assignment request." -ForegroundColor Red
    }
    else {
        Write-Host "`nERROR activating role: $errorMsg" -ForegroundColor Red
    }

    Write-Host "`nPress Enter to exit..."
    Read-Host | Out-Null
    exit
}

###############################################################################
# 10) Activation Success => Show Messages & Disconnect
###############################################################################
if ($activationSuccess) {
    Write-Host "`nRole activation request submitted successfully!" -ForegroundColor Green
    Write-Host "Activated Role: $($myRole.RoleDefinition.DisplayName)" -ForegroundColor DarkMagenta

    Write-Host "`nYou will be disconnected from Microsoft Graph in 5 seconds..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    Disconnect-MgGraph | Out-Null
    Write-Host "`nYou have been disconnected from Microsoft Graph!" -ForegroundColor Red

    Write-Host "`nPress Enter to exit..." -ForegroundColor Cyan
    Read-Host | Out-Null
    exit
}
