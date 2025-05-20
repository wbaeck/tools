#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Exports Intune Configuration Profiles, Remediation Scripts, and Platform Scripts for Windows 10/11 devices.
.DESCRIPTION
    This script connects to Microsoft Graph, queries Intune for all Windows 10/11 related configurations,
    and exports them to CSV and HTML reports for documentation purposes.
.NOTES
    File Name  : Export-IntuneWindowsConfigurations.ps1
    Author     : Generated script for Intune Admin
    Requires   : Microsoft Graph PowerShell SDK modules
#>

# Function to create folder if it doesn't exist
function Ensure-FolderExists {
    param (
        [string]$Path
    )
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Host "Created folder: $Path" -ForegroundColor Green
    }
}

# Function to format assignment data
function Format-AssignmentData {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Assignments
    )

    $formattedAssignments = @()

    foreach ($assignment in $Assignments) {
        $target = 'Unknown'
        $groupName = 'N/A'
        $intent = $assignment.intent

        if ($assignment.target.groupId) {
            try {
                $group = Get-MgGroup -GroupId $assignment.target.groupId -ErrorAction SilentlyContinue
                if ($group) {
                    $groupName = $group.DisplayName
                }
            }
            catch {
                $groupName = 'Unable to resolve group name'
            }

            if ($assignment.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget') {
                $target = "Group: $groupName (Include)"
            }
            elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
                $target = "Group: $groupName (Exclude)"
            }
        }
        elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allDevicesAssignmentTarget') {
            $target = 'All Devices'
        }
        elseif ($assignment.target.'@odata.type' -eq '#microsoft.graph.allLicensedUsersAssignmentTarget') {
            $target = 'All Users'
        }

        $formattedAssignments += [PSCustomObject]@{
            Target = $target
            Intent = $intent
        }
    }

    return $formattedAssignments
}

# Connect to Microsoft Graph
try {
    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
    Connect-MgGraph -Scopes 'DeviceManagementConfiguration.Read.All', 'DeviceManagementManagedDevices.Read.All', 'DeviceManagementApps.Read.All', 'Group.Read.All' -ErrorAction Stop

    # Check if we're connected
    $graphContext = Get-MgContext
    if (-not $graphContext) {
        throw 'Failed to connect to Microsoft Graph.'
    }

    Write-Host 'Successfully connected to Microsoft Graph' -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit 1
}

# Create output directory
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$outputDir = "IntuneExport-$timestamp"
Ensure-FolderExists -Path $outputDir

# 1. Export Device Configuration Profiles for Windows 10/11
Write-Host 'Exporting Windows 10/11 Configuration Profiles...' -ForegroundColor Cyan

$configProfiles = @()

# a) Get traditional device configurations
try {
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations'
    $devConfigs = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    foreach ($profile in $devConfigs.value) {
        # Include all Windows profiles
        if ($profile.'@odata.type' -match '#microsoft.graph.windows' -or
            $profile.platformType -eq 'windows10' -or
            $profile.platforms -contains 'windows10' -or
            $profile.deviceManagementApplicabilityRuleOsVersion.osMaximumVersion -match '10\.' -or
            $profile.deviceManagementApplicabilityRuleOsVersion.osMinimumVersion -match '10\.') {

            # Get assignments
            $assignmentUri = "$uri/$($profile.id)/assignments"
            $assignmentsResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -OutputType PSObject
            $assignments = $assignmentsResponse.value
            $formattedAssignments = Format-AssignmentData -Assignments $assignments

            $configProfiles += [PSCustomObject]@{
                Name                 = $profile.displayName
                Description          = $profile.description
                Type                 = $profile.'@odata.type' -replace '#microsoft.graph.', ''
                Platform             = 'Windows 10/11'
                CreatedDateTime      = $profile.createdDateTime
                LastModifiedDateTime = $profile.lastModifiedDateTime
                ID                   = $profile.id
                AssignmentCount      = $assignments.Count
                Assignments          = ($formattedAssignments | ConvertTo-Json -Compress)
            }
        }
    }
}
catch {
    Write-Host "Error retrieving device configurations: $_" -ForegroundColor Yellow
}

# b) Get configuration policies (Settings Catalog)
try {
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies'
    $configPolicies = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    foreach ($policy in $configPolicies.value) {
        # Check if it's for Windows
        if ($policy.platforms -contains 'windows10') {
            # Get assignments
            $assignmentUri = "$uri/$($policy.id)/assignments"
            $assignmentsResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -OutputType PSObject
            $assignments = $assignmentsResponse.value
            $formattedAssignments = Format-AssignmentData -Assignments $assignments

            $configProfiles += [PSCustomObject]@{
                Name                 = $policy.name
                Description          = $policy.description
                Type                 = 'Settings Catalog'
                Platform             = 'Windows 10/11'
                CreatedDateTime      = $policy.createdDateTime
                LastModifiedDateTime = $policy.lastModifiedDateTime
                ID                   = $policy.id
                AssignmentCount      = $assignments.Count
                Assignments          = ($formattedAssignments | ConvertTo-Json -Compress)
            }
        }
    }
}
catch {
    Write-Host "Error retrieving configuration policies: $_" -ForegroundColor Yellow
}

# c) Get endpoint security policies
try {
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/templates'
    $templates = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    foreach ($template in $templates.value) {
        if ($template.platformType -eq 'windows10EndpointProtection') {
            # Get all policies created from this template
            $policiesUri = "https://graph.microsoft.com/beta/deviceManagement/intents?`$filter=templateId eq '$($template.id)'"
            $policies = Invoke-MgGraphRequest -Method GET -Uri $policiesUri -OutputType PSObject

            foreach ($policy in $policies.value) {
                # Get assignments
                $assignmentUri = "https://graph.microsoft.com/beta/deviceManagement/intents/$($policy.id)/assignments"
                $assignmentsResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -OutputType PSObject
                $assignments = $assignmentsResponse.value
                $formattedAssignments = Format-AssignmentData -Assignments $assignments

                $configProfiles += [PSCustomObject]@{
                    Name                 = $policy.displayName
                    Description          = $policy.description
                    Type                 = "Endpoint Security - $($template.displayName)"
                    Platform             = 'Windows 10/11'
                    CreatedDateTime      = $policy.createdDateTime
                    LastModifiedDateTime = $policy.lastModifiedDateTime
                    ID                   = $policy.id
                    AssignmentCount      = $assignments.Count
                    Assignments          = ($formattedAssignments | ConvertTo-Json -Compress)
                }
            }
        }
    }
}
catch {
    Write-Host "Error retrieving endpoint security policies: $_" -ForegroundColor Yellow
}

# d) Get Windows update policies
try {
    $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$filter=isof('microsoft.graph.windowsUpdateForBusinessConfiguration')"
    $updatePolicies = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    foreach ($policy in $updatePolicies.value) {
        # Get assignments
        $assignmentUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($policy.id)/assignments"
        $assignmentsResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -OutputType PSObject
        $assignments = $assignmentsResponse.value
        $formattedAssignments = Format-AssignmentData -Assignments $assignments

        $configProfiles += [PSCustomObject]@{
            Name                 = $policy.displayName
            Description          = $policy.description
            Type                 = 'Windows Update Policy'
            Platform             = 'Windows 10/11'
            CreatedDateTime      = $policy.createdDateTime
            LastModifiedDateTime = $policy.lastModifiedDateTime
            ID                   = $policy.id
            AssignmentCount      = $assignments.Count
            Assignments          = ($formattedAssignments | ConvertTo-Json -Compress)
        }
    }
}
catch {
    Write-Host "Error retrieving Windows update policies: $_" -ForegroundColor Yellow
}

# e) Get administrative templates (Group Policy)
try {
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations'
    $gpoPolicies = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    foreach ($policy in $gpoPolicies.value) {
        # Get assignments
        $assignmentUri = "$uri/$($policy.id)/assignments"
        $assignmentsResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -OutputType PSObject
        $assignments = $assignmentsResponse.value
        $formattedAssignments = Format-AssignmentData -Assignments $assignments

        $configProfiles += [PSCustomObject]@{
            Name                 = $policy.displayName
            Description          = $policy.description
            Type                 = 'Administrative Template (Group Policy)'
            Platform             = 'Windows 10/11'
            CreatedDateTime      = $policy.createdDateTime
            LastModifiedDateTime = $policy.lastModifiedDateTime
            ID                   = $policy.id
            AssignmentCount      = $assignments.Count
            Assignments          = ($formattedAssignments | ConvertTo-Json -Compress)
        }
    }
}
catch {
    Write-Host "Error retrieving administrative templates: $_" -ForegroundColor Yellow
}

# f) Get device compliance policies
try {
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies'
    $compliancePolicies = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    foreach ($policy in $compliancePolicies.value) {
        if ($policy.'@odata.type' -match '#microsoft.graph.windows' -or
            $policy.platformType -eq 'windows10' -or
            $policy.platforms -contains 'windows10') {

            # Get assignments
            $assignmentUri = "$uri/$($policy.id)/assignments"
            $assignmentsResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -OutputType PSObject
            $assignments = $assignmentsResponse.value
            $formattedAssignments = Format-AssignmentData -Assignments $assignments

            $configProfiles += [PSCustomObject]@{
                Name                 = $policy.displayName
                Description          = $policy.description
                Type                 = 'Compliance Policy'
                Platform             = 'Windows 10/11'
                CreatedDateTime      = $policy.createdDateTime
                LastModifiedDateTime = $policy.lastModifiedDateTime
                ID                   = $policy.id
                AssignmentCount      = $assignments.Count
                Assignments          = ($formattedAssignments | ConvertTo-Json -Compress)
            }
        }
    }
}
catch {
    Write-Host "Error retrieving compliance policies: $_" -ForegroundColor Yellow
}

Write-Host "Found $(($configProfiles).Count) Windows 10/11 configuration profiles and policies" -ForegroundColor Cyan

# Export Configuration Profiles to CSV
$configProfiles | Export-Csv -Path "$outputDir\Windows10_11_ConfigurationProfiles.csv" -NoTypeInformation
Write-Host "Exported $(($configProfiles).Count) Windows 10/11 configuration profiles to $outputDir\Windows10_11_ConfigurationProfiles.csv" -ForegroundColor Green

# 2. Export Proactive Remediations (Detect and Remediate Scripts)
Write-Host 'Exporting Proactive Remediation Scripts...' -ForegroundColor Cyan

$remediationScripts = @()

# Use Graph API directly since the cmdlet isn't working
try {
    $graphApiVersion = 'beta'
    $graphEndpoint = 'deviceManagement/deviceHealthScripts'
    $uri = "https://graph.microsoft.com/$graphApiVersion/$graphEndpoint"
    $allRemediations = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    foreach ($remediation in $allRemediations.value) {
        # Get assignments
        $assignmentUri = "https://graph.microsoft.com/$graphApiVersion/$graphEndpoint/$($remediation.id)/assignments"
        $assignmentsResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -OutputType PSObject
        $assignments = $assignmentsResponse.value
        $formattedAssignments = Format-AssignmentData -Assignments $assignments

        # Get script contents
        $detectionScript = ''
        $remediationScript = ''

        if ($remediation.detectionScriptContent) {
            $detectionScript = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($remediation.detectionScriptContent))
        }

        if ($remediation.remediationScriptContent) {
            $remediationScript = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($remediation.remediationScriptContent))
        }

        $remediationScripts += [PSCustomObject]@{
            Name                     = $remediation.displayName
            Description              = $remediation.description
            DetectionScriptContent   = $detectionScript
            RemediationScriptContent = $remediationScript
            CreatedDateTime          = $remediation.createdDateTime
            LastModifiedDateTime     = $remediation.lastModifiedDateTime
            RunAsAccount             = $remediation.runAsAccount
            EnforceSignatureCheck    = $remediation.enforceSignatureCheck
            RunAs32Bit               = $remediation.runAs32Bit
            ID                       = $remediation.id
            AssignmentCount          = $assignments.Count
            Assignments              = ($formattedAssignments | ConvertTo-Json -Compress)
        }
    }
}
catch {
    Write-Host "Error retrieving remediation scripts: $_" -ForegroundColor Red
    $remediationScripts = @()
}

# Export Remediation Scripts to CSV
$remediationScripts | Export-Csv -Path "$outputDir\Windows10_11_RemediationScripts.csv" -NoTypeInformation
Write-Host "Exported $(($remediationScripts).Count) remediation scripts to $outputDir\Windows10_11_RemediationScripts.csv" -ForegroundColor Green

# 3. Export PowerShell Scripts (Platform Scripts)
Write-Host 'Exporting PowerShell Scripts (Platform Scripts)...' -ForegroundColor Cyan

$powerShellScripts = @()

# Use Graph API directly since the cmdlet isn't working
try {
    $graphApiVersion = 'beta'
    $graphEndpoint = 'deviceManagement/deviceManagementScripts'
    $uri = "https://graph.microsoft.com/$graphApiVersion/$graphEndpoint"
    $allPowerShellScripts = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

    foreach ($script in $allPowerShellScripts.value) {
        # Get assignments
        $assignmentUri = "https://graph.microsoft.com/$graphApiVersion/$graphEndpoint/$($script.id)/assignments"
        $assignmentsResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -OutputType PSObject
        $assignments = $assignmentsResponse.value
        $formattedAssignments = Format-AssignmentData -Assignments $assignments

        # Get script content
        $scriptContent = ''
        if ($script.scriptContent) {
            $scriptContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($script.scriptContent))
        }

        $powerShellScripts += [PSCustomObject]@{
            Name                  = $script.displayName
            Description           = $script.description
            ScriptContent         = $scriptContent
            CreatedDateTime       = $script.createdDateTime
            LastModifiedDateTime  = $script.lastModifiedDateTime
            RunAsAccount          = $script.runAsAccount
            EnforceSignatureCheck = $script.enforceSignatureCheck
            RunAs32Bit            = $script.runAs32Bit
            ID                    = $script.id
            AssignmentCount       = $assignments.Count
            Assignments           = ($formattedAssignments | ConvertTo-Json -Compress)
        }
    }
}
catch {
    Write-Host "Error retrieving PowerShell scripts: $_" -ForegroundColor Red
    $powerShellScripts = @()
}

# Export PowerShell Scripts to CSV
$powerShellScripts | Export-Csv -Path "$outputDir\Windows10_11_PowerShellScripts.csv" -NoTypeInformation
Write-Host "Exported $(($powerShellScripts).Count) PowerShell scripts to $outputDir\Windows10_11_PowerShellScripts.csv" -ForegroundColor Green

# 4. Generate HTML Reports
Write-Host 'Generating HTML Reports...' -ForegroundColor Cyan

# Function to create HTML report
function Create-HTMLReport {
    param (
        [string]$Title,
        [array]$Data,
        [string]$FilePath
    )

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>$Title</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0078D4; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #0078D4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .timestamp { color: #666; font-size: 0.8em; margin-bottom: 20px; }
        .section { margin-top: 40px; }
    </style>
</head>
<body>
    <h1>$Title</h1>
    <div class="timestamp">Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
"@

    if ($Data.Count -eq 0) {
        $html += '<p>No data found.</p>'
    }
    else {
        $html += '<table>'
        $html += '<tr>'
        $properties = $Data[0].PSObject.Properties.Name | Where-Object { $_ -ne 'ScriptContent' -and $_ -ne 'DetectionScriptContent' -and $_ -ne 'RemediationScriptContent' }
        foreach ($prop in $properties) {
            $html += "<th>$prop</th>"
        }
        $html += '</tr>'

        foreach ($item in $Data) {
            $html += '<tr>'
            foreach ($prop in $properties) {
                $value = $item.$prop
                if ($prop -eq 'Assignments') {
                    try {
                        $assignmentObj = $value | ConvertFrom-Json
                        $assignmentText = $assignmentObj | ForEach-Object { "• $($_.Target) ($($_.Intent))" }
                        $html += '<td>' + ($assignmentText -join '<br>') + '</td>'
                    }
                    catch {
                        $html += "<td>$value</td>"
                    }
                }
                else {
                    $html += "<td>$value</td>"
                }
            }
            $html += '</tr>'
        }
        $html += '</table>'
    }

    # Add script content sections for scripts
    if ($Title -match 'Script') {
        foreach ($item in $Data) {
            $html += "<div class='section'>"
            $html += "<h2>$($item.Name)</h2>"

            if ($item.PSObject.Properties.Name -contains 'DetectionScriptContent') {
                $html += '<h3>Detection Script</h3>'
                $html += "<pre>$($item.DetectionScriptContent)</pre>"
                $html += '<h3>Remediation Script</h3>'
                $html += "<pre>$($item.RemediationScriptContent)</pre>"
            }
            else {
                $html += '<h3>Script Content</h3>'
                $html += "<pre>$($item.ScriptContent)</pre>"
            }

            $html += '</div>'
        }
    }

    $html += @'
</body>
</html>
'@

    $html | Out-File -FilePath $FilePath -Encoding utf8
}

# Generate HTML Reports
Create-HTMLReport -Title 'Windows 10/11 Configuration Profiles' -Data $configProfiles -FilePath "$outputDir\Windows10_11_ConfigurationProfiles.html"
Create-HTMLReport -Title 'Windows 10/11 Remediation Scripts' -Data $remediationScripts -FilePath "$outputDir\Windows10_11_RemediationScripts.html"
Create-HTMLReport -Title 'Windows 10/11 PowerShell Scripts' -Data $powerShellScripts -FilePath "$outputDir\Windows10_11_PowerShellScripts.html"

Write-Host "Generated HTML reports in $outputDir" -ForegroundColor Green

# 5. Generate Summary Report
$summaryHtml = @"
<!DOCTYPE html>
<html>
<head>
    <title>Intune Windows 10/11 Configuration Summary</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0078D4; }
        h2 { color: #0078D4; margin-top: 30px; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #0078D4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .timestamp { color: #666; font-size: 0.8em; margin-bottom: 20px; }
        .card { background-color: #f9f9f9; border-radius: 5px; padding: 15px; margin-bottom: 20px; }
        .count { font-size: 2em; font-weight: bold; color: #0078D4; }
        .flex-container { display: flex; flex-wrap: wrap; }
        .flex-item { flex: 1; min-width: 250px; margin: 10px; }
    </style>
</head>
<body>
    <h1>Intune Windows 10/11 Configuration Summary</h1>
    <div class="timestamp">Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>

    <div class="flex-container">
        <div class="flex-item card">
            <div class="count">$(($configProfiles).Count)</div>
            <div>Configuration Profiles</div>
            <div><a href="Windows10_11_ConfigurationProfiles.html">View Details</a></div>
        </div>
        <div class="flex-item card">
            <div class="count">$(($remediationScripts).Count)</div>
            <div>Remediation Scripts</div>
            <div><a href="Windows10_11_RemediationScripts.html">View Details</a></div>
        </div>
        <div class="flex-item card">
            <div class="count">$(($powerShellScripts).Count)</div>
            <div>PowerShell Scripts</div>
            <div><a href="Windows10_11_PowerShellScripts.html">View Details</a></div>
        </div>
    </div>

    <h2>Top 10 Configuration Profiles</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Type</th>
            <th>Last Modified</th>
            <th>Assignment Count</th>
        </tr>
$(
    $configProfiles | Sort-Object -Property LastModifiedDateTime -Descending | Select-Object -First 10 | ForEach-Object {
        "<tr><td>$($_.Name)</td><td>$($_.Type)</td><td>$($_.LastModifiedDateTime)</td><td>$($_.AssignmentCount)</td></tr>"
    }
)
    </table>

    <h2>Top 10 Remediation Scripts</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Last Modified</th>
            <th>Assignment Count</th>
        </tr>
$(
    $remediationScripts | Sort-Object -Property LastModifiedDateTime -Descending | Select-Object -First 10 | ForEach-Object {
        "<tr><td>$($_.Name)</td><td>$($_.LastModifiedDateTime)</td><td>$($_.AssignmentCount)</td></tr>"
    }
)
    </table>

    <h2>All Files</h2>
    <ul>
        <li><a href="Windows10_11_ConfigurationProfiles.csv">Windows10_11_ConfigurationProfiles.csv</a></li>
        <li><a href="Windows10_11_ConfigurationProfiles.html">Windows10_11_ConfigurationProfiles.html</a></li>
        <li><a href="Windows10_11_RemediationScripts.csv">Windows10_11_RemediationScripts.csv</a></li>
        <li><a href="Windows10_11_RemediationScripts.html">Windows10_11_RemediationScripts.html</a></li>
        <li><a href="Windows10_11_PowerShellScripts.csv">Windows10_11_PowerShellScripts.csv</a></li>
        <li><a href="Windows10_11_PowerShellScripts.html">Windows10_11_PowerShellScripts.html</a></li>
    </ul>
</body>
</html>
"@

$summaryHtml | Out-File -FilePath "$outputDir\IntuneWindowsConfigSummary.html" -Encoding utf8
Write-Host "Generated summary report: $outputDir\IntuneWindowsConfigSummary.html" -ForegroundColor Green

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null
Write-Host 'Disconnected from Microsoft Graph' -ForegroundColor Cyan

# Final summary
Write-Host "`nExport Complete!" -ForegroundColor Green
Write-Host "All reports have been saved to: $((Get-Item -Path $outputDir).FullName)" -ForegroundColor Yellow
Write-Host "Open $outputDir\IntuneWindowsConfigSummary.html to view the summary report." -ForegroundColor Yellow