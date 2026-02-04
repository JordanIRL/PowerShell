[CmdletBinding()]
param(
    [string]$ExportPath = ".\LicenseAudit_$(Get-Date -Format 'dd-MM-yyyy_HH-mm').csv"
)

# To Add/Remove Licenses, use the GUID's from: 
# https://learn.microsoft.com/en-ie/entra/identity/users/licensing-service-plan-reference 
# Format as: "GUID" = "Product name"  | (String ID not used)

$LicenseSkuMapping = @{
    "05e9a617-0261-4cee-bb44-138d3ef5d965" = "Microsoft 365 E3"
    "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
    "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
    "c2fe850d-fbbb-4858-b67d-bd0c6e746da3" = "Microsoft 365 E3 EEA (no Teams)"
    "3271cf8e-2be5-4a09-a549-70fd05baaa17" = "Microsoft 365 E5 EEA (no Teams)"
    "d711d25a-a21c-492f-bd19-aae1e8ebaf30" = "Office 365 E3 EEA (no Teams)"
}

if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
    Write-Error "Microsoft.Graph not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"; exit 1
}
Import-Module Microsoft.Graph.Users -ErrorAction Stop

$ctx = Get-MgContext -ErrorAction SilentlyContinue
if (-not $ctx) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome
}
Write-Host "Connected as: $((Get-MgContext).Account)" -ForegroundColor Green

try {
    Write-Host "Retrieving licensed users..." -ForegroundColor Cyan
    $users = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,Mail,AssignedLicenses,AccountEnabled,SignInActivity" |
    Where-Object { $_.AssignedLicenses.Count -gt 0 }
    Write-Host "Found $($users.Count) licensed users." -ForegroundColor Green

    $allLicenseNames = @{}
    $userData = foreach ($user in $users) {
        $licenses = foreach ($lic in $user.AssignedLicenses) {
            $skuId = $lic.SkuId.ToString()
            if ($LicenseSkuMapping.ContainsKey($skuId)) {
                $LicenseSkuMapping[$skuId]
            }
        }
        if ($licenses) {
            $lastSignIn = $user.SignInActivity.LastSignInDateTime

            @{
                DisplayName           = $user.DisplayName
                Email                 = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
                AccountEnabled        = $user.AccountEnabled
                LastInteractiveSignIn = if ($lastSignIn) { $lastSignIn.ToString("dd-MM-yyyy HH:mm") } else { "Never" }
                Licenses              = $licenses
            }
            foreach ($l in $licenses) { $allLicenseNames[$l] = $true }
        }
    }

    $sortedLicenses = $allLicenseNames.Keys | Sort-Object
    $report = foreach ($u in $userData) {
        $row = [ordered]@{ DisplayName = $u.DisplayName; Email = $u.Email; AccountEnabled = $u.AccountEnabled; LastInteractiveSignIn = $u.LastInteractiveSignIn }
        $count = 0
        foreach ($lic in $sortedLicenses) {
            $has = $u.Licenses -contains $lic
            $row[$lic] = if ($has) { $count++; "Y" } else { "" }
        }
        $row["LicenseCount"] = $count
        [PSCustomObject]$row
    }

    if ($report.Count -gt 0) {
        $report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported to: $ExportPath" -ForegroundColor Green
    }
}
catch {
    Write-Error "Error: $_ `n$($_.ScriptStackTrace)"
}
