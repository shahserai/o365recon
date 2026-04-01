# o365recon - retrieve information on o365 accounts (and AzureAD too)
# Updated to use Microsoft Graph PowerShell SDK (2024)
#
# Original: 2021, 2018, 2017 @nyxgeek - TrustedSec
#   Special thanks to @spoonman1091 for ideas and contributions
# Updated: 2024 - Migrated from deprecated MSOnline/AzureAD modules to Microsoft Graph
#
# Requirements: 
# Open PowerShell window as Admin
# Then run "Install-Module Microsoft.Graph"
#
# Note: MSOnline and AzureAD modules are deprecated and will be retired.
#       Microsoft recommends using Microsoft Graph PowerShell SDK for all new development.
#
# Run the script. It will prompt you to authenticate. Log in. Get the loot.
#2024 : This fork introduces significant improvements to the PowerShell script responsible for retrieving and processing Office 365 user data. Previously, the script loaded all user data into memory, which could lead to excessive RAM usage, especially with large datasets. The revised version of the script addresses this issue by streaming user data directly to the output files without storing the entire dataset in memory.



param([switch] $azure = $false
      )


######################################################################################
# CONNECTING TO MICROSOFT SERVICES

echo "Connecting to Microsoft services:"

Write-Host -NoNewline "`t`t`tChecking for Microsoft.Graph Module ... "
if (Get-Module -ListAvailable -Name Microsoft.Graph) {
    Write-Host "`t`tDONE"
}else{
    Write-Host "`t`tFAILED"
    Write-Host "`tPlease install the Microsoft.Graph Module:`n`t`tInstall-Module Microsoft.Graph"
    exit
}

# Import the required Microsoft Graph modules
Write-Host -NoNewline "`t`t`tImporting Microsoft.Graph Modules ... "
try {
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Groups
    Import-Module Microsoft.Graph.Devices
    Import-Module Microsoft.Graph.Identity.DirectoryManagement
    Import-Module Microsoft.Graph.Applications
    Import-Module Microsoft.Graph.Organization
    Write-Host "`tDONE"
} catch {
    Write-Host "`tFAILED"
    Write-Host "`tError importing modules: $_"
    exit
}

# Connect to Microsoft Graph
try {
    Write-Host -NoNewline "`t`t`tConnecting to Microsoft Graph ... "
    $scopes = @(
        "User.Read.All",
        "Group.Read.All",
        "Device.Read.All",
        "Directory.Read.All",
        "Application.Read.All",
        "OrgContact.Read.All"
    )
    Connect-MgGraph -Scopes $scopes -NoWelcome
    Write-Host "`tDONE"
    
    # Verify connection
    $context = Get-MgContext
    if ($null -eq $context -or [string]::IsNullOrEmpty($context.TenantId)) {
        throw "Failed to establish a valid connection to Microsoft Graph"
    }
    Write-Host "`t`t`tConnected to Tenant: $($context.TenantId)"
    Write-Host "`t`t`tAuthenticated User: $($context.Account)"
} catch {
    Write-Host "`tFAILED"
    Write-Host "Could not connect to Microsoft Graph."
    throw $_
    exit
}


######################################################################################
# Setting up our working directory

[boolean]$pathIsOK = $false
$projectname = Read-host -prompt "Please enter a project name"
$inputclean = '[^a-zA-Z]'
$projectname = $projectname.Replace($inputclean,'')


while ($pathIsOK -eq $false){

    if (-not(Test-Path $projectname)){
        try{
            md $projectname > $null
            $CURRENTJOB = "./${projectname}/${projectname}"
            [boolean]$pathIsOK = $true
            }
        Catch{
            echo "whoops"
        }

    }else{
        $projectname = Read-host -prompt "File exists. Please enter a different project name"
        $inputclean = '[^a-zA-Z]'
        $projectname = $projectname.Replace($inputclean,'')
        [boolean]$pathIsOK = $false
    }

}



######################################################################################
# GET COMPANY AND DOMAIN INFO

Write-Host -NoNewline "`t`t`tRetrieving Company Info ... "
$organization = Get-MgOrganization
$organization | Select-Object -Property * |  Out-File -Append -FilePath .\${CURRENTJOB}.O365.CompanyInfo.txt
echo "`t`t`tDONE"


echo "------------------------------------------------------------------------------------"
echo "Overview" | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Company Name: $($organization.DisplayName)" | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Tenant ID: $($organization.Id)" | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
if ($organization.VerifiedDomains) {
    $initialDomain = ($organization.VerifiedDomains | Where-Object { $_.IsInitial -eq $true }).Name
    if (-not $initialDomain) {
        $initialDomain = ($organization.VerifiedDomains | Select-Object -First 1).Name
    }
    echo "Initial Domain: $initialDomain" | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
}
echo "Address: $($organization.Street), $($organization.City), $($organization.State) $($organization.PostalCode)" | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Phone Number: $($organization.PhoneNumber)" | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Technical Contact Addresses: $($organization.TechnicalNotificationMails -join ', ')" | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Marketing Contact Addresses: $($organization.MarketingNotificationMails -join ', ')" | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "------------------------------------------------------------------------------------"| Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Directory Sync"  | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "On-Premises Sync Enabled: $($organization.OnPremisesSyncEnabled)"  | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
if ($organization.OnPremisesSyncEnabled) {
    echo "Last Dir Sync Time: $($organization.OnPremisesLastSyncDateTime)`n"  | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
}
echo "------------------------------------------------------------------------------------"| Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Licensing Information" | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "(Note: License details available via Graph API with additional permissions)" | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "------------------------------------------------------------------------------------" | Tee -Append -FilePath .\${CURRENTJOB}.Report.txt

#get domain info
echo "Retrieving Domain Information:"
write-host -NoNewline "`t`t`tRetrieving Domain Information ... "
$organization.VerifiedDomains | Format-Table -Auto | Out-File -FilePath .\${CURRENTJOB}.O365.DomainInfo.txt
echo "`t`tDONE"

echo "------------------------------------------------------------------------------------"

######################################################################################
# USER INFO

echo "Retrieving User Information (this may take a while):" 

Write-Host -NoNewline "`t`t`tRetrieving and processing User List ..."
# Stream user data directly without storing in memory using Microsoft Graph
Get-MgUser -All -PageSize 999 | ForEach-Object {
    # Trim and output each UserPrincipalName to the simple list
    $_.UserPrincipalName.Trim(" ") | Out-File -Append -FilePath .\${CURRENTJOB}.O365.Users.txt 

    # Filter and export detailed list if not a HealthMailbox
    if ($_.UserPrincipalName -notlike "HealthMailbox*") {
        $_ | Select-Object -Property UserPrincipalName, DisplayName, Department, JobTitle, @{Name='PhoneNumber';Expression={$_.BusinessPhones -join ', '}}, @{Name='Office';Expression={$_.OfficeLocation}}, @{Name='PasswordNeverExpires';Expression={$_.PasswordPolicies -like "*DisablePasswordExpiration*"}}, @{Name='LastPasswordChangeTimestamp';Expression={$_.PasswordLastChangedDateTime}}, @{Name='LastDirSyncTime';Expression={$_.OnPremisesLastSyncDateTime}} | Export-Csv -Append -Path .\${CURRENTJOB}.O365.Users_Detailed.csv
    }

    # Output user SignInName and ProxyAddresses
    "$($_.UserPrincipalName),$($_.Mail -join ', ')" | Out-File -Append -FilePath .\${CURRENTJOB}.O365.Users_ProxyAddresses.txt

    # Output all properties in LDAP style format
    $_ | Select-Object -Property * | Out-File -Append -FilePath .\${CURRENTJOB}.O365.Users_LDAP_details.txt
}
echo "`t`t`tDONE"

echo "------------------------------------------------------------------------------------"

######################################################################################
# GROUP INFO

echo "Retrieving Group Information:"
Write-Host -NoNewline "`t`t`tRetrieving Group Names ... "
$grouplist = Get-MgGroup -All -PageSize 999
echo "`t`tDONE"


Write-Host -NoNewline "`t`t`tCreating Simple Group List ... "
foreach($line in $grouplist){$line.DisplayName.Trim(" ") | Out-File -Append -FilePath .\${CURRENTJOB}.O365.Groups.txt } 
echo "`t`tDONE"


Write-Host -NoNewline "`t`t`tRetrieving Extended Group Information ... "
$grouplist | Format-Table -Property DisplayName,Description,GroupTypes -AutoSize | out-string -width 1024 | Out-File -Append -FilePath .\${CURRENTJOB}.O365.Groups_Advanced.txt
echo "`tDONE"
echo "------------------------------------------------------------------------------------"



######################################################################################
# GROUP MEMBERSHIP

echo "Retrieving Group Membership (this may take a while): "

Write-Host -NoNewline "`t`t`tIterating Group Membership ... "
# enum4linux style group membership using Microsoft Graph
$grouplist | ForEach-Object {
    $CURRENTGROUP=$_.DisplayName
    $memberlist=$(Get-MgGroupMember -GroupId $_.Id -All); 
    if ($memberlist -ne $null){ 
        foreach ($item in $memberlist){
            # Get user details from the member object
            $memberEmail = $item.AdditionalProperties['userPrincipalName']
            if (-not $memberEmail) {
                $memberEmail = $item.Id
            }
            echo "$($CURRENTGROUP):$($memberEmail)" | Out-File -Append -FilePath .\${CURRENTJOB}.O365.GroupMembership.txt
        } 
    }
}
echo "`t`tDONE"
echo "------------------------------------------------------------------------------------"


######################################################################################
# ROLE MEMBERSHIP

echo "Retrieving Role Membership (this may take a longer while):"
Write-Host -NoNewline "`t`t`tIterating Admin Role Membership ... "
# Using Microsoft Graph to get directory roles and their members
Get-MgDirectoryRole | Where-Object { $_.DisplayName -like "*admin*" } | ForEach-Object {
    $testrole = $_.DisplayName
    Get-MgDirectoryRoleMember -DirectoryRoleId $_.Id | ForEach-Object {
        # Get user details for each member
        $user = Get-MgUser -UserId $_.Id -Property UserPrincipalName, DisplayName
        [PSCustomObject]@{
            RoleName = $testrole
            EmailAddress = $user.UserPrincipalName
        }
    }
} | Sort-Object RoleName, EmailAddress | Out-File -Append -FilePath .\${CURRENTJOB}.O365.Roles_Admins.txt
echo "`t`tDONE"
echo "------------------------------------------------------------------------------------"


######################################################################################
# DEVICE INFO

echo "Retrieving Device Information:"
Write-Host -NoNewline "`t`t`tRetrieving Device Information ... "
$o365devicelist = Get-MgDevice -All
echo "`t`tDONE"

Write-Host -NoNewline "`t`t`tCreating Simple Device list ... "
# if we just are lazy and use ft, then our output file will have whitespace at the end :-/
foreach($line in $o365devicelist){$line.DisplayName.Trim(" ") | Out-File -Append -FilePath .\${CURRENTJOB}.O365.DeviceList.txt } 
echo "`t`tDONE"

Write-Host -NoNewline "`t`t`tCreating Extended Device List ... "
$o365devicelist | Select-Object -Property DisplayName, @{Name='DeviceOsType';Expression={$_.OperatingSystem}}, @{Name='DeviceTrustType';Expression={$_.TrustType}}, @{Name='ApproximateLastLogonTimestamp';Expression={$_.ApproximateLastSignInDateTime}}, Enabled | Export-Csv -Path .\${CURRENTJOB}.O365.DeviceList_Advanced.csv
echo "`t`tDONE"

# Get device owner mapping
Write-Host -NoNewline "`t`t`tCreating user->device mapping ..."
#This pulls down a list of devices and looks up corresponding owner
$o365devicelist | ForEach-Object { 
    $OwnerObject = Get-MgDeviceRegisteredOwner -DeviceId $_.Id -ErrorAction SilentlyContinue
    $ownerName = ""
    if ($OwnerObject) {
        $ownerName = $OwnerObject.AdditionalProperties['displayName']
    }
    echo "$ownerName,$($_.DisplayName),$($_.OperatingSystem)"
} | Sort-Object | Out-File -FilePath .\${CURRENTJOB}.O365.DeviceList_Owners.csv
echo "`t`tDONE"
echo "------------------------------------------------------------------------------------"


######################################################################################
# User/Group/Device Statistics - for Report Only

echo "Overview of Environment" | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
# Count users from the file we created
$userCount = (Get-Content .\${CURRENTJOB}.O365.Users.txt | Measure-Object -Line).Lines
echo "Number of users: $userCount"  | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Number of groups: $($grouplist.Count)"| Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "Number of devices: $($o365devicelist.Count)"| Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "------------------------------------------------------------------------------------"| Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt

######################################################################################
# APPLICATION SECURITY REPORT

# UserPermission settings

echo "Checking Applications in Azure AD:"
Write-Host -NoNewline "`t`t`tRetrieving a list of Applications ... "   
$azureadapps = Get-MgApplication -All
$azureadapps | ForEach-Object { 
    [PSCustomObject]@{
        AppId = $_.AppId
        DisplayName = $_.DisplayName
        SignInAudience = $_.SignInAudience
    }
} | Out-File -Append -FilePath .\${CURRENTJOB}.O365.ApplicationList.txt
echo "Application Information" | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
$azureadapps | ForEach-Object { 
    [PSCustomObject]@{
        AppId = $_.AppId
        DisplayName = $_.DisplayName
    }
} | Out-File -Append -FilePath .\${CURRENTJOB}.Report.txt
echo "`tDONE"

Write-Host -NoNewline "`t`t`tCreating user->application mapping ..."
#This pulls down a list of applications and looks up corresponding owner
$azureadapps | ForEach-Object { 
    $OwnerObject = Get-MgApplicationOwner -ApplicationId $_.Id -ErrorAction SilentlyContinue
    $ownerUpn = ""
    $ownerName = ""
    if ($OwnerObject) {
        $ownerUpn = $OwnerObject.AdditionalProperties['userPrincipalName']
        $ownerName = $OwnerObject.AdditionalProperties['displayName']
    }
    echo "$ownerUpn,$ownerName,$($_.DisplayName)"
} | Sort-Object | Out-File -FilePath .\${CURRENTJOB}.O365.Application_Owners.csv
echo "`t`tDONE"

echo "" | Tee -Append -FilePath .\${CURRENTJOB}.Report.txt

# Check organization policies using Microsoft Graph
$policies = Get-MgPolicyAuthorizationPolicy -ErrorAction SilentlyContinue
if ($policies) {
    # Check default user role permissions
    if ($policies.DefaultUserRolePermissions) {
        if ($policies.DefaultUserRolePermissions.AllowedToCreateApps -eq $true) {
            echo "[!] Users in this tenant are allowed to create applications" | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
        }
        if ($policies.DefaultUserRolePermissions.AllowedToReadOtherUsers -eq $true) {
            echo "[!] Users in this tenant are allowed to read other user information." | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
        }
    }
}

# Check consent policy
$consentPolicy = Get-MgPolicyAuthorizationPolicyConsentPolicySetting -ErrorAction SilentlyContinue
if ($consentPolicy) {
    echo "[!] Consent policy configured. Review settings manually." | tee -Append -FilePath .\${CURRENTJOB}.Report.txt
}

echo "------------------------------------------------------------------------------------"


# look for admin groups
echo "Group Membership Checks:"
Write-Host -NoNewline "`t`t`tLooking for Admin users ... "
# enum4linux style group membership using Microsoft Graph
$grouplist | Where-Object { ( $_.DisplayName -like "*admin*" ) } | ForEach-Object {
    $CURRENTGROUP=$_.DisplayName
    $memberlist=$(Get-MgGroupMember -GroupId $_.Id -All); 
    if ($memberlist -ne $null){ 
        foreach ($item in $memberlist){
            $memberEmail = $item.AdditionalProperties['userPrincipalName']
            if (-not $memberEmail) {
                $memberEmail = $item.Id
            }
            echo "$($CURRENTGROUP):$($memberEmail)" | Out-File -Append -FilePath .\${CURRENTJOB}.O365.GroupMembership_AdminGroups.txt
        } 
    }
}
echo "`t`t`tDONE"

# vpn groups - look for alternate names like globalprotect etc
Write-Host -NoNewline "`t`t`tLooking for VPN groups ... "
# enum4linux style group membership using Microsoft Graph
$grouplist | Where-Object { ( $_.DisplayName -like "*vpn*" ) -Or ( $_.DisplayName -like "*cisco*" ) -Or ( $_.DisplayName -like "*globalprotect*" ) -Or ( $_.DisplayName -like "*palo*" ) } | ForEach-Object {
    $CURRENTGROUP=$_.DisplayName
    $memberlist=$(Get-MgGroupMember -GroupId $_.Id -All); 
    if ($memberlist -ne $null){ 
        foreach ($item in $memberlist){
            $memberEmail = $item.AdditionalProperties['userPrincipalName']
            if (-not $memberEmail) {
                $memberEmail = $item.Id
            }
            echo "$($CURRENTGROUP):$($memberEmail)" | Out-File -Append -FilePath .\${CURRENTJOB}.O365.GroupMembership_VPNGroups.txt
        } 
    }
}
echo "`t`t`tDONE"
echo "------------------------------------------------------------------------------------"
echo ""
echo "JOB COMPLETE: GO GET YOUR LOOT!"
ls .\${CURRENTJOB}*

# haec programma meum est. multa similia sunt, sed haec una mea est.
