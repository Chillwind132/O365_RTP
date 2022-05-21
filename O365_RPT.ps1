$TenantSiteURL =  "https://pwceur-admin.sharepoint.com"
$CSVFilePath = ".\data.csv"

function user_input {
    $user_selection_report = Read-Host "Would you like to generate user & groups report or user permissions report? Select 1 for user/groups report; Select 2 for permissions report"
    while (($user_selection_report -ne '1') -and ($user_selection_report -ne '2')) {
        Write-Host "Invalid input"
        $user_selection_report = Read-Host "Would you like to generate user & groups report or user permissions report? Select 1 for user/groups report; Select 2 for permissions report"
    }
    $user_selection_level = Read-Host "Select 1 for tenant level report; Select 2 for individual site col report"
    while (($user_selection_level -ne '1') -and ($user_selection_level -ne '2')) {
        Write-Host "Invalid input"
        $user_selection_level = Read-Host "Select 1 for tenant level report; Select 2 for individual site col report"
    }
    $user_selection_array = @()
    $user_selection_array += $user_selection_report
    $user_selection_array += $user_selection_level

    return $user_selection_array
}

function generate_users_report_tenant {
    Connect-SPOService -Url $TenantSiteURL 

    $SiteCollections = Get-SPOSite -Limit "3"

    $UserData = @()
    $users_list = @()
    $groups_list = @()
    $ExternalUsers = @()

    Write-Host "Number of sites in tenant:"$SiteCollections.count -ForegroundColor Blue

    forEach ($Site in $SiteCollections) {
        Write-Host "Checking:"$Site.URL -ForegroundColor Yellow
        $Users = Get-SPOUser -Site $Site 
    
        forEach ($item in $Users) {
            $users_d = $item | Select -ExpandProperty "DisplayName"
            $users_l = $item | Select -ExpandProperty "LoginName"
            $users_list += $users_d + ":" + $users_l
        }   
    
        $Groups = Get-SPOSiteGroup -Site $Site.URL 
    
        forEach ($item in $Groups) {
            $groups_title = $item.Title
            $groups_users = $item | Select -ExpandProperty "Users"
            $groups_roles = $item | Select -ExpandProperty "Roles"

            $groups_list += $groups_title + ":" + $groups_users + ":" + $groups_roles
        }

        $ExtUsers = Get-SPOUser -Limit All -Site $Site.URL | Where { $_.LoginName -like "*#ext#*" -or $_.LoginName -like "*urn:spo:guest*" }
        If ($ExtUsers.count -gt 0) {
            $ExternalUsers += $ExtUsers
        }
        else {
            $ExternalUsers += 'No external users present'
        }

        $UserData += New-Object PSObject -Property  @{
            'Site URL'          = $Site.URL
            'Users_DisplayName' = ($Users.DisplayName | Out-String).Trim()
            'Users_LoginName'   = ($Users.LoginName | Out-String).Trim()
            'Users'             = ($users_list | Out-String).Trim()
            'User_Groups'       = ($groups_list | Out-String).Trim()
            'External_users'    = ($ExternalUsers | Out-String).Trim()
        }

    }

    $UserData | Select-Object "Site URL", "Users_DisplayName", "Users_LoginName", "Users", "User_Groups", 'External_users' | Export-CSV $CSVFilePath -NoTypeInformation

    Write-Host "Done"
}


$user_selection_array = user_input

if (($user_selection_array[0] -eq 1) -and ($user_selection_array[1] -eq 1)) {
    Write-Host "You selected to generate user & groups report at tenant level" -ForegroundColor Yellow
    generate_users_report_tenant

}






