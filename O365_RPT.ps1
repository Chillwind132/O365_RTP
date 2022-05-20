$TenantSiteURL =  "https://pwceur-admin.sharepoint.com"
$CSVFilePath = ".\data.csv"

Connect-SPOService -Url $TenantSiteURL

$SiteCollections = Get-SPOSite -Limit "3"

$UserData = @()
$updated_items_list = @()
$groups_list = @()

forEach($Site in $SiteCollections)
{
    Write-Host "Checking:"$Site.URL
    $Users = Get-SPOUser -Site $Site 
    
    forEach($item in $Users){
        $users_d = $item | Select -ExpandProperty "DisplayName"
        $users_l = $item | Select -ExpandProperty "LoginName"
        $users_list += $users_d + ":" + $users_l
    }   
    
    $Groups = Get-SPOSiteGroup -Site $Site.URL 
    
    forEach($item in $Groups){
        $groups_title = $item.Title
        $groups_users = $item | Select -ExpandProperty "Users"
        $groups_roles = $item | Select -ExpandProperty "Roles"

        $groups_list += $groups_title + ":" + $groups_users + ":" + $groups_roles
    }

    $UserData += New-Object PSObject -Property  @{
        'Site URL' = $Site.URL
        'Users_DisplayName' = ($Users.DisplayName | Out-String).Trim()
        'Users_LoginName' =  ($Users.LoginName | Out-String).Trim()
        'Users' = ($updated_items_list | Out-String).Trim()
        'User_Groups' = ($groups_list | Out-String).Trim()
    }

}

$UserData | Select-Object "Site URL", "Users_DisplayName", "Users_LoginName", "Users", "User_Groups" | Export-CSV $CSVFilePath -NoTypeInformation

Write-Host "Done"