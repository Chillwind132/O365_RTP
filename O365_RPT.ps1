#Config Variables
$TenantSiteURL =  "https://pwceur-admin.sharepoint.com"
$CSVFilePath = ".\data.csv"

Connect-SPOService -Url $TenantSiteURL

$SiteCollections = Get-SPOSite -Limit "3"

$UserData = @()
$updated_items_list = @()
forEach($Site in $SiteCollections)
{
    Write-Host "Checking:"$Site.URL
    $Users = Get-SPOUser -Site $Site 
    
    forEach($item in $Users){
        $updated_item_d = $item | Select -ExpandProperty "DisplayName"
        $updated_item_l = $item | Select -ExpandProperty "LoginName"
        $updated_item = $updated_item_d + ":" + $updated_item_l
        $updated_items_list += $updated_item

    }   
      
    $UserData += New-Object PSObject -Property  @{
        'Site URL' = $Site.URL
        'Users_DisplayName' = ($Users.DisplayName | Out-String).Trim()
        'Users_LoginName' =  ($Users.LoginName | Out-String).Trim()
        'Users' = ($updated_items_list | Out-String).Trim()
    }
    
    Write-Host $updated_item
    Write-Host $updated_items_list

    Write-Host $Users.count
    Write-Host "Done"
}

$UserData | Select-Object "Site URL", "Users_DisplayName", "Users_LoginName", "Users" | Export-CSV $CSVFilePath -NoTypeInformation

Write-Host "Done"