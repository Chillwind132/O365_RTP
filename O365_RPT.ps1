#Config Variables
$TenantSiteURL =  "https://pwceur-admin.sharepoint.com"
$CSVFilePath = ".\data.csv"

Connect-SPOService -Url $TenantSiteURL

$SiteCollections = Get-SPOSite -Limit "3"

$UserData = @()

forEach($Site in $SiteCollections)
{
    Write-Host "Checking:"$Site.URL
    $Users = Get-SPOUser -Site $Site 
    $UserData += New-Object PSObject -Property  @{
        'Site URL' = $Site.URL
        'Users_DisplayName' = ($Users.DisplayName | Out-String).Trim()
        'Users_LoginName' =  ($Users.LoginName | Out-String).Trim()
    }
    Write-Host "Done"
}

$UserData | Select-Object "Site URL", "Users_DisplayName", "Users_LoginName" | Export-CSV $CSVFilePath -NoTypeInformation

Write-Host "Done"