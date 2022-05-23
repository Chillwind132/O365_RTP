$CSVFilePath =  ".\data.csv"

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

function generate_users_report_sitecol ($SiteURL) {

    $UserData = @()
    $users_list = @()
    $groups_list = @()
    $ExternalUsers = @()
    
    Connect-SPOService -Url $TenantSiteURL
    $Site = Get-SPOSite -Identity $SiteURL
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

    $UserData | Select-Object "Site URL", "Users_DisplayName", "Users_LoginName", "Users", "User_Groups", 'External_users' | Export-CSV $CSVFilePath -NoTypeInformation

    Write-Host "Done"
}
function generate_users_report_tenant($TenantSiteURL) {
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

Function Get-PnPPermissions([Microsoft.SharePoint.Client.SecurableObject]$Object) {
    #Reference -> https://www.sharepointdiary.com/2018/09/sharepoint-online-site-collection-permission-report-using-powershell.html
    Switch ($Object.TypedObject.ToString()) {
        "Microsoft.SharePoint.Client.Web" { $ObjectType = "Site" ; $ObjectURL = $Object.URL; $ObjectTitle = $Object.Title }
        "Microsoft.SharePoint.Client.ListItem" {
            If ($Object.FileSystemObjectType -eq "Folder") {
                $ObjectType = "Folder"
                
                $Folder = Get-PnPProperty -ClientObject $Object -Property Folder
                $ObjectTitle = $Object.Folder.Name
                $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $Object.Folder.ServerRelativeUrl)
            }
            Else { 
                Get-PnPProperty -ClientObject $Object -Property File, ParentList
                If ($Object.File.Name -ne $Null) {
                    $ObjectType = "File"
                    $ObjectTitle = $Object.File.Name
                    $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $Object.File.ServerRelativeUrl)
                }
                else {
                    $ObjectType = "List Item"
                    $ObjectTitle = $Object["Title"]
                    $DefaultDisplayFormUrl = Get-PnPProperty -ClientObject $Object.ParentList -Property DefaultDisplayFormUrl                    
                    $ObjectURL = $("{0}{1}?ID={2}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $DefaultDisplayFormUrl, $Object.ID)
                }
            }
        }
        Default {
            $ObjectType = "List or Library"
            $ObjectTitle = $Object.Title
            $RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder    
            $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $RootFolder.ServerRelativeUrl)
        }
    }
   
    Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments
 
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments
     
    $PermissionCollection = @()
    Foreach ($RoleAssignment in $Object.RoleAssignments) {
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
 
        $PermissionType = $RoleAssignment.Member.PrincipalType
    
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name
 
        $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access" }) -join ","
 
        If ($PermissionLevels.Length -eq 0) { Continue }
 
        If ($PermissionType -eq "SharePointGroup") {
            $GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName
                 
            If ($GroupMembers.count -eq 0) { Continue }
            $GroupUsers = ($GroupMembers | Select -ExpandProperty Title) -join ","
 
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($GroupUsers)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
            $PermissionCollection += $Permissions
        }
        Else {
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($RoleAssignment.Member.Title)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
            $PermissionCollection += $Permissions
        }
    }
    $PermissionCollection | Export-CSV $CSVFilePath -NoTypeInformation -Append
}
   
Function Generate-PnPSitePermissionRpt() {
    [cmdletbinding()]
 
    Param 
    (   
        [Parameter(Mandatory = $false)] [String] $SiteURL,
        [Parameter(Mandatory = $false)] [String] $ReportFile           
    ) 
    
    Try {
        Connect-PnPOnline -URL $SiteURL -Interactive
        $Web = Get-PnPWeb
        $Recursive = $true
        $ScanItemLevel = $true
        $IncludeInheritedPermissions = $true
 
        Write-host -f Yellow "Getting Site Collection Administrators..."
        $SiteAdmins = Get-PnPSiteCollectionAdmin
         
        $SiteCollectionAdmins = ($SiteAdmins | Select -ExpandProperty Title) -join ","
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Object("Site Collection")
        $Permissions | Add-Member NoteProperty Title($Web.Title)
        $Permissions | Add-Member NoteProperty URL($Web.URL)
        $Permissions | Add-Member NoteProperty HasUniquePermissions("TRUE")
        $Permissions | Add-Member NoteProperty Users($SiteCollectionAdmins)
        $Permissions | Add-Member NoteProperty Type("Site Collection Administrators")
        $Permissions | Add-Member NoteProperty Permissions("Site Owner")
        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
               
        $Permissions | Export-CSV $CSVFilePath -NoTypeInformation
   
        Function Get-PnPListItemsPermission([Microsoft.SharePoint.Client.List]$List) {
            Write-host -f Yellow "`t `t Getting Permissions of List Items in the List:"$List.Title
  
            $ListItems = Get-PnPListItem -List $List -PageSize 500
  
            $ItemCounter = 0
            ForEach ($ListItem in $ListItems) {
                If ($IncludeInheritedPermissions) {
                    Get-PnPPermissions -Object $ListItem
                }
                Else {
                    $HasUniquePermissions = Get-PnPProperty -ClientObject $ListItem -Property HasUniqueRoleAssignments
                    If ($HasUniquePermissions -eq $True) {
                        #Call the function to generate Permission report
                        Get-PnPPermissions -Object $ListItem
                    }
                }
                $ItemCounter++
                Write-Progress -PercentComplete ($ItemCounter / ($List.ItemCount) * 100) -Activity "Processing Items $ItemCounter of $($List.ItemCount)" -Status "Searching Unique Permissions in List Items of '$($List.Title)'"
            }
        }
 
        Function Get-PnPListPermission([Microsoft.SharePoint.Client.Web]$Web) {
            $Lists = Get-PnPProperty -ClientObject $Web -Property Lists
            $ExcludedLists = @("Access Requests", "App Packages", "appdata", "appfiles", "Apps in Testing", "Cache Profiles", "Composed Looks", "Content and Structure Reports", "Content type publishing error log", "Converted Forms",
                "Device Channels", "Form Templates", "fpdatasources", "Get started with Apps for Office and SharePoint", "List Template Gallery", "Long Running Operation Status", "Maintenance Log Library", "Images", "site collection images"
                , "Master Docs", "Master Page Gallery", "MicroFeed", "NintexFormXml", "Quick Deploy Items", "Relationships List", "Reusable Content", "Reporting Metadata", "Reporting Templates", "Search Config List", "Site Assets", "Preservation Hold Library",
                "Site Pages", "Solution Gallery", "Style Library", "Suggested Content Browser Locations", "Theme Gallery", "TaxonomyHiddenList", "User Information List", "Web Part Gallery", "wfpub", "wfsvc", "Workflow History", "Workflow Tasks", "Pages")
             
            $Counter = 0
            ForEach ($List in $Lists) {
                If ($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title) {
                    $Counter++
                    Write-Progress -PercentComplete ($Counter / ($Lists.Count) * 100) -Activity "Exporting Permissions from List '$($List.Title)' in $($Web.URL)" -Status "Processing Lists $Counter of $($Lists.Count)"
 
                    If ($ScanItemLevel) {
                        Get-PnPListItemsPermission -List $List
                    }
 
                    If ($IncludeInheritedPermissions) {
                        Get-PnPPermissions -Object $List
                    }
                    Else {
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments
                        If ($HasUniquePermissions -eq $True) {
                            Get-PnPPermissions -Object $List
                        }
                    }
                }
            }
        }
   
        Function Get-PnPWebPermission([Microsoft.SharePoint.Client.Web]$Web) {
            Write-host -f Yellow "Getting Permissions of the Web: $($Web.URL)..." 
            Get-PnPPermissions -Object $Web
   
            Write-host -f Yellow "`t Getting Permissions of Lists and Libraries..."
            Get-PnPListPermission($Web)
 
            If ($Recursive) {
                $Subwebs = Get-PnPProperty -ClientObject $Web -Property Webs
 
                Foreach ($Subweb in $web.Webs) {
                    If ($IncludeInheritedPermissions) {
                        Get-PnPWebPermission($Subweb)
                    }
                    Else {
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $SubWeb -Property HasUniqueRoleAssignments
   
                        If ($HasUniquePermissions -eq $true) {
                            Get-PnPWebPermission($Subweb)
                        }
                    }
                }
            }
        }
        Get-PnPWebPermission $Web
        Write-host -f Green "`n*** Site Permission Report Generated Successfully!***"
    }
    Catch {
        write-host -f Red "Error Generating Site Permission Report!" $_.Exception.Message
    }
}

function generate_permissions_report_sitecol($SiteURL) {
    Generate-PnPSitePermissionRpt -SiteURL $SiteURL -ReportFile $CSVFilePath
}

$user_selection_array = user_input

if (($user_selection_array[0] -eq 1) -and ($user_selection_array[1] -eq 1)) { #user report for tenant
    Write-Host "You selected to generate user & groups report at tenant level" -ForegroundColor Yellow
    $user_input = Read-Host "Input the target SharePoint Online admin center URL (For example: https://pwceur-admin.sharepoint.com)"
    generate_users_report_tenant($user_input)

} ElseIf (($user_selection_array[0] -eq 2) -and ($user_selection_array[1] -eq 2)) { # permissions report for a site col
    Write-Host "You selected to generate permissions report at site col level" -ForegroundColor Yellow
    $user_input = Read-Host "Input the target SharePoint Online Site URL"
    generate_permissions_report_sitecol($user_input)
} ElseIf (($user_selection_array[0] -eq 1) -and ($user_selection_array[1] -eq 2)) { #user report for a site col
    Write-Host "You selected to generate user & groups report for a site col" -ForegroundColor Yellow
    $user_input = Read-Host "Input the target SharePoint Online Site URL"
    generate_users_report_sitecol($user_input)
} 






