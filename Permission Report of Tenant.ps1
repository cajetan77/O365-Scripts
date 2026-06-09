

$ClientId = "dc223b11-5ab5-4a33-988a-3474b25eb9be"
$Thumbprint = "2D2C9AD033452336BD161D4E8CA88E164398FC43"
$Tenant = "caje77sharepoint.onmicrosoft.com"


function Display-ProcessingTime {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][DateTime]$StartDate,
        [Parameter(Mandatory = $false)][DateTime]$EndDate,
        [Parameter(Mandatory = $false)][String]$AdditionalText,
        [Parameter(Mandatory = $false)][String]$BatchInformation,
        [Parameter(Mandatory = $false)][string]$Indentation
    );
    if ($EndDate -eq $null) { $EndDate = (Get-Date); }
    $TotalTime = ($EndDate - $StartDate);

    if ($AdditionalText -eq $null) { $AdditionalText = ""; }
    if ($BatchInformation -eq $null) { $BatchInformation = ""; }
    

    $Hours = $TotalTime.Hours.ToString(); if ($TotalTime.Hours -lt 10) { $Hours = ("0" + $Hours); }
    $Minutes = $TotalTime.Minutes.ToString(); if ($TotalTime.Minutes -lt 10) { $Minutes = ("0" + $Minutes); }
    $Seconds = $TotalTime.Seconds.ToString(); if ($TotalTime.Seconds -lt 10) { $Seconds = ("0" + $Seconds); }
    if ($Indentation -eq $null) { $Indentation = ""; }

    $DisplayStr = (
        $Indentation +
        $BatchInformation +
        $AdditionalText +         
        $TotalTime.Days.ToString() +
        "." + $Hours + 
        ":" + $Minutes +
        ":" + $Seconds + 
        "." + $TotalTime.Milliseconds.ToString()
    );

    Write-Host ($DisplayStr);
}

function Get-ContentTypeTree {
    [cmdletbinding()]
    param (                                
        [Parameter(Mandatory = $true)]$ContentType
    );

    $ContentTypeTree = (New-Object PSCustomObject -Property @{
            IdPath   = "";
            NamePath = "";
        })
    
    $ParentContentType = (Get-PnPProperty -ClientObject $ContentType -Property "Parent")
    if ($ParentContentType.Name -ne "System") {
        $ParentContentTypeTree = (Get-ContentTypeTree -ContentType $ParentContentType);        
        $ContentTypeTree.NamePath = $ParentContentTypeTree.NamePath + " > ";
        $ContentTypeTree.IDPath = $ParentContentTypeTree.IDPath + " > ";
    }

    $ContentTypeTree.NamePath += ($ContentType.Name);
    $ContentTypeTree.IDPath += ($ContentType.Id.StringValue);    
    return $ContentTypeTree;
}

function Call-SharePointHelper-RestApiMethod {
    [cmdletbinding()]
    param (                                
        [Parameter(Mandatory = $true)]$RestApiUrl
    );

    $theItems = @();    
    $NumRestCalls = 0;
    do { 
        $RestApiResult = (Invoke-PnPSPRestMethod -Url $RestApiUrl -ErrorAction Stop); 
        $theItems += ($RestApiResult.value); 
        $RestApiUrl = $RestApiResult.'odata.nextLink'; 

        $NumRestCalls++;
        if (($NumRestCalls % 10) -eq 0) {
            Start-Sleep -Seconds 120;    
        }
    } 
    while (-not [String]::IsNullOrEmpty($RestApiUrl));

    return $theItems;
}

function Get-SharePointHelper-RoleDefinitions {
    $theResult = @{};
    $RoleTypeDefs = (Get-PnPRoleDefinition -ErrorAction Stop);
    foreach ($currRoleTypeDef in $RoleTypeDefs) {
        $theResult[($currRoleTypeDef.Id.ToString())] = $currRoleTypeDef.Name;
    }

    return $theResult;
}

function Get-SharePointHelper-GetRoleAssignmentsForObject {
    [cmdletbinding()]
    param (                                
        [Parameter(Mandatory = $true)]$SiteUrl
        , [Parameter(Mandatory = $true)]$ObjectType
        , [Parameter(Mandatory = $false)]$LibraryGuid
        , [Parameter(Mandatory = $false)]$ItemID
        , [Parameter(Mandatory = $false)]$RoleDefinitions
    );


    if ($RoleDefinitions -eq $null) {
        $RoleDefinitions = (Get-SharePointHelper-RoleDefinitions -ErrorAction Stop);
    }


    $RestApiUrl = ($SiteUrl + "/_api/");
    if ($ObjectType -eq "Web") {
        $RestApiUrl += "web";
    }
    elseif (($ObjectType -eq "List") -or ($ObjectType -eq "Library")) {
        $RestApiUrl += ("web/lists(guid'" + $LibraryGuid + "')");
    }
    elseif (($ObjectType -eq "Item")) {
        $RestApiUrl += ("web/lists(guid'" + $LibraryGuid + "')/items(" + $ItemID + ")");
    }

    $RestApiUrl += ("/roleassignments?`$expand=Member/users,RoleDefinitionBindings&`$top=1000");    

    $ObjectPermissions = (Call-SharePointHelper-RestApiMethod $RestApiUrl); 

    
    $theResult = @{};

    foreach ($currPerm in $ObjectPermissions) {
        $MoreThanLimitedAccess = $false;
        $Permissions = @();

        foreach ($currRoleDef in $currPerm.RoleDefinitionBindings) {
            $currRoleDefPerm = ($RoleDefinitions[($currRoleDef.Id.ToString())]);
            if (($currRoleDefPerm -ne "Limited Access") -and ($currRoleDefPerm -ne "Web-Only Limited Access")) {
                $MoreThanLimitedAccess = $true;                        

                #Exclude Limited Access
                $Permissions += $currRoleDefPerm;
            }            
        }

        if ($MoreThanLimitedAccess) {
            
            #https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee541430(v=office.15)
            $AccessType = "";
            switch ([int]$currPerm.Member.PrincipalType) {
                0 { $AccessType = "None"; break; }
                1 { $AccessType = "User"; break; }
                2 { $AccessType = "DistributionList"; break; }
                4 { $AccessType = "SecurityGroup"; break; }
                8 { $AccessType = "SharePointGroup"; break; }
                15 { $AccessType = "All"; break; }
            }

            $theResult[$currPerm.Member.LoginName] = (New-Object PSCustomObject -Property @{
                    MemberLoginName = ($currPerm.Member.LoginName);
                    Title           = ($currPerm.Member.Title);
                    AccessType      = $AccessType;
                    Permissions     = $Permissions;
                });
        }
    }

    return $theResult;
}

function Format-StringForExport {
    [cmdletbinding()]
    param (                                
        [Parameter(Mandatory = $false)][String]$str
    );

    if ([String]::IsNullOrEmpty($str)) {
        return "";
    }
    else {
        return $str;
    }
}

function Get-PermissionAuditEntry {
    
    [cmdletbinding()]
    param (                                
        [Parameter(Mandatory = $true)][String]$Scope
        , [Parameter(Mandatory = $true)][String]$SiteUrl
        , [Parameter(Mandatory = $true)][String]$SiteName
        , [Parameter(Mandatory = $false)][String]$ListRootFolderName
        , [Parameter(Mandatory = $false)][String]$ListDisplayName
        , [Parameter(Mandatory = $false)][String]$ContentTypeName
        , [Parameter(Mandatory = $false)][String]$IsFolder
        , [Parameter(Mandatory = $false)][String]$IsSharingLink
        , [Parameter(Mandatory = $false)][String]$ItemName
        , [Parameter(Mandatory = $false)][String]$ItemUrl
        , [Parameter(Mandatory = $false)][String]$ItemID
        , [Parameter(Mandatory = $true)][String]$AccessType
        , [Parameter(Mandatory = $true)][String]$Name
        , [Parameter(Mandatory = $true)][String]$Permissions
        , [Parameter(Mandatory = $false)][String]$Members
        , [Parameter(Mandatory = $false)][String]$Notes
        , [Parameter(Mandatory = $false)][String]$Delimiter
    );

    return (
        $Scope + 
        $Delimiter + $SiteUrl +
        $Delimiter + $SiteName +
        $Delimiter + (Format-StringForExport -str $ListRootFolderName) +
        $Delimiter + (Format-StringForExport -str $ListDisplayName) +
        $Delimiter + (Format-StringForExport -str $ContentTypeName) +
        $Delimiter + (Format-StringForExport -str $IsFolder) +
        $Delimiter + (Format-StringForExport -str $IsSharingLink) +
        $Delimiter + (Format-StringForExport -str $ItemName) +
        $Delimiter + (Format-StringForExport -str $ItemUrl) +
        $Delimiter + (Format-StringForExport -str $ItemID) +
        $Delimiter + $AccessType +
        $Delimiter + $Name +
        $Delimiter + $Permissions +
        $Delimiter + $Members +
        $Delimiter + $Notes
    );
}

function Process-PermissionsForSite {
    [cmdletbinding()]
    param (                                
        [Parameter(Mandatory = $true)][String]$SiteUrl
        , [Parameter(Mandatory = $true)][String]$ExportDirectoryAndName
        , [Parameter(Mandatory = $true)][String]$Delimiter        
        , [switch]$ProcessLists
        , [switch]$ProcessItems
    );

    Write-Host ("Processing: " + $SiteUrl);

    $StartTime = (Get-Date);
    $ExcludedLists = @{};
    $ExcludedLists["Access Requests"] = 1;
    $ExcludedLists["_catalogs/appdata"] = 1;
    $ExcludedLists["_catalogs/appfiles"] = 1;
    $ExcludedLists["_catalogs/design"] = 1;
    $ExcludedLists["_catalogs/masterpage"] = 1;
    $ExcludedLists["_catalogs/wp"] = 1;
    $ExcludedLists["_catalogs/theme"] = 1;
    $ExcludedLists["_catalogs/users"] = 1;
    $ExcludedLists["_catalogs/solutions"] = 1;
    $ExcludedLists["_catalogs/hubsite"] = 1;
    $ExcludedLists["IWConvertedForms"] = 1;   
    $ExcludedLists["_catalogs/MaintenanceLogs"] = 1;
    $ExcludedLists["FormServerTemplates"] = 1;
    $ExcludedLists["PreservationHoldLibrary"] = 1;
    $ExcludedLists["Teams Wiki Data"] = 1;
    $ExcludedLists["Lists/DO_NOT_DELETE_SPLIST_SITECOLLECTION_AGGREGATED_CON"] = 1;
    $ExcludedLists["_catalogs/lt"] = 1;
    $ExcludedLists["Lists/SharePointHomeCacheList"] = 1;
    $ExcludedLists["Lists/TaxonomyHiddenList"] = 1;
    $ExcludedLists["_catalogs/wte"] = 1;  
    $thePermissions = @();
    $MAX_RETRY_CONNECTION = 5;

    $SiteRetryCounter = 0;
    $SiteConnectedAndProcessed = $false;
    

    $Lists = $null;
    $Web = $null;
    $SiteName = "";
    
    while ((-not $SiteConnectedAndProcessed) -and ($SiteRetryCounter -lt $MAX_RETRY_CONNECTION)) {
        try {
            
            $SitePermissions = @();

            #AppCreds = (Get-HelperSPOnline-AppCredential -StoredCredentialName $StoredCredentialName);
            Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant
            Write-Host ("`t" + "Connected");


            $Web = (Get-PnPWeb -Includes Title, ServerRelativePath -ErrorAction Stop);
            $SiteName = ($Web.Title);

            $RoleDefinitions = (Get-SharePointHelper-RoleDefinitions);

            $SiteGroups = (Get-PnPGroup -ErrorAction Stop);
            $SiteGroupsDictionary = @{};
            foreach ($currGroup in $SiteGroups) {
                if ($SiteGroupsDictionary[$currGroup.LoginName] -eq $null) {
                    $SiteGroupsDictionary[$currGroup.LoginName] = (New-Object PSCustomObject -Property @{
                            Title     = ($currGroup.Title);
                            LoginName = ($currGroup.LoginName);
                            Members   = @();
                        });

                    $GroupMembers = @(Get-PnPGroupMember -Group $currGroup -ErrorAction Stop);
                    foreach ($mem in $GroupMembers) {
                        #$SiteGroupsDictionary[$currGroup.LoginName].Members += ($mem.PrincipalType.ToString() + ": " + $mem.Title + " (" + $mem.LoginName + ")");
                        $SiteGroupsDictionary[$currGroup.LoginName].Members += ($mem.PrincipalType.ToString() + ": " + $mem.Title);
                    }
                }
                else {
                    throw [System.Exception]::new("Duplicate group name detected: " + $currGroup.LoginName);
                }
            }

            Write-Host ("`t" + "Site Groups Processed");


            $CommonSitePermissionExportParams = @{
                SiteUrl   = $SiteUrl;
                SiteName  = $SiteName;       
                Delimiter = $Delimiter;                
                Scope     = "Web";
                ItemUrl   = ($Web.ServerRelativePath.DecodedUrl);
            };

            $SitePermissionCollection = (Get-SharePointHelper-GetRoleAssignmentsForObject -SiteUrl $SiteUrl -ObjectType "Web" -RoleDefinitions $RoleDefinitions);
            foreach ($currPerm in $SitePermissionCollection.GetEnumerator()) {
                $Members = @();
                if ($currPerm.Value.AccessType -eq "SharePointGroup") {
                    $Members = ($SiteGroupsDictionary[($currPerm.Key)].Members);
                }

                $SitePermissions += (Get-PermissionAuditEntry @CommonSitePermissionExportParams -AccessType ($currPerm.Value.AccessType) -Name ($currPerm.Value.Title) -Permissions ($currPerm.Value.Permissions -join ";") -Members ($Members -join "{@]"));
            }

            $Lists = (Get-PnPList -Includes Id, BaseType, RootFolder, Title -ErrorAction Stop);

            $thePermissions += ($SitePermissions);
            $SiteConnectedAndProcessed = $true;

            Write-Host ("`t" + "Site Permissions Processed");
            Write-Host "";
        }
        catch {
            
            $SiteRetryCounter++;
            Start-Sleep -Seconds 30;
            Write-Host ("`t" + "Faed to Connect to Site (" + $SiteRetryCounter + ")") -ForegroundColor Red;
        }
    }

    if ($SiteConnectedAndProcessed) {
        if ($ProcessLists) {
        
            Write-Host ("`t" + "Processing Lists");

            $ListsProcessedDictionary = @{};
            $ListsProcessed = $false;        
            $ListRetryCounter = 0;

            while ((-not $ListsProcessed) -and ($ListRetryCounter -lt $MAX_RETRY_CONNECTION)) {
                try {
                
                    $ListPermissions = @();
                    $WebRelativeUrl = ($Web.ServerRelativePath);

                    foreach ($currList in $Lists) {                   

                        $ListRootFolderName = (Get-PnPProperty -ClientObject ($currList.RootFolder) -Property ServerRelativePath);            
                        if ($WebRelativeUrl.DecodedUrl -eq "/") {
                            $ListRootFolderName = ($ListRootFolderName.DecodedUrl.ToString().Substring(1)); #Removes the leading '/'
                        }
                        else {
                            $ListRootFolderName = ($ListRootFolderName.DecodedUrl.ToString().Replace($WebRelativeUrl.DecodedUrl + "/", ""));
                        }                        

                        if (($ExcludedLists[$ListRootFolderName] -eq $null) -and ($ListsProcessedDictionary[$ListRootFolderName] -eq $null)) {
                            
                            Write-Host ("`t`t" + "Processing '" + $ListRootFolderName + "'");
                        
                            $CommonListPermissionExportParams = @{
                                SiteUrl            = $SiteUrl;
                                SiteName           = $SiteName;
                                ListRootFolderName = $ListRootFolderName;
                                ListDisplayName    = ($currList.Title);                                
                                ItemUrl            = ($WebRelativeUrl.DecodedUrl + "/" + $ListRootFolderName);
                                Delimiter          = $Delimiter;
                                Scope              = ($currList.BaseType);
                            };


                            $ListHasUniquePermissions = (Get-PnPProperty -ClientObject $currList -Property HasUniqueRoleAssignments);
                            if ($ListHasUniquePermissions) {                                                   

                                $ListPermissionCollection = (Get-SharePointHelper-GetRoleAssignmentsForObject -SiteUrl $SiteUrl -ObjectType "List" -LibraryGuid ($currList.Id) -RoleDefinitions $RoleDefinitions);
                                foreach ($currPerm in $ListPermissionCollection.GetEnumerator()) {
                                    $Members = @();
                                    if ($currPerm.Value.AccessType -eq "SharePointGroup") {
                                        $Members = ($SiteGroupsDictionary[($currPerm.Key)].Members);
                                    }

                                    $ListPermissions += (Get-PermissionAuditEntry @CommonListPermissionExportParams -AccessType ($currPerm.Value.AccessType) -Name ($currPerm.Value.Title) -Permissions ($currPerm.Value.Permissions -join ";") -Members ($Members -join "{@]"));
                                }
                            }

                            if ($ProcessItems) {

                                Write-Host ("`t`t`t" + "Processing Items");

                                $ContentTypesObj = (Get-PnPContentType -List ("/" + $ListRootFolderName) -ErrorAction Stop);
                                $FolderContentTypes = @{};
                                foreach ($ct in $ContentTypesObj) {
                                    $ContentTypeTree = (Get-ContentTypeTree -ContentType $ct);
                                    if ($ContentTypeTree.NamePath.StartsWith("Item > Folder")) {
                                        $FolderContentTypes[$ct.Id.StringValue] = 1;
                                    }
                                }

                                $RestApiRequestUrl = ($SiteUrl + "/_api/web/lists(guid'" + ($currList.Id) + "')/items?`$select=Id,FileRef,FileLeafRef,Title,HasUniqueRoleAssignments,ContentType/Name,ContentType/Id,Modified&`$expand=ContentType&`$top=5000"); 
                                $ListItemWithUniquePermissions = (Call-SharePointHelper-RestApiMethod $RestApiRequestUrl);

                                $ToProcess = 0;
                                foreach ($currItem in $ListItemWithUniquePermissions) {
                                    if ($currItem.HasUniqueRoleAssignments) {
                                        $ToProcess++;
                                    }
                                }

                                if ($ToProcess -gt 0) {
                                    Write-Host ("`t`t`t" + "To Process: " + $ToProcess);
                        
                                    $ItemCounter = 0;
                                    foreach ($currItem in $ListItemWithUniquePermissions) {
                                        if ($currItem.HasUniqueRoleAssignments) {
                                            $isFolder = ($FolderContentTypes[$currItem.ContentType.Id.StringValue] -ne $null);                                            
                                            $ListItemPermissionCollection = (Get-SharePointHelper-GetRoleAssignmentsForObject -SiteUrl $SiteUrl -ObjectType "Item" -LibraryGuid ($currList.Id) -ItemID ($currItem.Id) -RoleDefinitions $RoleDefinitions);

                                            $Scope = "Item";
                                            if ($isFolder) { $Scope = "Folder"; }

                                            $ItemName = ($currItem.FileLeafRef);
                                            if ($currList.BaseType -eq "GenericList") {
                                                $ItemName = ($currItem.Title);
                                            }

                                            foreach ($currPerm in $ListItemPermissionCollection.GetEnumerator()) {
                                                $Members = @();
                                                if ($currPerm.Value.AccessType -eq "SharePointGroup") {
                                                    $Members = ($SiteGroupsDictionary[($currPerm.Key)].Members);
                                                }

                                                $IsSharingLink = ($currPerm.Value.Title.StartsWith("SharingLinks."));

                                                $ItemPermissionParams = @{
                                                    Scope              = $Scope;
                                                    SiteUrl            = $SiteUrl;
                                                    SiteName           = $SiteName;
                                                    ListRootFolderName = $ListRootFolderName;
                                                    ListDisplayName    = ($currList.Title);
                                                    Name               = ($currPerm.Value.Title);
                                                    Permissions        = ($currPerm.Value.Permissions -join ";");
                                                    Members            = ($Members -join "{@]");
                                                    ContentTypeName    = ($currItem.ContentType.Name);
                                                    IsFolder           = ($isFolder.ToString());
                                                    IsSharingLink      = ($IsSharingLink.ToString());
                                                    ItemID             = ($currItem.Id);
                                                    ItemName           = ($ItemName);
                                                    ItemUrl            = ($currItem.FileRef);
                                                    Delimiter          = $Delimiter;                                        
                                                    AccessType         = ($currPerm.Value.AccessType)
                                                };

                                                $ListPermissions += (Get-PermissionAuditEntry @ItemPermissionParams);
                                            }

                                            $ItemCounter++;
                                            if (($ItemCounter % 100) -eq 0) {
                                                Write-Host ("`t`t`t`t" + "Processed " + $ItemCounter + " of " + $ToProcess);
                                            }
                                        }                                        
                                    }
                                }
                            }

                            $thePermissions += ($ListPermissions);
                            $ListsProcessedDictionary[$ListRootFolderName] = 1;                            
                        }
                    }

                    $ListsProcessed = $true;
                }
                catch {
                    Write-Host ("`t" + "Failed to retrieve permissions for lists (" + $ListRetryCounter + ")") -ForegroundColor Red;
                    Write-Host $_;

                    $ListRetryCounter++;
                    Start-Sleep -Seconds 30;                    
                }
            }
        }
    }
    else {
        throw [System.Exception]::new("Failed to connect to Site after retry attempts exhausted");
    }

    Disconnect-PnPOnline;

    $thePermissions | Out-File -LiteralPath $ExportDirectoryAndName -Encoding utf8 -Append;
    
    Display-ProcessingTime -StartDate $StartTime -AdditionalText ("`t" + "Site Processed in: ");
}

function Export-TenantPermissions {
    [cmdletbinding()]
    param (                                
        [Parameter(Mandatory = $true)][String]$TenantUrl
        , [Parameter(Mandatory = $true)][String]$ExportDirectoryAndName
        , [Parameter(Mandatory = $true)][String]$Delimiter
        , [Parameter(Mandatory = $false)][String]$UrlContains
        , [switch]$ProcessLists
        , [switch]$ProcessItems
        , [switch]$DontRecreateExtract
    );


    #$AppCreds = (Get-HelperSPOnline-AppCredential -StoredCredentialName $StoredCredentialName);
    Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant
    $AllSites = (Get-PnPTenantSite | Where-Object { $_.Url -eq $UrlContains });

    if (-not [String]::IsNullOrEmpty($UrlContains)) {
        $UrlContains = ($UrlContains.ToLower());
        $AllSites = ($AllSites | Where { $_.Url.ToLower().Contains($UrlContains) });
    }

    if (-not $DontRecreateExtract) {
        (
            "Scope" + 
            $Delimiter + "SiteUrl" +
            $Delimiter + "SiteName" +
            $Delimiter + "LibraryInternalName" +
            $Delimiter + "LibraryDisplayName" +
            $Delimiter + "ContentTypeName" +
            $Delimiter + "IsFolder" +
            $Delimiter + "IsSharingLink" +
            $Delimiter + "ItemName" +
            $Delimiter + "ItemUrl" +
            $Delimiter + "ItemID" +
            $Delimiter + "AccessType" +
            $Delimiter + "Name" +
            $Delimiter + "Permissions" +
            $Delimiter + "Members" +
            $Delimiter + "Notes"
        ) | Out-File -LiteralPath $ExportDirectoryAndName -Encoding utf8;
    }

    foreach ($currSite in $AllSites) {
        
        $ProcessSiteParams = @{
            SiteUrl                = ($currSite.Url);
            ExportDirectoryAndName = $ExportDirectoryAndName;
            Delimiter              = $Delimiter;
            ProcessLists           = $ProcessLists;
            ProcessItems           = $ProcessItems;
        };

        Process-PermissionsForSite @ProcessSiteParams;
        Start-Sleep -Seconds 30;
    }
}

$TenantProcessParams = @{
    TenantUrl              = "https://caje77sharepoint.sharepoint.com";
  
    ExportDirectoryAndName = ".\Caje - Permission Export.csv"; #Location on your PC where script wl save report. Should end in .CSV
    Delimiter              = "{#]"; #Used to separate columns in each row of the output. You'll need to separate columns via this value when formatting the report.
    UrlContains            = "https://caje77sharepoint.sharepoint.com/sites/App-3"; #Optional site fter. Comment out this line to run across whole tenancy.
    ProcessLists           = $true;
    ProcessItems           = $true; #Note: ProcessLists must be true for this to work
};

Export-TenantPermissions @TenantProcessParams;