#Requires -Version 7.0
# Make sure to create your secret.json file before running the script.
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.Read.All to enumerate all users in the tenant
#    Sites.ReadWrite.All to return all the item sharing details
#    Files.Read.All to examine all fles in each drive
#    (optional) Directory.Read.All to obtain a domain list and check whether an item is shared externally

[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$ExpandFolders,[int]$depth)

#$ErrorView = 'DetailedView'
function processChildren {

    Param(
    #Graph User object
    [Parameter(Mandatory=$true)]$Drive,
    #URI for the drive
    [Parameter(Mandatory=$true)][string]$URI,
    #Use the ExpandFolders switch to specify whether to expand folders and include their items in the output.
    [switch]$ExpandFolders,
    #Use the Depth parameter to specify the folder depth for expansion/inclusion of items.
    [int]$depth)

    $URI = "$URI/children"
    $children = @()
    #fetch children, make sure to handle multiple pages
    do {
        $result = Invoke-GraphApiRequest -Uri "$URI" -Verbose:$VerbosePreference
        $URI = $result.'@odata.nextLink'
        $children += $result
    } while ($URI)
    if (!$children) { Write-Verbose "No items found for $($drive.id), skipping..."; continue }

    #handle different children types
    $output = @()
    $cFolders = $children.value | ? {$_.Folder}
    $cFiles = $children.value | ? {$_.File} #doesnt return notebooks
    $cNotebooks = $children.value | ? {$_.package.type -eq "OneNote"}

    #Process Folders
    foreach ($folder in $cFolders) {
        $output += (processFolder -Drive $Drive -folder $folder -ExpandFolders:$ExpandFolders -depth $depth -Verbose:$VerbosePreference)
    }

    #Process Files
    foreach ($file in $cFiles) {
        $output += (processFile -Drive $Drive  -file $file -Verbose:$VerbosePreference)
    }

    #Process Notebooks
    foreach ($notebook in $cNotebooks) {
        $output += (processFile -Drive $Drive  -file $notebook -Verbose:$VerbosePreference)
    }

    return $output
}

function processFolder {

    Param(
    #Graph User object
    [Parameter(Mandatory=$true)]$Drive,
    #Folder object
    [Parameter(Mandatory=$true)]$folder,
    #Use the ExpandFolders switch to specify whether to expand folders and include their items in the output.
    [switch]$ExpandFolders,
    #Use the Depth parameter to specify the folder depth for expansion/inclusion of items.
    [int]$depth)

    #prepare the output object
    $fileinfo = New-Object psobject
    $fileinfo | Add-Member -MemberType NoteProperty -Name "WebPath" -Value $Drive.webUrl
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $folder.name
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "Folder"
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Shared" -Value (&{If($folder.shared) {"Yes"} Else {"No"}})

    #if the Shared property is set, fetch permissions
    if ($folder.shared) {
        $permlist = getPermissions $Drive.id $folder.id -Verbose:$VerbosePreference

        #Match user entries against the list of domains in the tenant to populate the ExternallyShared value
        $regexmatches = $permlist | % {if ($_ -match "\(?\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*\)?") {$Matches[0]}}
        if ($permlist -match "anonymous") { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
        else {
            if (!$domains) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No domain info" }
            elseif ($regexmatches -notmatch ($domains -join "|")) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
            else { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No" }
        }
        $fileinfo | Add-Member -MemberType NoteProperty -Name "Permissions" -Value ($permlist -join ",")
    }
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $folder.webUrl

    #Since this is a folder item, check for any children, depending on the script parameters
    if (($folder.folder.childCount -gt 0) -and $ExpandFolders -and ((3 - $folder.parentReference.path.Split("/").Count + $depth) -gt 0)) {
        Write-Verbose "Folder $($folder.Name) has child items"
        $uri = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/items/$($folder.id)"
        $folderItems = processChildren -Drive $Drive -URI $uri -ExpandFolders:$ExpandFolders -depth $depth -Verbose:$VerbosePreference
    }

    #handle the output
    if ($folderItems) { $f = @(); $f += $fileinfo; $f += $folderItems; return $f }
    else { return $fileinfo }
}

function processFile {

    Param(
    #Graph User object
    [Parameter(Mandatory=$true)]$Drive,
    #File object
    [Parameter(Mandatory=$true)]$file)

    #prepare the output object
    $fileinfo = New-Object psobject
    $fileinfo | Add-Member -MemberType NoteProperty -Name "WebPath" -Value $drive.webUrl
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $file.name
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value (&{If($file.package.Type -eq "OneNote") {"Notebook"} Else {"File"}})
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Shared" -Value (&{If($file.shared) {"Yes"} Else {"No"}})

    #if the Shared property is set, fetch permissions
    if ($file.shared) {
        $permlist = getPermissions $drive.id $file.id -Verbose:$VerbosePreference

        #Match user entries against the list of domains in the tenant to populate the ExternallyShared value
        $regexmatches = $permlist | % {if ($_ -match "\(?\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*\)?") {$Matches[0]}}
        if ($permlist -match "anonymous") { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
        else {
            if (!$domains) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No domain info" }
            elseif ($regexmatches -notmatch ($domains -join "|")) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
            else { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No" }
        }
        $fileinfo | Add-Member -MemberType NoteProperty -Name "Permissions" -Value ($permlist -join ",")
    }
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $file.webUrl

    #handle the output
    return $fileinfo
}

function getPermissions {

    Param(
    #Use the UserId parameter to provide an unique identifier for the user object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$DriveId,
    #Use the ItemId parameter to provide an unique identifier for the item object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ItemId)

    #fetch permissions for the given item
    $uri = "https://graph.microsoft.com/beta/drives/$($DriveId)/items/$($ItemId)/permissions"
    $permissions = (Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference).Value

    #build the permissions string
    $permlist = @()
    foreach ($entry in $permissions) {
        #Sharing link
        if ($entry.link) {
            $strPermissions = $($entry.link.type) + ":" + $($entry.link.scope)
            if ($entry.grantedToIdentitiesV2) { $strPermissions = $strPermissions + " (" + (((&{If($entry.grantedToIdentitiesV2.siteUser.email) {$entry.grantedToIdentitiesV2.siteUser.email} else {$entry.grantedToIdentitiesV2.User.email}}) | select -Unique) -join ",") + ")" }
            if ($entry.hasPassword) { $strPermissions = $strPermissions + "[PasswordProtected]" }
            if ($entry.link.preventsDownload) { $strPermissions = $strPermissions + "[BlockDownloads]" }
            if ($entry.expirationDateTime) { $strPermissions = $strPermissions + " (Expires on: $($entry.expirationDateTime))" }
            $permlist += $strPermissions
        }
        #Invitation
        elseif ($entry.invitation) { $permlist += $($entry.roles) + ":" + $($entry.invitation.email) }
        #Direct permissions
        elseif ($entry.roles) {
            if ($entry.grantedToV2.siteUser.Email) { $roleentry = $entry.grantedToV2.siteUser.Email }
            elseif ($entry.grantedToV2.User.Email) { $roleentry = $entry.grantedToV2.User.Email }
            #else { $roleentry = $entry.grantedToV2.siteUser.DisplayName }
            else { $roleentry = $entry.grantedToV2.siteUser.loginName } #user claim
            $permlist += $($entry.Roles) + ':' + $roleentry #apparently the email property can be empty...
        }
        #Inherited permissions
        elseif ($entry.inheritedFrom) { $permlist += "[Inherited from: $($entry.inheritedFrom.path)]" } #Should have a Roles facet, thus covered above
        #some other permissions?
        else { Write-Verbose "Permission $entry not covered by the script!"; $permlist += $entry }
    }

    #handle the output
    return $permlist
}

function Renew-Token {
    #prepare the request
    $url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

    $Scopes = New-Object System.Collections.Generic.List[string]
    $Scope = "https://graph.microsoft.com/.default"
    $Scopes.Add($Scope)

    $body = @{
        grant_type = "client_credentials"
        client_id = $appID
        client_secret = $client_secret
        scope = $Scopes
    }

    try {
        Set-Variable -Name authenticationResult -Scope Global -Value (Invoke-WebRequest -Method Post -Uri $url -Debug -Verbose -Body $body -ErrorAction Stop)
        $token = ($authenticationResult.Content | ConvertFrom-Json).access_token
    }
    catch { $_; return }

    if (!$token) { Write-Host "Failed to aquire token!"; return }
    else {
        Write-Verbose "Successfully acquired Access Token"

        #Use the access token to set the authentication header
        Set-Variable -Name authHeader -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application\json'}
    }
}

function Invoke-GraphApiRequest {
    param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Uri,
    [int]$MaxRetries = 5
    )

    if (!$AuthHeader) { Write-Verbose "No access token found, aborting..."; throw }
    
    $retryCount = 0
    $response = $null
    do {
        try { $result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop }
            catch {
            $errResp = $_.Exception
            $statusCode = [int]$errResp.Response.StatusCode
            $errorCode = ($_.ErrorDetails | ConvertFrom-Json).error.code
            $StatusMessage = $errResp.Response.ReasonPhrase
            $statusCode
            $errorCode
            $StatusMessage

            if ($errorCode -match "ResourceNotFound|Request_ResourceNotFound") { Write-Verbose "Resource $uri not found, skipping..."; return } #404, continue
            #also handle 429, throttled (Too many requests)
            elseif ($statusCode -eq 429) {
                $delay = [int]$_.Exception.Response.Headers["Retry-After"] 
                Write-Warning "Encountered 429 (`"TooManyRequests`"). Waiting $delay seconds."
                Start-Sleep -Seconds $delay
            } elseif ($statusCode -eq 503) {
                $delay = 30
                Write-Warning "Encountered 503 error. Waiting $delay seconds."
                Start-Sleep -Seconds $delay
            }
            elseif ($errorCode -eq "BadRequest") { return } #400, we should terminate... but stupid Graph sometimes returns 400 instead of 404
            elseif ($errorCode -eq "Forbidden") { Write-Verbose "Insufficient permissions to run the Graph API call, aborting..."; throw } #403, terminate
            elseif ($errorCode -match "InvalidAuthenticationToken|unauthenticated|activityLimitReached") {
                    Write-Warning "Trying to renew access token.."
                    Renew-Token
                    if (!$AuthHeader) { Write-Verbose "Failed to renew token, aborting..."; throw }
                    else { Write-Verbose "Access Token renewed. Continuing..." ; throw }
            }
            else { $errResp ; throw }
        }
        $retryCount++
    } while ($retryCount -lt $MaxRetries)

    if ($result) {
        if ($result.Content) { ($result.Content | ConvertFrom-Json) }
        else { return $result }
    }
    else { return }
}

#==========================================================================
#Main script starts here
#==========================================================================

$secretsJson = Get-Content -Raw -Path ".\secrets.json" | ConvertFrom-Json # create a secrets.json file in the same folder containing these secrets. Make sure to add this to .gitignore!
$tenantID = $secretsJson.'Sharepoint-ODFB-tenantID' #your tenantID or tenant root domain
$appID = $secretsJson.'Sharepoint-ODFB-appID' #the GUID of your app. For best result, use app with Sites.ReadWrite.All scope granted.
$client_secret = $secretsJson.'Sharepoint-ODFB-client_secret' #client secret for the app

Renew-Token

#Used to determine whether sharing is done externally, needs Directory.Read.All scope.
$domains = (Invoke-GraphApiRequest -uri "https://graph.microsoft.com/v1.0/domains" -Verbose:$VerbosePreference).Value | ? {$_.IsVerified -eq "True"} | select -ExpandProperty Id
#$domains = @("xxx.com","yyy.com")

#Adjust the input parameters
if ($ExpandFolders -and ($depth -le 0)) { $depth = 0 }

$GraphDrives = @()

$uri = "https://graph.microsoft.com/v1.0/sites/getAllSites?top=999"

$result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop

$sites = $result.value | ? {$_.webUrl -notlike "https://*-my.sharepoint.com/*"}
#$sites.count
foreach ($site in $sites) {
    $drivesUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives"

    $drivesResponse = (Invoke-GraphAPIRequest -Uri $drivesUri -Verbose:$VerbosePreference -ErrorAction Stop).value

    foreach ($drive in $drivesResponse) {
        $GraphDrives += $drive | ? {$_.name -ne "Preservation Hold Library"}
    }
}

#Get the drive for each user and enumerate files
$Output = @()
$count = 1; $PercentComplete = 0;
foreach ($drive in $GraphDrives) {

    #Progress message
    $ActivityMessage = "Retrieving data for drive $($drive.webUrl). Please wait..."
    $StatusMessage = ("Processing drive {0} of {1}: {2}" -f $count, @($GraphDrives).count, $drive.webUrl)
    $PercentComplete = ($count / @($GraphDrives).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Check whether the drive is drive provisioned
    $uri = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root"
    $siteDrive = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop

    #If no items in the drive, skip
    if (!$siteDrive -or ($siteDrive.folder.childCount -eq 0)) { Write-Verbose "No items to report on for drive $($drive.webUrl), skipping..."; continue }

    #enumerate items in the drive and prepare the output
    $pOutput = processChildren -Drive $drive -URI $uri -ExpandFolders:$ExpandFolders -depth $depth
    $Output += $pOutput
}

#Return the output
#$Output | select OneDriveOwner,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath | ? {$_.Shared -eq "Yes"} | Ogv -PassThru
$global:varSPSharedItems = $Output | select WebPath,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath | ? {$_.Shared -eq "Yes"}
#$Output | select OneDriveOwner,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPSharedItems.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
#return $global:varSPSharedItems

$downloadsPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path

$CSVPath = $downloadsPath + "\"+ "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPSharedItems.csv"

$Output | ? {$_.Shared -eq "Yes"} | select WebPath,Name,ItemType,ExternallyShared,Permissions,ItemPath | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8 -UseCulture

ii $CSVPath