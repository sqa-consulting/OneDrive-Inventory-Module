Import-Module Microsoft.Online.SharePoint.PowerShell -WarningAction SilentlyContinue
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue

function Invoke-Logging($msg, $code = 0)
{
$colour = switch($code)
{
0 {<#debug#>"Cyan"}
1 {<#info#>"White"}
2 {<#success#>"Green"}
3 {<#warning#>"Yellow"}
4 {<#fail#>"Red"}
}
$output = "{0} :: {1}" -f (Get-Date -Format yyyy-MM-dd__HH-mm-ss), $msg
if ($code -ge 1) {Write-Host $output -ForegroundColor $colour}
Write-Output $output  >> "OneDriveInventory.log"
}

class OneDriveSite
{
    [string]$siteOwner
    [string]$url
    [Object[]] $Files
    [Object] $job
    [PSCredential] $credentials


    OneDriveSite([Object]$site, [Object]$ODc)
    {
    $this.url = $site.url
    $this.siteOwner = $site.owner
    $this.credentials = $ODc.credentials
    }

    asyncReturnFiles()
    {
        Invoke-Logging ("Fetching all results from the background pull of files for site at '{0}'" -f $this.url)
        # This is blocking and should only be called when $this.job.j.isCompleted -eq $True
        $this.Files = $this.job.p.endinvoke([System.IAsyncResult]($this.job.j))
        $code = if ($this.Files.count -gt 0) {2} else {3}
        Invoke-Logging ("Fetched {0} results from the background pull of files for site at '{1}'" -f $this.Files.count, $this.url) $code

    }

    asyncFetchItems()
    {
    Invoke-Logging ("Beginning the asynchronous fetch of files for site '{0}'" -f $this.url)
    $p = [PowerShell]::Create()
    $script = {
        param([PSCredential]$credentials, $url, $siteOwner)
        $connection = Connect-PnPOnline -Url $Url -credential $credentials -ReturnConnection
        $items = Get-PnPListItem -List Documents -PageSize 1000 -Connection $connection -Fields "FileLeafRef","FileDirRef","SMTotalFileStreamSize","Editor","Modified","FSObjType"
        $files = foreach ($item in $items)
            {
            if ($item.fieldvalues.FSObjType -eq 0)
                {
                $ext = if ($Item["FileLeafRef"].contains(".")) { $Item["FileLeafRef"].split(".")[-1] }
                New-Object -type psobject -Property @{
                "SiteOwner" = $siteOwner;
                "Extension" = $ext;
                #"FileName" = $Item["FileLeafRef"];
                "Dir" = $Item["FileDirRef"];
                "SizeInMB" = ($Item["SMTotalFileStreamSize"] / 1MB).ToString("N");
                #"LastModifiedBy" = $Item.FieldValues.Editor.LookupValue;
                "Email" = $Item.FieldValues.Editor.Email;
                "DateModified" = [DateTime]$Item["Modified"]}
                }
            }
        $items.clear()
        [GC]::Collect()
        $files
    }
    $null = $p.AddScript($script).AddParameters(($this.credentials, $this.url, $this.siteOwner))
    $this.job = New-Object -TypeName psobject -property @{"P"=$p;"J"=$p.BeginInvoke()}
    Invoke-Logging ("Background asynchronous fetch of files for site '{0}' has been invoked" -f $this.url) 2
    }
}

class OneDriveClient {
    [PSCredential] $credentials
    [string] $username
    [Object[]] $sites
    [hashtable] $oneDrives = @{}
    [OneDriveSite]$currentOneDrive

    OneDriveClient([PSCredential]$credential, $sharepoint_url) {
        $this.username = $credential.username
        $this.credentials = $credential
        Invoke-Logging "Connecting to sharepoint"
        Try
        {
        Connect-SPOService -Url $sharepoint_url -credential $this.credentials -ErrorAction Stop
        }
        Catch
        {
        Throw "Cannot proceed, credentials did not work for the provided site"
        return
        }
        Invoke-Logging "Connected to sharepoint" 2
        Invoke-Logging "Fetching all sharepoint sites" 2
        $this.fetchSites()
    }

    fetchSites() {
        Invoke-Logging "Fetching a list of sharepoint sites and storing within the object"
        $this.sites = Get-SPOSite -IncludePersonalSite $true -Limit all
        Invoke-Logging ("Identified {0} sharepoint sites, of which {1} are personal OneDrive storage locations" -f $this.sharepointSites().count, $this.oneDriveSites().count) 2
    }

    [Object[]] sharepointSites() {
        Invoke-Logging ("Fetching a list of {0} sharepoint sites" -f $this.sites.count)
        return $this.sites
    }

    [Object[]] oneDriveSites() {
        Invoke-Logging ("Fetching a list of {0} sharepoint sites with OneDrive URLs" -f ($this.sites | Where-Object Url -like "*-my.sharepoint.com/personal/*").count)
        return ($this.sites | Where-Object Url -like "*-my.sharepoint.com/personal/*")
    }

    grantSiteAdmin($site) {
        Invoke-Logging ("Granting Site admin priviledge to {0} for the site '{1}'" -f $this.username, $site.url)
        Set-SPOUser -Site $site.Url -LoginName $this.username -IsSiteCollectionAdmin $true
        Invoke-Logging ("Granted Site admin priviledge to {0} for the site '{1}'" -f $this.username, $site.url) 2
        }

    revokeSiteAdmin($site) {
        Invoke-Logging ("Revoking Site admin priviledge to {0} for the site '{1}'" -f $this.username, $site.url)
        Try
        {
            Set-SPOUser -Site $site.Url -LoginName $this.username -IsSiteCollectionAdmin $false
            Invoke-Logging ("Revoked Site admin priviledge to {0} for the site '{1}'" -f $this.username, $site.url) 2
        }
        Catch
        {
            Invoke-Logging ("Unable to revoke Site admin priviledge granted to {0} for the site '{1}'" -f $this.username, $site.url) 4
        }
    }

    connectOneDrive($site) {
        Invoke-Logging ("Connecting to OneDrive site '{0}'" -f $site.url)
        if ($this.oneDrives.ContainsKey($site.url))
        {
            Invoke-Logging ("Already connected to OneDrive site '{0}'" -f $site.url) 3
            $this.currentOneDrive = $this.oneDrives[$site.url]
        }
        else
        {
            $od = New-Object OneDriveSite($site, $this)
            $this.oneDrives[$site.url] = $od
            $this.currentOneDrive = $this.oneDrives[$site.url]
            Invoke-Logging ("OneDrive object created for site '{0}'" -f $site.url) 2
        }

    }

    ###### Fetch paged files and write to disk
    fetchFiles ([string] $csvPath, [Object[]] $sites )
    {
        Invoke-Logging "Initialising sites"
        foreach ($site in $sites)
        {
            $this.grantSiteAdmin($site)
            $this.connectOneDrive($site)
        }

        Invoke-Logging "Triggering background connect and file inventorying"
        foreach ($od in $this.oneDrives.Values)
        {
        $Od.asyncFetchItems()
        }

        Invoke-Logging "Waiting for background connect and file inventorying to complete"
        ####Progress bar
        $a=$true
        $remaining = 0
        $total = $sites.count
        $lastuser = "ERROR"
        while($a)
        {
        $a = $false
        $x = 0
            foreach ($od in $this.oneDrives.values)
            {
                if ($od.job.j.IsCompleted -eq $false)
                {
                $x ++
                $a = $true
                $lastuser = $OD.siteOwner
                }
            }
        if ($remaining -gt $x) {Invoke-Logging ("Fetched files for {0} of {1} users" -f ($total - $x),$total)}
        $remaining = $x
        [int]$complete = $total - $remaining
        if ($x -eq 1) { Write-Progress -Activity "Fetched files" -ParentId 1 -CurrentOperation ("Waiting on results for user '{0}'" -f $lastuser) -PercentComplete ((100 / $total) * ($complete))}
        else { Write-Progress -Activity "Fetched files" -ParentId 1 -CurrentOperation ("Fetched files for {0} of {1} users, {2} remaining" -f $complete,$total,[int]$remaining) -PercentComplete ((100 / $total) * ($complete))}

        Start-Sleep -Milliseconds 500
        }
        Invoke-Logging "All background jobs complete" 2

        Invoke-Logging "Revoking site admins"
        foreach ($site in $sites)
        {
            $this.revokeSiteAdmin($site)
        }

        Invoke-Logging "Fetching background pipelines and populating objects with file inventories"
        foreach ($OD in $this.OneDrives.Values)
            {
            $OD.asyncReturnFiles()
            }

        Invoke-Logging "Building array to hold files ready for ouput"
        $CSV = foreach ($OD in $this.OneDrives.Values)
            {
            $OD.Files
            }

        Invoke-Logging ("Exporting CSV of {0} files" -f $CSV.count)
        $CSV | Export-Csv $csvPath -NoClobber -NoTypeInformation
        Invoke-Logging ("Exported CSV to '{0}'" -f $csvPath)

    }

    fetchAllFiles ([string] $csvDir, [int] $pagingInterval)
    {
    Invoke-Logging ("Fetching files for OneDrive sites at a paging interval of {0} sites and outputting to '{1}'" -f $pagingInterval, $csvDir)
    $iterations = [math]::Ceiling(($this.oneDriveSites().count / $pagingInterval))
    Invoke-Logging ("Given a paging interval of {0} and a total site count of {1}, I shall require {2} iterations" -f $pagingInterval, $this.oneDriveSites().count, $iterations) 2
    foreach ($x in 1..$iterations)
        {
        Invoke-Logging ("Processing iteration {0}" -f $x)
        Write-Progress -id 1 -Activity "Batch processor" -CurrentOperation ("Fetching files for batch {0} of {1} batches, {2} remaining" -f $x,$iterations,($iterations - $x)) -PercentComplete ((100 / $iterations) * ($x))
        $csvPath = ("{0}\{1}.csv" -f $csvDir, $x).TrimStart("\")
        Invoke-Logging ("Batch CSV path is '{0}'" -f $csvPath) 2
        $first = ($x - 1 ) * $pagingInterval
        $last = ($x * $pagingInterval) - 1 #9,19,29
        Invoke-Logging ("Fetching OneDrive records {0} to {1}" -f $first,$last) 2
        $this.fetchFiles($csvPath,$this.oneDriveSites()[$first .. $last])
        Invoke-Logging "Batch complete, cleaning up to reduce memory footprint further"
        $this.oneDrives = @{}
        $this.currentOneDrive = $null
        [GC]::Collect()
        }
    [GC]::Collect()
    }

    revokeAllPermissions ()
{
    [GC]::Collect()
    foreach ($site in $this.oneDriveSites())
    {
    $this.revokeSiteAdmin($site)
    }
}
}


function Get-OneDriveInventory()
{
<#
 .Synopsis
  Fetch a full file inventory for OneDrive storage within Office365.

 .Description
  Connects to the Office365 Sharepoint instance and fetches a list of all sharepoint sites.  Admin permission is granted
  to all OneDrive sites within sharepoint and a full file inventory is retrieved and stored in CSVs.  Parallel processing
  is used to reduce execution time.

 .Parameter Credential
  The credential object for the connection (Sharepoint Admin)

 .Parameter Site
  The sharepoint site

 .Parameter Page
  The max number of OneDrive sites to enumerate at one time.  Larger page sizes will result in additional memory resource
  constraint.  Smaller page sizes will result in additional run time.

 .Parameter OutputPath
  The base output directory for the resulting CSV files.  Defaults to working dir.

 .Example
   # Fetch inventories in batches of 200.
   Get-OneDriveInventory -Site http://sharepointcustomer-admin.sharepoint.com -Page 200
 .Example
   # Provide a credential object
   Get-OneDriveInventory -Site http://sharepointcustomer-admin.sharepoint.com -Page 200 -Credential $Credentials
 .Example
   # Provide an alternate CSV output path than the current working dir
   Get-OneDriveInventory -Site http://sharepointcustomer-admin.sharepoint.com -Page 200 -OutputPath "C:\Temp"


#>
param(
    [Parameter(Mandatory)]
    [string] $Site,
    [Parameter()]
    [PSCredential] $Credential,
    [Parameter()]
    [int] $Page = 100,
    [Parameter()]
    [string] $OutputPath = ""
    )

if (! $Credential) {$Credential = Get-Credential}
if (! $Credential) { throw "Credentials not provided"}
Try
{
$ODc = New-Object OneDriveClient($Credential, $Site) -ErrorAction Stop
}
Catch
{
Throw "Could not connect to Sharepoint site with the given credentials"
return
}
$timestamp = Get-Date -Format yyyyMMddHHmmss
$OutputPath = $OutputPath + "OneDriveInventory_" + $timestamp
New-Item -ItemType Directory -Path $OutputPath | Out-Null
$ODc.fetchAllFiles($OutputPath,$Page)
Remove-Variable "ODc"
[GC]::Collect()
}

function Remove-OneDriveAdminPermissions()
{
<#
 .Synopsis
  Remove admin permissions for the user across all OneDrive sites.

 .Description
  Connects to the Office365 Sharepoint instance and fetches a list of all sharepoint sites.  Admin permission is removed
  to all OneDrive sites within sharepoint.

 .Parameter Credential
  The credential object for the connection (Sharepoint Admin)

 .Parameter Site
  The sharepoint site

 .Example
   Remove-OneDriveAdminPermissions -Site http://sharepointcustomer-admin.sharepoint.com

#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact='Low')]
param(
    [Parameter(Mandatory)]
    [string] $Site,
    [Parameter()]
    [PSCredential] $Credential
    )

if (! $Credential) {$Credential = Get-Credential}
if (! $Credential) { throw "Credentials not provided"}
Try
{
$ODc = New-Object OneDriveClient($Credential, $Site) -ErrorAction Stop
}
Catch
{
Throw "Could not connect to Sharepoint site with the given credentials"
return
}
$ODc.revokeAllPermissions()
Remove-Variable "ODc"
[GC]::Collect()
}


Export-ModuleMember Get-OneDriveInventory
Export-ModuleMember Remove-SiteAdminPermissions
