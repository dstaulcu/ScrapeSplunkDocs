$VerbosePreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$ReleaseNotesOnly = $False

$StartDate=(GET-DATE)

function Get-SplunkDoc-Versions {
    # scrape the splunk enterprise document main page
    $url = 'http://docs.splunk.com/Documentation/Splunk'
    $page = Invoke-WebRequest -Uri $url

    # get content of the section which contains options for doc versions
    $versioncontainer = ($page.AllElements | Where-Object {$_.class -match "(version-select-container)"}).innerhtml

    # cleanup option value list so that we have a comma delimmted list of internal/external version names
    $Matches = ([regex]"value=([0-9\.]+)").match($versioncontainer)
    $Values = ($versioncontainer | select-string -pattern '(?smi)value=([0-9\.]+)' -AllMatches).Matches.Value
    $Values = $Values -replace "value=",""
  
    return $Values
}

function Get-SplunkDoc-UserSelection {
    param($formtitle="SplunkDocs Downloader",$prompt="Please select a version:",$listitems)

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = $formtitle #"SplunkDocs Downloader"
    $objForm.Size = New-Object System.Drawing.Size(350,200) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
        {$x=$objListBox.SelectedItem;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
        {$objForm.Close()}})

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$x=$objListBox.SelectedItem;$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CancelButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(280,20) 
    $objLabel.Text = $prompt # "Please select a version:"
    $objForm.Controls.Add($objLabel) 

    $objListBox = Out-Null
    $objListBox = New-Object System.Windows.Forms.ListBox 
    $objListBox.Location = New-Object System.Drawing.Size(10,40) 
    $objListBox.Size = New-Object System.Drawing.Size(260,20) 
    $objListBox.Height = 80

    foreach ($item in $listitems) {
        [void] $objListBox.Items.Add($item)
    }

    $objForm.Controls.Add($objListBox) 

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    return $objListBox.SelectedItem
}

$client = new-object System.Net.WebClient
$scriptpath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
$catalogfile = $scriptpath + '\' + 'catalog.csv'
$catalog = Import-Csv -Path $catalogfile
$downloadfolder = $scriptpath + '\' + 'downloads'

# if the download folder exists, remove it
if ((Test-Path $downloadfolder -PathType Container) -eq $true) {
    Remove-Item -Path $downloadfolder -Force -Recurse
}
# create the download folder anew
New-Item -Path $scriptpath -Name "downloads" -ItemType "directory" | Out-Null

# scrape the default splunk enterprise documents site for versions of document containers
$VersionContainer = Get-SplunkDoc-Versions

# prompt the user to select which document container they want to download from
$listitems = $VersionContainer
$selecteditem = Get-SplunkDoc-UserSelection -formtitle "SplunkDocs Downloader" -prompt "Select version to download" -listitems $listitems
if (!($selecteditem)) { write-verbose "User cancelled action; exiting." ; Exit }
write-host "User selected document container version $($selecteditem)."

# download the main page of the selected document container
$containerUrl = "http://docs.splunk.com/Documentation/Splunk/$($selecteditem)"
Write-Debug "Downloading main page of document container: $($containerUrl)"
$containerPage = Invoke-WebRequest -Uri $containerUrl

# for each manual download link, download file associated with url.
foreach ($link in $containerPage.links) {
    if ($link.href -like '*/Documentation/Splunk/*/*') {

        $docname = $link.outerText.trim()
        $docname = "$($docname)_v$($selecteditem).pdf"

        if (($ReleaseNotesOnly -eq $True) -and ($docname -notlike "*ReleaseNotes*")) {
            continue                        
        } 

        $downloadfile = $downloadfolder + '\' + $docname
        write-host "Downloading $($docname)."
       
        $docUrl = "http://docs.splunk.com$($link.href)"
        $docPage = Invoke-WebRequest -Uri $docUrl
        $docManualPdfUrl = ($docPage.Links | Where-Object {$_.class -eq "download"} | Where-Object {$_.outerText -match "Download manual as PDF"}).href
        $docManualPdfUrl = "http://docs.splunk.com$($docManualPdfUrl)"
        $docManualPdfUrl = $docManualPdfUrl -replace "&amp;","&"

      
        $client.DownloadFile($docManualPdfUrl,$downloadfile) 

    }
}

# get all available versions after the selected version into an array
$afterselect = $false
$laterversions = @()
foreach ($version in $VersionContainer) {
    $version = $version.split(",")[0]
    if ($version -eq $selecteditem) {
        $afterselect = $true
        continue
    }
    if ($afterselect -eq $true) {
        $laterversions += $version
    }
}


foreach ($laterversion in $laterversions) {

    # download the main page of the selected document container
    $containerUrl = "http://docs.splunk.com/Documentation/Splunk/$($laterversion)"
    Write-Debug "Downloading main page of document container: $($containerUrl)"
    $containerPage = Invoke-WebRequest -Uri $containerUrl

    # for each manual download link, download file associated with url.
    $RelNotesItem = $containerPage.Links | Where-Object {$_.href -like "*/Documentation/Splunk/*/ReleaseNotes*"}
    $docname = $RelNotesItem.outerText.trim()
    $docname = "$($docname)_v$($laterversion).pdf"
    $downloadfile = $downloadfolder + '\' + $docname

    write-host "Downloading $($docname)."
        
    $docUrl = "http://docs.splunk.com$($RelNotesItem.href)"
    $docPage = Invoke-WebRequest -Uri $docUrl
    $docManualPdfUrl = ($docPage.Links | Where-Object {$_.class -eq "download"} | Where-Object {$_.outerText -match "Download manual as PDF"}).href
    $docManualPdfUrl = "http://docs.splunk.com$($docManualPdfUrl)"
    $docManualPdfUrl = $docManualPdfUrl -replace "&amp;","&"     
    $client.DownloadFile($docManualPdfUrl,$downloadfile) 
}


# get a list of files we downloaded
$files = get-childitem -path $downloadfolder

# create category folders into which we can copy manuals
$folders = $catalog | Select-Object -Unique -Property Folder
foreach ($folder in $folders) {
    $folderpath = $downloadfolder + '\' + $folder.Folder
    if ((Test-Path $folderpath -PathType Container) -eq $false) {
        New-Item -Path $downloadfolder -Name $folder.Folder -ItemType "directory" | Out-Null
    }
}

# for each file, copy into appropriate folders then delete from temp location
foreach ($file in $files) {
    foreach ($item in $catalog) {
        if ($file.name -match $item.Document) {
            write-host ('copying ' + $file.name + ' to ' + $item.Folder + ' folder.')
            Copy-Item $file.FullName -Destination ($downloadfolder + '\' + $item.Folder + '\' + $file.Name)
        }
    }
    Remove-Item $file.FullName
}

# compress the data store of manuals
$zipfile = "$($scriptpath)\downloads.zip"
write-host "compressing documents within $($zipfile)"
if ((Test-Path $zipfile) -eq $true) { Remove-Item $zipfile -Force }
Add-Type -Assembly "System.IO.Compression.FileSystem"
[System.IO.Compression.ZipFile]::CreateFromDirectory($downloadfolder, $zipfile)
$zipfileversion = $zipfile -replace ".zip","v$($selecteditem).zip"
Get-Item $zipfile | Rename-Item -NewName $zipfileversion

# remove the temporary downlad folder
Remove-Item $downloadfolder -Force -Recurse

# summarize the transaction
$EndDate=(GET-DATE)
$timespan = NEW-TIMESPAN -Start $StartDate -End $EndDate
$elapsed_seconds = [math]::round($timespan.TotalSeconds, 2)
write-host ('operation completed in ' + $elapsed_seconds + ' seconds!')
