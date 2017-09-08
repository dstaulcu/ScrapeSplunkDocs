$VerbosePreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"

$StartDate=(GET-DATE)


function Get-SplunkDoc-Versions {
    # scrape the splunk enterprise document main page
    $url = 'http://docs.splunk.com/Documentation/Splunk'
    $page = Invoke-WebRequest -Uri $url

    # get content of the section which contains options for doc versions
    $versioncontainer = ($page.AllElements | Where-Object {$_.class -match "(version-select-container)"}).innerhtml

    # cleanup option value list so that we have a comma delimmted list of internal/external version names
    $versioncontainer = $versioncontainer -split "`<option value`=" 
    $versioncontainer = $versioncontainer -match "/option"
    $versioncontainer = $versioncontainer -replace "</option>",""
    $versioncontainer = $versioncontainer -replace "(\s+</select>)",""
    $versioncontainer = $versioncontainer -replace "`"",""
    $versioncontainer = $versioncontainer -replace ">",","
    $versioncontainer = $versioncontainer -replace " selected=",""
    $versioncontainer = $versioncontainer -replace "(\n)",""

    return $versioncontainer
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
$listitems = $VersionContainer -replace "(.*),",""
$selecteditem = Get-SplunkDoc-UserSelection -formtitle "SplunkDocs Downloader" -prompt "Select version to download" -listitems $listitems
if (!($selecteditem)) { write-verbose "User cancelled action; exiting." ; Exit }
$selecteditem = ($VersionContainer -like "*$selecteditem*").split(",")[0]
write-host "User selected document container version $($selecteditem)."

# download the main page of the selected document container
$containerUrl = "http://docs.splunk.com/Documentation/Splunk/$($selecteditem)"
Write-Debug "Downloading main page of document container: $($containerUrl)"
$containerPage = Invoke-WebRequest -Uri $containerUrl

# search version container for instances of document page links

foreach ($link in $containerPage.links) {
    if ($link.href -like '*/Documentation/Splunk/*/*') {

        $docname = $link.outerText.trim()
        $docname = "$($docname).pdf"
        $downloadfile = $downloadfolder + '\' + $docname

        write-host "Downloading $($docname) for version $($selecteditem)."
        
        $docUrl = "http://docs.splunk.com$($link.href)"
        $docPage = Invoke-WebRequest -Uri $docUrl
        $docManualPdfUrl = ($docPage.Links | Where-Object {$_.class -eq "download"} | Where-Object {$_.outerText -match "Download manual as PDF"}).href
        $docManualPdfUrl = "http://docs.splunk.com$($docManualPdfUrl)"
        $docManualPdfUrl = $docManualPdfUrl -replace "&amp;","&"
      
        $client.DownloadFile($docManualPdfUrl,$downloadfile) 

    }
}


$files = get-childitem -path $downloadfolder

# check to see if folders exist. if not create them
$folders = $catalog | Select-Object -Unique -Property Folder
foreach ($folder in $folders) {
    $folderpath = $downloadfolder + '\' + $folder.Folder
    if ((Test-Path $folderpath -PathType Container) -eq $false) {
        New-Item -Path $downloadfolder -Name $folder.Folder -ItemType "directory" | Out-Null
    }
}

foreach ($file in $files) {
    foreach ($item in $catalog) {
        if ($item.Document + '.pdf'-eq $file.Name) {
            write-host ('copying ' + $file.name + ' to ' + $item.Folder + ' folder.')
            Copy-Item $file.FullName -Destination ($downloadfolder + '\' + $item.Folder + '\' + $file.Name)
        }
    }
    Remove-Item $file.FullName
}

write-host ('compressing documents within ' + $scriptpath + "\downloads.zip")
if ((Test-Path ($scriptpath + "\downloads.zip")) -eq $true) {
    Remove-Item ($scriptpath + "\downloads.zip") -Force
}
Add-Type -Assembly "System.IO.Compression.FileSystem"
[System.IO.Compression.ZipFile]::CreateFromDirectory($downloadfolder, $scriptpath + "\downloads.zip")

Remove-Item $downloadfolder -Force -Recurse

$EndDate=(GET-DATE)

$timespan = NEW-TIMESPAN –Start $StartDate –End $EndDate
$elapsed_seconds = [math]::round($timespan.TotalSeconds, 2)

write-host ('operation completed in ' + $elapsed_seconds + ' seconds!')