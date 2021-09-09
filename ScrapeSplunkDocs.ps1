<#
.SYNOPSIS
	Download PDF files for Splunk documentation
.PARAMETER latest
	Download latest if true.  Prompt for version if false.
.EXAMPLE
     .\<scriptname>.ps1 -latest $true
.EXAMPLE
     .\<scriptname>.ps1 -latest $false
#>

param([boolean]$latest=$true)

$VerbosePreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$ProgressPreference = 'SilentlyContinue' # Subsequent calls do not display UI.
$ReleaseNotesOnly = $False
$runningJobLimit = 5   # website appears to penalize with 403 for ~10 minutes if you exceed some number of connections in a certain duration of time.

$StartDate=(GET-DATE)

function Get-SplunkDoc-Versions {
    param($productValue="splunk")
    # scrape the splunk enterprise document main page
    $url = "http://docs.splunk.com/Documentation/$($productValue)"
    $page = Invoke-WebRequest -Uri $url

    # get content of the section which contains options for doc versions
    $versioncontainer = ($page.AllElements | Where-Object {$_.class -match "(version-select-container)"}).innerhtml

    # cleanup option value list so that we have a comma delimmted list of internal/external version names
    $Matches = ([regex]"value=([0-9\.]+)").match($versioncontainer)
    $Values = ($versioncontainer | select-string -pattern '(?smi)value=([0-9\.]+)' -AllMatches).Matches.Value
    $Values = $Values -replace "value=",""
  
    return $Values
}

function Get-SplunkDoc-ProductNames {
    # scrape the splunk enterprise document main page
    $url = 'http://docs.splunk.com/Documentation/Splunk'
    $page = Invoke-WebRequest -Uri $url

    $versionSelectContainer = $page.AllElements | Where-Object {$_.class -match "(version-select-container)"}
    $versionSelectContainer.innerHTML -match "(<SELECT id=product-select>\s+.*<\/SELECT>?)"
    $productSelect = $matches[1]
    $m = $productSelect | select-string -pattern "<OPTION (selected )?value=([^>]+)>([^<]+)<\/OPTION>" -AllMatches

    # transform the array of links to a pipe delimited string
    $Records = @()
    for ($i = 0; $i -le $m.matches.Count -1; $i++)
    {   
        $Record = @{
            OptionValue = $m.Matches.Captures[$i].groups[2].Value     
            OptionName = $m.Matches.Captures[$i].groups[3].Value    
        }
        $Records += New-Object -TypeName PSObject -Property $Record     
    }

    return $Records
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
$scriptpath = $env:temp
$downloadfolder = $scriptpath + '\' + 'downloads'

# if the download folder exists, remove it
if ((Test-Path $downloadfolder -PathType Container) -eq $true) {
    Remove-Item -Path $downloadfolder -Force -Recurse
}
# create the download folder anew
New-Item -Path $scriptpath -Name "downloads" -ItemType "directory" | Out-Null


$ProductNames = Get-SplunkDoc-ProductNames

$ProductNames = $ProductNames | ?{$_.OptionName -ne $null }
$ProductNames = $ProductNames | ?{$_.OptionName -notmatch "\((Legcy|EOL|depricated)\)" }

$SelectedProducts = $ProductNames | Select OptionName, OptionValue | ?{$_.OptionName -notmatch "(\(Legacy\)|\(EOL\)|\(deprecated\)| SDK)"} | Out-GridView -PassThru

if (-not($SelectedProducts)) {
    write-host "User cancelled product selection. Exiting."
    exit
}


$jobs = @()

foreach ($SelectedProduct in $SelectedProducts) {

    write-host "Working on doc download for product $($SelectedProduct.optionName)."

    $downloadFolderProduct = "$($downloadfolder)\$($SelectedProduct.OptionValue)"

    write-host "downloadFolderProduct: $($downloadFolderProduct)"

    # if the download subfolder exists, remove it
    if ((Test-Path $downloadFolderProduct -PathType Container) -eq $true) {
        Remove-Item -Path $downloadFolderProduct -Force -Recurse
    }
    # create the product download folder anew
    New-Item -Path $downloadFolderProduct -ItemType Directory | Out-Null

    # scrape the default splunk enterprise documents site for versions of document containers
    $VersionContainer = Get-SplunkDoc-Versions -productValue $SelectedProduct.OptionValue

    # prompt the user to select which document container they want to download from
    $listitems = $VersionContainer
    if ($latest -eq $false) {
        $selecteditem = Get-SplunkDoc-UserSelection -formtitle "SplunkDocs Downloader" -prompt "Select version to download" -listitems $listitems
        if (!($selecteditem)) { write-verbose "User cancelled action; exiting." ; Exit }
    } else {
        $SelectedItem = $listitems[-1]
    }
    write-host "-User selected document container version $($selecteditem)."

    # download the main page of the selected document container
    $containerUrl = "http://docs.splunk.com/Documentation/$($SelectedProduct.OptionValue)/$($selecteditem)"
    Write-Debug "Downloading main page of document container: $($containerUrl)"
    $containerPage = Invoke-WebRequest -Uri $containerUrl

    # for each manual download link, download file associated with url.
    foreach ($link in $containerPage.links) {
        if ($link.href -like "*/Documentation/$($SelectedProduct.OptionValue)/*/*") {

            $docname = $link.outerText.trim()
            $docname = "$($docname)_v$($selecteditem).pdf"

            if (($ReleaseNotesOnly -eq $True) -and ($docname -notlike "*ReleaseNotes*")) {
                continue                        
            } 

            $downloadfile = "$($downloadFolderProduct)\$($docname)"
            write-host "-Downloading $($docname)."
       
            $docUrl = "http://docs.splunk.com$($link.href)"
            $docPage = Invoke-WebRequest -Uri $docUrl
            $docManualPdfUrl = ($docPage.Links | Where-Object {$_.class -eq "download"} | Where-Object {$_.outerText -match "Download manual as PDF"}).href
            $docManualPdfUrl = "http://docs.splunk.com$($docManualPdfUrl)"
            $docManualPdfUrl = $docManualPdfUrl -replace "&amp;","&"

            # build download url from base + appid + app version
            $Down_URL = $docManualPdfUrl

            # build the file path to download the item to
            $filepath = $downloadfile

            $bSleeping = $false
            do
            {
                $runningJobCount = ($jobs | Get-Job | ?{$_.state -eq "Running"}).count
                if ($runningJobCount -gt $runningJobLimit) {
                    if ($bSleeping -eq $false) {
                        write-host "-running job count is $($runningJobCount). Sleeping until count is less than $($runningJobLimit)."
                    }
                    $bSleeping = $true
                    Start-Sleep -Seconds 1    
                }
            }
            until ($runningJobCount -lt $runningJobLimit)

            $jobs += Start-Job -ScriptBlock { $WebRequest = Invoke-WebRequest -Uri $using:Down_URL -OutFile $using:filepath }

        }
    }
}

<#
$jobs | remove-job -force
#>

# summarize the transaction
$EndDate=(GET-DATE)
$timespan = NEW-TIMESPAN -Start $StartDate -End $EndDate
$elapsed_seconds = [math]::round($timespan.TotalSeconds, 2)
write-host "operation completed in $($elapsed_seconds) seconds. PDF files written to $($downloadfolder)."
