<#
.Synopsis
   Scrapes splunk enterprise documentation from website, copies pdf files into categorized folders, then compresses content.
#>


$docrevlevel = '6.1beta'

$StartDate=(GET-DATE)

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

# scrape the splunk enterprise document site
$url = 'http://docs.splunk.com/Documentation/Splunk/latest'
$page = Invoke-WebRequest -Uri $url

foreach ($link in $page.links) {
    if ($link.href -like '*/Documentation/Splunk/*') {
        $extract = ([regex]"Splunk\/(\S+)\/(\S+)\/(\S+)").match($link.href)

        if ($extract.Success -eq $true) { 

            $version = $extract.Groups[1].Value
            $class = $extract.Groups[2].Value
            $doc = $extract.Groups[3].Value
            $url = 'https://docs.splunk.com/index.php?title=Documentation:Splunk:' + $class + ':' + $doc + ':' + $docrevlevel + '&action=pdfbook&version=' + $version + 'product=Splunk'

            $filename = ($link.innerText).Trim() + '.pdf'
            $downloadfile = $downloadfolder + '\' + $filename
            write-host ('downloading ' + $filename)

            $client = new-object System.Net.WebClient 
            $client.DownloadFile($url,$downloadfile) 
            }
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