
if (!$page) {
    $url = 'http://conf.splunk.com/sessions/2016-sessions.html'
    $page = Invoke-WebRequest -Uri $url
}

foreach ($link in $page.links) {
#    if (($link.innerHTML -eq 'Recording') -or ($link.innerHTML -eq 'Slides')) {
    if (($link.innerHTML -eq 'Slides')) {
        $filepath = $link | Select-Object -ExpandProperty href
        $filename = ([regex]"(recordings|slides)\/(.*)").match($filepath).Groups[2].Value
        $domain = 'http://conf.splunk.com'
        $url = "$domain$filepath"
        write-host ('downloading ' + $filename)
        $downloadfile = "$env:TEMP\$filename"
        $client = new-object System.Net.WebClient 
        $client.DownloadFile($url,$downloadfile) 
    }
}
