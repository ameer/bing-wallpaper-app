$url = "https://bingwallpaper.microsoft.com/api/BWC/getHPImages?screenWidth=1920&screenHeight=1080&env=live"
$json = Invoke-RestMethod -Uri $url -Method Get
$urlbase = $json.images[0].urlbase
Add-Type -AssemblyName System.Windows.Forms
$width = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Size.Width
$height = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Size.Height
$urlbase = $urlbase -replace 'w=\d+', "w=$width" -replace 'h=\d+', "h=$height"
$imageUrl = $urlbase + "&format=jpg"
$title = $json.images[0].enddate
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($imageUrl, "C:\Users\$Env:UserName\Pictures\$title.jpg")

$filePath = "C:\Users\$Env:UserName\Pictures\$title.jpg"
$regKey = "HKCU:\Control Panel\Desktop"
Set-ItemProperty -Path $regKey -Name Wallpaper -Value $filePath
rundll32.exe user32.dll, UpdatePerUserSystemParameters
