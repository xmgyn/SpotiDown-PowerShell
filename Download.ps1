$filePath = "Download - Links.txt"
$lines = Get-Content -Path $filePath
$headers = @{
    'Referer' = 'https://spotifydown.com/'
    'Origin' = 'https://spotifydown.com/'
}
$counter = 0
foreach ($line in $lines) {
    $counter += 1
    $jsonResponse = Invoke-RestMethod -Uri $line -Headers $headers -Method Get
    #Write-Host "Link: $line"
    Invoke-WebRequest $jsonResponse.'link' -OutFile ".\Downloads\Audio $counter.mp3" 
    Invoke-WebRequest $jsonResponse.'metadata'.'cover' -OutFile ".\Downloads\Image $counter.jpg" 
    $album_name = $jsonResponse.'metadata'.'album'
    & ".\ffmpeg.exe" -i ".\Downloads\Audio $counter.mp3" -i ".\Downloads\Image $counter.jpg" -map 0:0 -map 1:0 -codec copy -id3v2_version 3 -metadata album=$album_name ".\Downloads\out $counter.mp3"
    $target = ".\Downloads\out $counter.mp3"
    $newnamel = $jsonResponse.'metadata'.'title'
    $newname = $newnamel + '.mp3'
    Rename-Item -Path $target -NewName $newname
    Remove-Item ".\Downloads\Audio $counter.mp3", ".\Downloads\Image $counter.jpg"
}
