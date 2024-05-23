# Getting Details.csv
Try {
    $filePath = "Details.csv"
    $lines = Import-csv -Path $filePath -ErrorAction Stop
} Catch {
    Write-Host "Details.csv Not Found"
    Write-Host "Terminating..."
    Exit
}

# Retriving Registry Variable
Try {
    $counter = (Get-ItemProperty -Path "HKCU:\SOFTWARE\SpotiDown" -ErrorAction Stop).Iter
} Catch {
    Write-Host "Catch"
    $counter = 0
    $RegPath = 'HKEY_CURRENT_USER\SOFTWARE\SpotiDown'
    [microsoft.win32.registry]::SetValue($RegPath, 'Iter', $counter, [Microsoft.Win32.RegistryValueKind]::DWORD)
}

Write-Host $counter

# Setting Headers
$headers = @{
    'Referer' = 'https://spotifydown.com/'
    'Origin' = 'https://spotifydown.com/'
}

foreach ($line in $lines) {
    Write-Host "------------------------------------------------------"
    $properties = $line | Get-Member -MemberType Properties
    $songDetails = New-Object -TypeName PSObject

    # Increment The Counter
    $counter++

    for($i=0; $i -lt $properties.Count;$i++)
    {
        $column = $properties[$i]
        $columnvalue = $line | Select -ExpandProperty $column.Name
        $songDetails | Add-Member -MemberType NoteProperty -Name $column.Name -Value $columnvalue
    }
    
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

# Termination Handling
Register-EngineEvent PowerShell.Exiting -Action {
    Set-ItemProperty -Path "HKCU:\SOFTWARE\SpotiDown" -Name "Iter" -Value $counter
}
