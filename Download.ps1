# Getting Details.csv
Try {
    $filePath = "Details.csv"
    $lines = Import-csv -Path $filePath -ErrorAction Stop
} Catch {
    Write-Host -BackgroundColor Red "Details.csv Not Found"
    Write-Host -BackgroundColor Red "Terminating..."
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

# Creating Report
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$reportName = "Report_"+$timestamp+".txt"
New-Item -ItemType File -Path $reportName -Force

foreach ($line in $lines) {
    Write-Host -BackgroundColor Green "------------------------------------------------------"
    $properties = $line | Get-Member -MemberType Properties
    $songDetails = New-Object -TypeName PSObject

    # Converting A CSV Row To PSObject
    for($i=0; $i -lt $properties.Count;$i++)
    {
        $column = $properties[$i]
        $columnvalue = $line | Select -ExpandProperty $column.Name
        $songDetails | Add-Member -MemberType NoteProperty -Name $column.Name -Value $columnvalue
    }
    
    try {
        $uri = 'https://api.spotifydown.com/download/' + $songDetails.'Spotify Track Id'
        $jsonResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    }
    catch {
        Write-Host -BackgroundColor Red "Oops Network Error"
        Write-Host -BackgroundColor Red "Terminating..."
        Exit
    }

    # Retry If Error
    $timelag = 4
    while ($jsonResponse.'success' -eq "false" ) {
        Write-Host "Retrying in $timelag sec..."
        Start-Sleep -Seconds $timelag
        $timelag += 7;
        try {
            $jsonResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        } catch {
            Write-Host -BackgroundColor Red "Oops Network Error"
            Write-Host -BackgroundColor Red "Terminating..."
            Exit
        }
    }

    Write-Host -BackgroundColor Green "Link Received Successfully"

    # Increment The Counter
    $counter++    

    Write-Host -BackgroundColor Green "Song : " + $songDetails.'Song'
    Write-Host -BackgroundColor Green "Count : " + $counter
    
    # Downloading Audio And Image
    Invoke-WebRequest $jsonResponse.'link' -OutFile ".\Downloads\Audio $counter.mp3" 
    Invoke-WebRequest $jsonResponse.'metadata'.'cover' -OutFile ".\Downloads\Image $counter.jpg" 

    # Running FFMPEG
    $album_name = $songDetails.'Album'
    & ".\Tool\ffmpeg.exe" -i ".\Downloads\Audio $counter.mp3" -i ".\Downloads\Image $counter.jpg" -map 0:0 -map 1:0 -codec copy -id3v2_version 3 -metadata album=$album_name ".\Downloads\out $counter.mp3"
    
    # Running Check
    Write-Host -BackgroundColor Cyan "Running Check..."

    try {
        $songN = '.\Downloads\' + $songDetails.'Song'+'.mp3'
        #$songName = $songDetails.'Song' + '.mp3'

        #Get-ChildItem -Recurse -Name ".\Downloads\out $counter.mp3" -ErrorAction Stop
        $objFolder = (New-Object -ComObject Shell.Application).NameSpace(".\Downloads")
        $shellfile = $objFolder.parsename("out $counter.mp3")
        $duration = $objFolder.GetDetailsOf($shellfile, 27) 
        $MetaData = [PSCustomObject]@{
            Duration = $duration
        }
        $minutes, $seconds = ($songDetails.'Time').Split(':')
        $timeSpan = [TimeSpan]::FromMinutes($minutes).Add([TimeSpan]::FromSeconds($seconds))
        $outputString = $timeSpan.ToString('hh\:mm\:ss')

        Write-Host -BackgroundColor Blue "Song Duration : "+$MetaData.'Duration'
        Write-Host -BackgroundColor Blue "Expt Duration : "+$outputString
        Write-Host -BackgroundColor Blue "Difference : "

        if ($outputString -eq $MetaData.'Duration') {
            Write-Output -BackgroundColor Green "All Good!!!"
            $target = ".\Downloads\out $counter.mp3"
            $newnameWExt = $songDetails.'Song'
            $newname = $newnameWExt + '.mp3'
            Rename-Item -Path $target -NewName $newname -ErrorAction Stop
        }
        else {
            Remove-Item ".\Downloads\out $counter.jpg" -ErrorAction Stop
            throw "Audio Was A Chunk..."
        }
        Remove-Item ".\Downloads\Audio $counter.mp3", ".\Downloads\Image $counter.jpg" -ErrorAction Stop
    }
    catch {
        Write-Host -f Red "An error occurred: $_"
        Write-Host -BackgroundColor Red "!!! Error !!!"
        $descrip = "Song :" + $songDetails.'Song' + "`n" + "Link :" + $songDetails.'Spotify Track Id' + "`n"+ "---------------------------------------------"
        Add-Content -Path reportName $descrip
    }
}

# Handle The Report

# Termination Handling
Register-EngineEvent PowerShell.Exiting -Action {
    Set-ItemProperty -Path "HKCU:\SOFTWARE\SpotiDown" -Name "Iter" -Value $counter
}