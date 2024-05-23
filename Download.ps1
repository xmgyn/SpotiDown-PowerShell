param([Int32]$no = 1)

# Getting Details.csv
Try {
    $filePath = "Details.csv"
    $lines = Import-csv -Path $filePath -ErrorAction Stop
} Catch {
    Write-Host -ForegroundColor Red "Details.csv Not Found"
    Write-Host -ForegroundColor Red "Terminating..."
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

# Setting Headers
$headers = @{
    'Referer' = 'https://spotifydown.com/'
    'Origin' = 'https://spotifydown.com/'
}

# Creating Report
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$reportName = "Report_"+$timestamp+".txt"
New-Item -ItemType File -Path $reportName -Force

for($x=$no; $x -le $lines.length;$x++) {
    $songDetails = $lines[$x-1]

    Write-Host -ForegroundColor Green -BackgroundColor White "------------------------------------------------------"
    # $properties = $line | Get-Member -MemberType Properties
    # $songDetails = New-Object -TypeName PSObject

    # # Converting A CSV Row To PSObject
    # for($i=0; $i -lt $properties.Count;$i++)
    # {
    #     $column = $properties[$i]
    #     $columnvalue = $line | Select -ExpandProperty $column.Name
    #     $songDetails | Add-Member -MemberType NoteProperty -Name $column.Name -Value $columnvalue
    # }
    
    try {
        $uri = 'https://api.spotifydown.com/download/' + $songDetails.'Spotify Track Id'
        $jsonResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
    }
    catch {
        Write-Host -ForegroundColor Red "Oops Network Error"
        Write-Host -ForegroundColor Red "Terminating..."
        Exit
    }

    $status = $jsonResponse.'success'

    # Retry If Error
    $timelag = 4
    while (!$status) {
        Write-Host "Retrying in $timelag sec..."
        Start-Sleep -Seconds $timelag
        $timelag += 7;
        try {
            $jsonResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        } catch {
            Write-Host -ForegroundColor Red "Oops Network Error"
            Write-Host -ForegroundColor Red "Terminating..."
            Exit
        }
    }

    Write-Host -ForegroundColor Green "Link Received Successfully"

    # Increment The Counter
    $counter++    
    Write-Host -ForegroundColor Green "Song : "$songDetails.'Song'
    Write-Host -ForegroundColor Green "Count : "$counter
    
    # Downloading Audio And Image
    try {
        Write-Host -ForegroundColor Magenta -BackgroundColor White "Downloading Song..."
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest $jsonResponse.'link' -OutFile ".\Downloads\Audio $counter.mp3" -ErrorAction Stop
        Write-Host -ForegroundColor Magenta -BackgroundColor White "Downloading Cover..."
        Invoke-WebRequest $jsonResponse.'metadata'.'cover' -OutFile ".\Downloads\Image $counter.jpg" -ErrorAction Stop
    }
    catch {
        Write-Host -ForegroundColor Red "!!! Connection Closed !!!"
        $descrip = "Song :" + $songDetails.'Song' + "`n" + "Link :" + $songDetails.'Spotify Track Id' + "`n"+ "---------------------------------------------"
        Add-Content -Path $reportName $descrip
        continue
    } 

    # Running FFMPEG
    $album_name = $songDetails.'Album'
    & ".\Tool\ffmpeg.exe" -i ".\Downloads\Audio $counter.mp3" -i ".\Downloads\Image $counter.jpg" -map 0:0 -map 1:0 -codec copy -id3v2_version 3 -metadata album=$album_name ".\Downloads\out $counter.mp3"
    
    # Running Check
    Write-Host -ForegroundColor Cyan "Running Check..."

    try {
        $currentLocation = Get-Location
        $currentLocation = Join-Path -Path $currentLocation -ChildPath "Downloads"
        $objFolder = (New-Object -ComObject Shell.Application).NameSpace($currentLocation)
        $shellfile = $objFolder.parsename("out $counter.mp3")
        $duration = $objFolder.GetDetailsOf($shellfile, 27) 
        $MetaData = [PSCustomObject]@{
            Duration = $duration
        }
        $minutes, $seconds = ($songDetails.'Time').Split(':')
        $timeSpan = [TimeSpan]::FromMinutes($minutes).Add([TimeSpan]::FromSeconds($seconds))
        $outputString = $timeSpan.ToString('hh\:mm\:ss')
        $hours0, $minutesO, $secondsO = ($MetaData.'Duration').Split(':')
        $difference = (([int]$minutesO)*60 + $secondsO)-(([int]$minutes)*60 + $seconds)

        Write-Host -ForegroundColor Yellow -BackgroundColor DarkGreen "Song Duration : "$MetaData.'Duration'
        Write-Host -ForegroundColor Yellow -BackgroundColor DarkGreen "Expt Duration : "$outputString
        Write-Host -ForegroundColor Yellow -BackgroundColor DarkGreen "Difference : "$difference

        if (($outputString -eq $MetaData.'Duration') -Or (($difference -le 5) -And ($difference -ge -5))) {
            Write-Host -ForegroundColor Green -BackgroundColor Magenta "Audio Was Correct..."
            $target = ".\Downloads\out $counter.mp3"
            $newnameWExt = $songDetails.'Song'
            $cleanedFilename = $newnameWExt.Split([IO.Path]::GetInvalidFileNameChars()) -join ' '
            $artistname = $songDetails.'Artist' -split ",\s*"
            $cleanedartistname = $artistname[0].Split([IO.Path]::GetInvalidFileNameChars()) -join ' '
            $newname = $cleanedartistname+" - "+ $cleanedFilename + '.mp3'
            $limitnamechar = $newname.Substring(0, [System.Math]::Min(80, $newname.Length))
            Rename-Item -Path $target -NewName $limitnamechar -ErrorAction Stop
        }
        else {
            Remove-Item ".\Downloads\out $counter.mp3" -ErrorAction Stop
            throw "Audio Was A Chunk..."
        }
        Write-Host -ForegroundColor Magenta "Cleaning..."
        Remove-Item ".\Downloads\Audio $counter.mp3", ".\Downloads\Image $counter.jpg" -ErrorAction Stop
        Write-Host -ForegroundColor Green "All Good!!!"
    }
    catch {
        Write-Host -ForegroundColor Red "!!! Error !!!"
        Write-Host -f Red "An error occurred: $_"
        $descrip = "Song :" + $songDetails.'Song' + "`n" + "Link :" + $songDetails.'Spotify Track Id' + "`n"+ "---------------------------------------------"
        Add-Content -Path $reportName $descrip
        Remove-Item ".\Downloads\Audio $counter.mp3", ".\Downloads\Image $counter.jpg", ".\Downloads\out $counter.mp3"
    }
}

# Handle The Report

# Termination Handling
Register-EngineEvent PowerShell.Exiting -Action {
    Write-Host "Closed"
    Set-ItemProperty -Path "HKCU:\SOFTWARE\SpotiDown" -Name "Iter" -Value $counter
    Write-Host "Closed"
}