param(
    [Int32]$sno = 1,
    [Int32]$eno = 1
)
if (([Int32]$sno -lt 1) -Or ([Int32]$eno -lt 1)) {
    Write-Host -ForegroundColor Red "Param Was Wrong..."
    Exit
}

function ErrorReportClose {
    param (
        [string]$message = ""
    )
    Write-Host -ForegroundColor Red $message
    Write-Host -ForegroundColor Red "Terminating..."
    Exit
}

# Getting Details.csv
Try {
    $filePath = "Details.csv"
    $lines = Import-csv -Path $filePath -ErrorAction Stop
} Catch {
    ErrorReportClose("Details.csv Not Found")
}

if ($eno -eq 1) { $eno = $lines.length }

# Retriving Registry Variable
# Try {
#     $counter = (Get-ItemProperty -Path "HKCU:\SOFTWARE\SpotiDown" -ErrorAction Stop).Iter
# } Catch {
#     Write-Host "Catch"
#     $counter = 0
#     $RegPath = 'HKEY_CURRENT_USER\SOFTWARE\SpotiDown'
#     [microsoft.win32.registry]::SetValue($RegPath, 'Iter', $counter, [Microsoft.Win32.RegistryValueKind]::DWORD)
# }

$counter = 0

# Setting Headers
$headers = @{
    'Referer' = 'https://spotifydown.com/'
    'Origin' = 'https://spotifydown.com/'
}

# Creating Report
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$reportName = "Report_"+$timestamp+".txt"
New-Item -ItemType File -Path $reportName -Force
$timestampFD = Get-Date -Format "dd/MMM HH:mm:ss"
$des = "Date :" + $timestampFD + "`n" + "Start :" + $sno + "`n"+ "End :" + $eno + "`n" + "`n"
Add-Content -Path $reportName $des 
Write-Host -BackgroundColor Cyan -ForegroundColor Blue $des

for($x=$sno; $x -le $eno;$x++) {
    $songDetails = $lines[([Int32]$x)-1]

    Write-Host -ForegroundColor Green -BackgroundColor White "------------------------------------------------------"
    try {
        $uri = 'https://api.spotifydown.com/download/' + $songDetails.'Spotify Track Id'
        $jsonResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
    }
    catch {
        ErrorReportClose("Oops Network Error")
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
            ErrorReportClose("Oops Network Error")
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
        $descrip = "Song :" + $songDetails.'Song' + "`n" + "Link :" + $songDetails.'Spotify Track Id' + "`n"+ "Command :" + "powershell -executionpolicy bypass -File '.\Download.ps1' -sno " + $x + " -eno " + $x + "`n"+ "---------------------------------------------"
        Add-Content -Path $reportName $descrip
        continue
    } 

    # Running FFMPEG
    $album_name = $songDetails.'Album'
    $album_name = $album_name.Split([IO.Path]::GetInvalidFileNameChars()) -join ' '
    $album_name = $album_name.Substring(0, [System.Math]::Min(40, $album_name.Length))
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
            # Make It Short
            $newnameWExt = $songDetails.'Song'
            $cleanedFilename = $newnameWExt.Split([IO.Path]::GetInvalidFileNameChars()) -join ' '
            $artistname = $songDetails.'Artist' -split ",\s*"
            $cleanedartistname = $artistname[0].Split([IO.Path]::GetInvalidFileNameChars()) -join ' '
            $newname = $cleanedartistname+" - "+ $cleanedFilename
            $limitnamechar = $newname.Substring(0, [System.Math]::Min(80, $newname.Length))
            $limitnamecharWExt = $limitnamechar + '.mp3'
            Rename-Item -Path $target -NewName $limitnamecharWExt -ErrorAction Stop
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
        $descrip = "Song :" + $songDetails.'Song' + "`n" + "Link :" + $songDetails.'Spotify Track Id' + "`n"+ "Command :" + "powershell -executionpolicy bypass -File '.\Download.ps1' -sno " + $x + " -eno " + $x + "`n"+ "---------------------------------------------"
        Add-Content -Path $reportName $descrip
        Remove-Item ".\Downloads\Audio $counter.mp3", ".\Downloads\Image $counter.jpg", ".\Downloads\out $counter.mp3" -ErrorAction SilentlyContinue
    }
}

# Handle The Report

# Termination Handling
# Register-EngineEvent PowerShell.Exiting -Action {
#     Write-Host "Closed"
#     Set-ItemProperty -Path "HKCU:\SOFTWARE\SpotiDown" -Name "Iter" -Value $counter
#     Write-Host "Closed"
# }