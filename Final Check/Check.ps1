# Getting Details.csv
Try {
    $filePath = ".\..\Details.csv"
    $csv = Import-csv -Path $filePath -ErrorAction Stop
} Catch {
    Write-Host -ForegroundColor Red "Details.csv Not Found"
    Write-Host -ForegroundColor Red "Terminating..."
    Exit
}

# Creating Report
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$reportName = "Report_"+$timestamp+".txt"
New-Item -ItemType File -Path $reportName -Force

$counter = 0

foreach($line in $csv)
{ 
    $songDetails = $line

    Write-Host -ForegroundColor Green -BackgroundColor White "------------------------------------------------------"
        
    $counter++

    Write-Host "Item : "$counter
    $songDetails = "Song : "+$songDetails.'Song'+", ID : " +$songDetails.'Spotify Track Id'
    Write-Host $songDetails

    try {
        $newnameWOExt = $songDetails.'Song'
        $cleanedFilename = $newnameWOExt.Split([IO.Path]::GetInvalidFileNameChars()) -join ' '
        $artistname = $songDetails.'Artist' -split ",\s*"
        $cleanedartistname = $artistname[0].Split([IO.Path]::GetInvalidFileNameChars()) -join ' '
        $newname = $cleanedartistname+" - "+ $cleanedFilename + '.mp3'
        $limitnamechar = $newname.Substring(0, [System.Math]::Min(80, $newname.Length))
        $limitnamecharWExt = $limitnamechar + '.mp3'

        $currentLocation = Get-Location
        $currentLocation = Join-Path -Path $currentLocation -ChildPath "..\Downloads"
        $objFolder = (New-Object -ComObject Shell.Application).NameSpace($currentLocation)
        $shellfile = $objFolder.parsename($limitnamecharWExt)
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
            $songLocation = Join-Path -Path $currentLocation -ChildPath $limitnamecharWExt 
            Move-Item -Path $songLocation -Destination '..\Downloads\FinalCheckOK' 
        } else {
            throw "Length Short..."
        }
        Write-Host -ForegroundColor Green "All Good!!!"
    } 
    catch {
        Write-Host -ForegroundColor Red "!!! Error !!!"
        Write-Host -f Red "An error occurred: $_"
        $descrip = "Song :" + $songDetails.'Song' + "`n" + "Link :" + $songDetails.'Spotify Track Id' + "`n"+ "---------------------------------------------"
        Add-Content -Path $reportName $descrip
    }
}