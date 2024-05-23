$path = ".\Song Details.csv"
$csv = Import-csv -path $path
$counter = 0
foreach($line in $csv)
{ 
    Write-Host "------------------------------------------------------"
    $properties = $line | Get-Member -MemberType Properties
    $songDetails = New-Object -TypeName PSObject
    $counter++
    for($i=0; $i -lt $properties.Count;$i++)
    {
        $column = $properties[$i]
        $columnvalue = $line | Select -ExpandProperty $column.Name
        $songDetails | Add-Member -MemberType NoteProperty -Name $column.Name -Value $columnvalue
    }
    Write-Host "Item : "
    Write-Host $counter
    Write-Host $songDetails
    $text = "Song : " + $songDetails.'Song' +", ID : " +$songDetails.'Spotify Track Id'
    try {
        $songN = '.\Downloads\' + $songDetails.'Song'+'.mp3'
        Get-ChildItem -Recurse -Name $songN -ErrorAction Stop
        $songName = $songDetails.'Song' + '.mp3'
        $objFolder = (New-Object -ComObject Shell.Application).NameSpace("C:\Users\Mrigayan\Downloads\SpotifyDown\Downloads")
        $shellfile = $objFolder.parsename($songName)
        $duration = $objFolder.GetDetailsOf($shellfile, 27) 
        $MetaData = [PSCustomObject]@{
            Duration = $duration
        }
        $minutes, $seconds = ($songDetails.'Time').Split(':')
        $timeSpan = [TimeSpan]::FromMinutes($minutes).Add([TimeSpan]::FromSeconds($seconds))
        $outputString = $timeSpan.ToString('hh\:mm\:ss')
        Write-Output $MetaData.'Duration'
        Write-Output $outputString
        if ($outputString -eq $MetaData.'Duration') {
            Write-Output "No Errors OK!!!"
            Move-Item -Path $songN -Destination '.\Downloads\OK' 
        }
        else {
            Write-Output "!!! Error !!!"
            $Desc = $text + "`n"+ "Expected: " + $MetaData.'Duration' + "`n" + "Actual:" + $outputString + "`n"+ "---------------------------------------------"
            Add-Content -Path '.\Check - Errors.txt' $Desc
        }
    } 
    catch {
        Write-Host -f Yellow "An error occurred: $_"
        Write-Host "File Not Found"
        Add-Content -Path '.\Check - Not Found.txt' $text
    }
}