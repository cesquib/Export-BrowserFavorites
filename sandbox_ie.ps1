function Export-IEFavorites{
    [cmdletbinding()]
    Param(
    [Parameter(Position=0,ValueFromPipeline=$True)]
    $ieFavoritesDir
    )
    Process {
        Write-Output "ieFavoritesDir = " $ieFavoritesDir
        $ieNode = Get-ChildItem -Recurse $ieFavoritesDir
        $Shell = New-Object -ComObject WScript.Shell
        ForEach ($ieChild in $ieNode) 
        {
            $Properties = @{
                ieName = $ieChild.Name; #filename.ext
                ieFull = $ieChild.FullName; #full path with file.ext
                iePath = $ieChild.DirectoryName; #full path only.
                ieModified = $ieChild.LastWriteTime; #MM/DD/YYY HH:MM:SS
                ieCreated = $ieChild.CreationTime; #MM/DD/YYY HH:MM:SS
                ieTarget = $Shell.CreateShortcut($ieChild).targetpath;
              } #=> properties
            New-Object PSObject -Property $Properties
        } #=> ForEach ieChild
    
    } #=> process
    End {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null
    }#=> end

}#=> Export-IEFavorites
$favorites = "C:\users\MichaelEsquibel\favorites"
$favorites | Export-IEFavorites
