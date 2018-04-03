function Export-IEFavorites{
    Param()
    Try {
      $ieShortcuts = Get-ChildItem -Recurse "$ENV:USERPROFILE\favorites" -Include *.url
      $Shell = New-Object -ComObject WScript.Shell
      foreach ($ieShortcut in $ieShortcuts)
      {
        $Properties = @{
          ieName = $ieShortcut.Name; #filename.ext
          ieFull = $ieShortcut.FullName; #full path with file.ext
          iePath = $ieShortcut.DirectoryName; #full path only.
          ieModified = $ieShortcut.LastWriteTime; #MM/DD/YYY HH:MM:SS
          ieCreated = $ieShortcut.CreationTime; #MM/DD/YYY HH:MM:SS
          ieTarget = $Shell.CreateShortcut($ieShortcut).targetpath #web/lan target path.
        } #=> properties
      New-Object PSObject -Property $Properties
      }#=> shortcut

    } #=>try
    Catch {
      Write-Error -Message "There was an error getting bookmarks from the users profile."
      Write-Verbose -Message "There was an error getting bookmarks from the users profile.`nSystem error message: $($_.Exception.ToString())"
    } #=>catch
    Finally {
      [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null
    } #=>finally

<#
$ieItems = Get-IEFavorites
ForEach ($ieItem in $ieItems) {
    Write-Host $ieItem.ShortcutFull
    }
#>
}
$ieItems = Export-IEFavorites
ForEach ($ieItem in $ieItems) {
    Write-Output "-------------------"
    Write-Output "ieName: $($ieItem.ieName)"
    Write-Output "ieFull: $($ieItem.ieFull)"
    Write-Output "ieSplitPath: $(Split-Path -path $ieItem.iePath -Leaf)"
    Write-Output "iePath: $($ieItem.iePath)"
    Write-Output "ieModified: $($ieItem.ieModified)"
    Write-Output "ieCreated: $($ieItem.ieCreated)"
    Write-Output "ieTarget: $($ieItem.ieTarget)"
}