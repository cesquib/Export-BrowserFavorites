function Get-DesktopShortcuts{
    $Shortcuts = Get-ChildItem -Recurse "C:\users\mesquibel\favorites" -Include *.url
    $Shell = New-Object -ComObject WScript.Shell
    foreach ($Shortcut in $Shortcuts)
    {
        $Properties = @{
        ShortcutName = $Shortcut.Name;
        ShortcutFull = $Shortcut.FullName;
        ShortcutPath = $shortcut.DirectoryName
        Target = $Shell.CreateShortcut($Shortcut).targetpath
        }
        New-Object PSObject -Property $Properties
    }

[Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null
}

$ieItems = Get-DesktopShortcuts
ForEach ($ieItem in $ieItems) {
    Write-Host $ieItem.ShortcutFull
    }

