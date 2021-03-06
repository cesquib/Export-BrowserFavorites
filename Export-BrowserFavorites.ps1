﻿#requires -version 5

<#
.SYNOPSIS
  Exports IE, Firefox, and/or Chrome bookmarks/favorites to a Netscape Bookmark File Format for easy importing into most modern browsers.
.DESCRIPTION
  Provides a method to backup IE, Firefox, and/or Chrome bookmarks into Netscape Bookmark File Format (html).
  Use in cases where you need to re-create a user's local machine profile and they don't have bookmark sync'ign enabled or maybe use multiple browsers that do not provide a sync option.
  This script assumes you're backing up favorites/bookmarks for the user that initiated the script.  All variables containing path information for different browser bookmark files use the environment variable $ENV:USERPROFILE and should be changed if you want to backup another user's bookmarks.
.PARAMETER browser [string]<required>
    Which supported browser(s) do you want to work with;
    Acceptable inputs - IE|FF|Chrome|All or a combination of multiple browser types.  If this paramter has "all" combined with another browser it will assume 'all' (e.g. '.\backup-favorites.ps1 -browser all,ie' will backup ALL supported browsers)
.PARAMETER htmlOutput [string]<required>
    Location where you want to store the Netscape Bookmark formatted file.
.PARAMETER saveCSV [bool]
    If $TRUE this will keep the temporary CSV file created?
.PARAMETER uniqueonly [bool]
    IF $TRUE will attempt a cleanup of bookmarks to not output duplicates.  Duplicate values are based on URL only.
.PARAMETER getprereqs [bool]
    Default: $false
    Will check for 3rd party tools like the SQLite dll.  If not found will automatically download the necessary files.
    We will always attempt to place needed files inside the user's %temp%\PSLib folder.
.PARAMETER userprofile [string]
    default: <null>
    If left empty (default) the script assumes you are running a backup of bookmarks for the user that is running this script.
    Otherwise, please enter the user's full profile path but only the ROOT (e.g. 'C:\users\somename')


.NOTES
  Version:        1.0
  Author:         Mike Esquibel
  Creation Date:  3/26/2018
  Modified Date:  
  Purpose/Change: Initial script development
  Misc:

  Firefox stores bookmarks in an SQLite DB.  This script was testing using the Precompiled Binaries for 64-bit Windows (.Net Framework 4.5) - filename: sqlite-netFx45-binary-bundle-x64-2012-1.0.108.0.zip

  .EXAMPLE

#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------
Param (
  #Script parameters go here
  [ValidateSet("IE","FF","Chrome","All")]
  [Parameter(mandatory=$true)]
  [array] $browser,
  [Parameter(mandatory=$true)]
  [string] $htmloutput,
  [bool] $saveCSV=$false,
  [Parameter(mandatory=$true)]
  [bool] $uniqueonly,
  [bool] $currentuser=$true
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
$tmpCSV = "$ENV:TEMP\bookmark_export--$((Get-Date).ToString("yyyyMMdd")).csv"
$htmlHeader = @'
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<!--This is an automatically generated file.
    It will be read and overwritten.
    Do Not Edit! -->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<Title>Bookmarks</Title>
<H1>Bookmarks</H1>
<DL><p>
'@



#----------------------------------------------------------[Declarations]----------------------------------------------------------

#-----------------------------------------------------------[Functions]------------------------------------------------------------

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

#-----------------------------------------------------------[Execution]------------------------------------------------------------
#The bookmark file is always necessary so we'll create that first and insert the main header.
Try {
  $htmlHeader | Out-File -FilePath $htmlOut -Force -Encoding utf8
} Catch {
  Write-Error -Message "Unable to create bookmarks HTML file located in $htmlout .  Please check the path and make sure you have the proper permissions to write to that location."
  Write-Verbose -Message "Failed to write to 'C:\windows\system32\test.html'. Please check permissions. `n`nSystem error message: $($_.Exception.ToString())"
  Exit
}


if ($browser -contains 'all' -or 'ie') {
  $ieItems = Export-IEFavorites
  ForEach ($ieItem in $ieItems) {
    
  }#=> ForEach ieItem
}#=> if IE/all
