#requires -version 4
<#
.SYNOPSIS
  Exports IE, Firefox, and/or Chrome bookmarks/favorites to a Netscape Bookmark File Format for easy importing into most modern browsers.
.DESCRIPTION
  Provides a method to backup IE, Firefox, and/or Chrome bookmarks into Netscape Bookmark File Format (html).
  Use in cases where you need to re-create a user's local machine profile and they don't have bookmark sync'ign enabled or maybe use multiple browsers that do not provide a sync option.
  This script assumes you're backing up favorites/bookmarks for the user that initiated the script.  All variables containing path information for different browser bookmark files use the environment variable $ENV:USERPROFILE and should be changed if you want to backup another user's bookmarks.
.PARAMETER browser
    Which supported browser(s) do you want to work with;
    Acceptable inputs - IE|FF|Chrome|All or a combination of multiple browser types.  If this paramter has "all" combined with another browser it will error out and fail. (e.g. '.\backup-favorites.ps1 -browser all,ie' will fail)
.PARAMETER htmlOutput
    Location where you want to store the Netscape Bookmark formatted file.
.INPUTS
  <Inputs if any, otherwise state None>
.OUTPUTS
  <Outputs if any, otherwise state None>
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
)


#---------------------------------------------------------[Initialisations]--------------------------------------------------------


#----------------------------------------------------------[Declarations]----------------------------------------------------------



#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
Function <FunctionName> {
  Param ()
  Begin {
    Write-Host '<description of what is going on>...'
  }
  Process {
    Try {
      <code goes here>
    }
    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception)"
      Break
    }
  }
  End {
    If ($?) {
      Write-Host 'Completed Successfully.'
      Write-Host ' '
    }
  }
}
#>
function Get-IEFavorites{
    $Shortcuts = Get-ChildItem -Recurse $ENV:USERPROFILE\favorites -Include *.url
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
<#

$ieItems = Get-IEFavorites
ForEach ($ieItem in $ieItems) {
    Write-Host $ieItem.ShortcutFull
    }
#>
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Script Execution goes here