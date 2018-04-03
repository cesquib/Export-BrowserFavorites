Function Get-IEFavorite {
    <#
        .SYNOPSIS
            Display a list of Internet Shortcuts and folders

        .DESCRIPTION
            Display a list of Internet Shortcuts and folders

        .PARAMETER Name
            Name of the links you wish to find.

        .NOTES
            Author: Boe Prox
            Name: Get-IEFavorite
            Created: 24 Dec 2013
            Version History 
                1.0 -- 24 Dec 2013
                    -Initial Version
                1.1 -- 26 Dec 2013
                    -Directory and File parameters for better filtering
                    -Try/Catch to handle errors with attempting to parse .lnk files
            
        .EXAMPLE
            Get-IEFavorite

            Name     : The WSUS Support Team Blog - Site Home - TechNet Blogs.url
            IsFolder : False
            IsLink   : True
            Url      : http://blogs.technet.com/b/sus/
            Path     : C:\Users\PROXB\Favorites\Links\The WSUS Support Team Blog - Site Home - TechNet Blogs.url

            Name     : WSUS Product Team Blog - Site Home - TechNet Blogs.url
            IsFolder : False
            IsLink   : True
            Url      : http://blogs.technet.com/b/wsus/
            Path     : C:\Users\PROXB\Favorites\Links\WSUS Product Team Blog - Site Home - TechNet Blogs.url 

            Description
            -----------
            Displays all of the favorites
        
        .EXAMPLE
            Get-IEFavorite -Directory

            Description
            -----------
            Displays all folders in Favorites

        .EXAMPLE
            Get-IEFavorite -Name WSUS*

            Description
            -----------
            Displays all Favorites with a name beginning with WSUS
    #>
    #Requires -Version 3.0
    [OutputType('System.IO.InternetShortcutFile','System.IO.InternetShortcutFolder')]
    [cmdletbinding(
        DefaultParameterSetName = 'All'
    )]
    Param (
        [parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string[]]$Name="*",
        [parameter(ParameterSetName= 'Directory')]
        [switch]$Directory,
        [parameter(ParameterSetName= 'File')]
        [switch]$File
    )
    Begin {
        $IEFav =  [Environment]::GetFolderPath('Favorites','None')
        $params = @{
            Recurse = $True
            Path = $IEFav
        }
        If ($PSBoundParameters.ContainsKey('Directory')) {
            $params['Directory'] = $True
        }
        If ($PSBoundParameters.ContainsKey('File')) {
            $params['File'] = $True
        }
        clear
    }
    Process {
        ForEach ($item in $Name) {
            $params['Filter'] = $item
            Get-ChildItem @params | ForEach {
                $object = $_
                Try {
                    If ($object.PSIsContainer) {
                        $Object = [pscustomobject]@{
                            Name = $object.Name
                            IsFolder = [bool]$object.PSIsContainer
                            IsLink = [bool]$False 
                            Url = $Null
                            Path = $Object.FullName
                            Modified = $object.LastWriteTime
                            Created = $object.CreationTime
                        }  
                        $Object.pstypenames.insert(0,'System.IO.InternetShortcutFolder')
                    } Else {
                        $Object = [pscustomobject]@{
                            Name = $object.Name
                            IsFolder = [bool]$object.PSIsContainer
                            IsLink = [bool]$True 
                            Url = ($object | Select-String "^URL").Line.Trim("URL=")
                            Path = $Object.FullName
                            Modified = $object.LastWriteTime
                            Created = $object.CreationTime                            
                        }  
                        $Object.pstypenames.insert(0,'System.IO.InternetShortcutFile')            
                    }  
                    $Object
                } Catch {}
            }
        }
    }
    End {}
}

$favorites = Get-IEFavorite -Directory
ForEach ($favorite in $favorites) {
    Write-Output $favorite
}