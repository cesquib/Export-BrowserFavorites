#Path to chrome bookmarks
$pathToJsonFile = "$env:localappdata\Google\Chrome\User Data\Default\Bookmarks"

#Helper vars
$temp = "C:\temp\google-bookmarks.json"
$timestamp = Get-Date -Format yyyymmmdd_hhmmss

$global:bookmarks = @()

#A nested function to enumerate bookmark folders
Function Get-BookmarkFolder {
[cmdletbinding()]
Param(
[Parameter(Position=0,ValueFromPipeline=$True)]
$Node
)

Process 
{

 foreach ($child in $node.children) 
 {
   #get parent folder name
   $parent = $node.Name
   $folder = If (!$node.Folder) {""} Else {$node.Folder}
   $folder = $folder + $parent + "/" 
   $child | Add-Member @{Folder= $folder}
   if ($child.type -eq 'Folder') 
   {
     # Write-Verbose "Processing $($child.Name)"
     Get-BookmarkFolder $child
   }
   else 
   {
        $hash= [ordered]@{
          Folder = $parent
          Name = $child.name
          URL = $child.url
          Path = $child.folder.substring(0,$child.folder.Length-1)
          Added = "{0:yyyyMMddHHmmssfff}" -f [datetime]::FromFileTime(([double]$child.Date_Added)*10)
        }
        #add ascustom object to collection
        $global:bookmarks += New-Object -TypeName PSobject -Property $hash
  } #else url
 } #foreach
 } #process
} #end function

$data = Get-content $pathToJsonFile -Encoding UTF8 | out-string | ConvertFrom-Json

#process top level "folders"
$data.roots.bookmark_bar | Get-BookmarkFolder
$data.roots.other | Get-BookmarkFolder
$data.roots.synced | Get-BookmarkFolder

#create a new JSON file
$empty | Set-Content $temp -Force
'{
"bookmarks":' | Add-Content $temp

#these should be the top level "folders"
$global:bookmarks | ConvertTo-Json | Add-Content $temp

'}' | Add-Content $temp

Write-Verbose $temp

Get-Content $temp -Raw |
ConvertFrom-Json |
select -ExpandProperty bookmarks |
Export-CSV $env:USERPROFILE\Desktop\ChromeBookmarks_$timestamp.csv -NoTypeInformation

Remove-Item $temp