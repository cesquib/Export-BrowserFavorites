Add-Type -Path "C:\tools\scripts\lib\x64\sqlite\System.Data.SQLite.dll"
$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection -ArgumentList "Read Only=True;"
$con.ConnectionString = "Data Source=$ENV:USERPROFILE\AppData\Roaming\Mozilla\Firefox\Profiles\fa5bb0fm.default\places.sqlite"# CHANGE THIS 
$con.Open()
$sql = $con.CreateCommand()
$sql.CommandText = "
SELECT a.id AS ID, a.title AS Title, b.url AS URL, datetime(a.dateAdded/1000000,'unixepoch') AS DateAdded, datetime(a.LastModified/1000000,'unixepoch') AS LastModified, parents.title AS folder
FROM moz_bookmarks AS a 
JOIN moz_places AS b ON a.fk = b.id
JOIN moz_bookmarks parents ON parents.id = a.parent"
$adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
#we create the dataset
$data = New-Object System.Data.DataSet
#and then fill the dataset
[void]$adapter.Fill($data)
#we can, of course, then display the first one hundred rows in a grid
ForEach ($entry in $($data.tables[0].Rows | Sort-Object -Property ID)) {

    $entry.ID
    $entry.Title
    $entry.URL
    $entry.DateAdded
    $entry.LastModified
    $entry.Folder

}
$con.Close()
$con.Dispose()