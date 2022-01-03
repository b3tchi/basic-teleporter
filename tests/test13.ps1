# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test13\testdb.accdb"
$jsonOrigPath = "$PSScriptRoot.\..\tests\test13\data2.json"
$jsonNewPath = "$PSScriptRoot.\..\tests\test13\data.json"

# if (Test-Path -Path $testDb){
#   Remove-Item -Path $testDb -Force
# }

'starting'

#create property
$app = CreateAccess
$app.Visible = $true

'testing - create'

$app.OpenCurrentDatabase($testDb)

$db=$app.CurrentDb()

$proj=$app.CurrentProject

$json=@{}

$json['TableProperties']=(DbListProperties $db.TableDefs)
$json['QueryProperties']=(DbListProperties $db.QueryDefs)
$json['DbProperties']=(DbListProperties ($db))
$json['DbProjectProperties']=(DbListProperties ($proj))
$json['DbFieldTableProperties']=(DbListPropertiesLevelBellow $db.TableDefs "Fields")
# $json['DbFieldQueryProperties']=(DbListPropertiesLevelBellow $db.QueryDefs "Fields")
#
$jsonOrig = (Get-Content "$jsonOrigPath" -Raw) | ConvertFrom-Json

Merge-Object $jsonOrig ($json | ConvertTo-Json | ConvertFrom-Json)

$jsonOrig | ConvertTo-Json | Format-Json | Out-File "$jsonNewPath"

$app.CloseCurrentDatabase()

'exitting'

$app.Quit()
