# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test14\testdb.accdb"

if (Test-Path -Path $testDb){
  Remove-Item -Path $testDb -Force
}

$sourceDir = "$PSScriptRoot.\..\tests\test14"
# $sourceFile = "vbe-project.json"

$file = GetFilesList $sourceDir "tabledefs" $("*.json")

$data = (Get-Content "$PSScriptRoot.\..\lib_data.json") | ConvertFrom-Json
$tdata=$data.TableProperties
$fdata=$data.DbFieldTableProperties
# $file.count

'starting'

#create property
$app = CreateAccess
$app.Visible = $true

'testing - create'

$app.NewCurrentDatabase($testDb)

$file[0]
DbTableDef_Import $file[0] $app $tdata $fdata

$file[1]
DbTableDef_Import $file[1] $app $tdata $fdata

$app.CloseCurrentDatabase()

'exitting'

$app.Quit()
