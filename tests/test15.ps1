# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test15\testdb.accdb"

if (Test-Path -Path $testDb){
  Remove-Item -Path $testDb -Force
}

$sourceDir = "$PSScriptRoot.\..\tests\test15"
# $sourceFile = "vbe-project.json"


$data = (Get-Content "$PSScriptRoot.\..\lib_data.json") | ConvertFrom-Json
$tdata=$data.TableProperties
$fdata=$data.DbFieldTableProperties
# $file.count

'starting'

#create property
$app = CreateAccess
$app.Visible = $true
$app.NewCurrentDatabase($testDb)

'testing - createlocal'

$file = GetFilesList $sourceDir "tbldefs" $("*.json")

$file[0]
DbTableDef_Import $file[0] $app $tdata $fdata

$file[1]
DbTableDef_Import $file[1] $app $tdata $fdata

'testing createrelations'
$filerel = GetFilesList $sourceDir "relations" $("*.json")

$filerel
DbRelation_Import $filerel $app

'exitting'
$app.CloseCurrentDatabase()

$app.Quit()
