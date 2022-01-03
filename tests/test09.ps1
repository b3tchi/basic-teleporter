# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test09\testdb.accdb"

if (Test-Path -Path $testDb){
  Remove-Item -Path $testDb -Force
}

$sourceDir = "$PSScriptRoot.\..\tests\test09"
$sourceFile = "proj-properties.json"

$file = GetFileAsList $sourceDir $sourceFile
# $file.count

'testing - create'

#create property
$app = CreateAccess
$app.Visible = $true

$app.NewCurrentDatabase($testDb)

DbProjectProperties_Import $file $app

$app.CloseCurrentDatabase()
$app.Quit()

'testing - update'

#try to eddit
$app = CreateAccess
$app.Visible = $true
$app.OpenCurrentDatabase($testDb)

DbProjectProperties_Import $file $app

$app.CloseCurrentDatabase()
$app.Quit()
