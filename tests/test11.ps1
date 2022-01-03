# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test14\testdb.accdb"

if (Test-Path -Path $testDb){
  Remove-Item -Path $testDb -Force
}

$sourceDir = "$PSScriptRoot.\..\tests\test14"
$sourceFile = "t_Table2.json"

# $file.count

'starting'

#create property
$app = CreateAccess
$app.Visible = $true

'testing - create'

$app.NewCurrentDatabase($testDb)

DbVbeProject_Import (Join-Path $sourceDir "t_Table2.json") $app

$app.CloseCurrentDatabase()

'exitting'

$app.Quit()
