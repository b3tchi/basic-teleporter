# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test05\testdb.accdb"
$sourceFile = "$PSScriptRoot.\..\tests\test05\mod_PSClassCreator.bas"

Remove-Item -Path $testDb

$objectName = (GetObjectNameFormFileName $sourceFile).split(".")[0]

$app = CreateAccess
$app.Visible = $true

$app.NewCurrentDatabase($testDb)

$proj = GetProject $app
$docmd = $app.DoCmd

DbModule_Import $sourceFile $proj $docmd

[int]$moduleTypeAc = 5

$docmd.Save([int]$moduleTypeAc,$objectName)

$app.CloseCurrentDatabase()
$app.Quit()
