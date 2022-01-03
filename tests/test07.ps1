# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test07\testdb.accdb"
$sourceFile = "$PSScriptRoot.\..\tests\test07\queries\create.sql"

Remove-Item -Path $testDb

# $objectName = (GetObjectNameFormFileName $sourceFile).split(".")[0]

$app = CreateAccess
$app.Visible = $true

$app.NewCurrentDatabase($testDb)

# $docmd = $app.DoCmd
$sourceFile = "$PSScriptRoot.\..\tests\test07\queries\create.sql"

DbQuery_Import $sourceFile $app

$sourceFile = "$PSScriptRoot.\..\tests\test07\queries\ddltable.sql"

DbQuery_Import $sourceFile $app
# [int]$moduleTypeAc = 5
# $docmd.Save([int]$moduleTypeAc,$objectName)

$app.CloseCurrentDatabase()
$app.Quit()
