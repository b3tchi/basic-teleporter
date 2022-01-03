# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test08\testdb.accdb"

Remove-Item -Path $testDb -Force

# $objectName = (GetObjectNameFormFileName $sourceFile).split(".")[0]

$app = CreateAccess
$app.Visible = $true
#
$app.NewCurrentDatabase($testDb)

# $docmd = $app.DoCmd

$sourceDir = "$PSScriptRoot.\..\tests\test08"
$sourceFile = "project.json"

$file = GetFileAsList $sourceDir $sourceFile

$file.count

DbProject_Import $file $app

$app.CloseCurrentDatabase()
$app.Quit()
