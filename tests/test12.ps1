# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$testDb = "$PSScriptRoot.\..\tests\test12\testdb.accdb"

if (Test-Path -Path $testDb){
  Remove-Item -Path $testDb -Force
}

$sourceDir = "$PSScriptRoot.\..\tests\test12"
$sourceFile = "vbe-references.json"

$file = GetFileAsList $sourceDir $sourceFile
# $file.count

'starting'

#create property
$app = CreateAccess
$app.Visible = $true

'testing - create'

$app.NewCurrentDatabase($testDb)

DbVbeReferences_Import $file $app

$app.CloseCurrentDatabase()
# $app.Quit()

#create property
'testing - edit'
# $app = CreateAccess
# $app.Visible = $true

$app.OpenCurrentDatabase($testDb)

DbVbeReferences_Import $file $app

$app.CloseCurrentDatabase()

'exitting'

$app.Quit()
# 'testing - update'
#
# #try to eddit
# $app = CreateAccess
# $app.Visible = $true
# $app.OpenCurrentDatabase($testDb)
#
# DbProjectProperties_Import $file $app
#
# $app.CloseCurrentDatabase()
# $app.Quit()
