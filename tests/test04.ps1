# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

$ls = [Environment]::NewLine

#default source dir
$sourceFile = "$PSScriptRoot.\..\tests\test04\clsLog.cls"

$objectName = GetObjectNameFormFileName $sourceFile

$contentF = DbModule_ParseSourceFile $sourceFile $objectName

$contentF.Split($ls).count

#default source dir
$sourceFile = "$PSScriptRoot.\..\tests\test04\mod_PSClassCreator.bas"

$objectName = GetObjectNameFormFileName $sourceFile

$contentF = DbModule_ParseSourceFile $sourceFile $objectName

# $contentF.Split($ls).count
# $contentF.count

# $contentF.GetType()

$tempFileName = [System.IO.Path]::GetTempFileName()

$sencd = [Text.Encoding]::Default

[System.IO.File]::WriteAllLines($tempFileName, [string[]]$contentF, $sencd)
