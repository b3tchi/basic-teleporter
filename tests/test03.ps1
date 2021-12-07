# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$sourceFile = "$PSScriptRoot.\..\src\modules\cls%3CCocat.cls"

GetObjectNameFormFileName $sourceFile

#default source dir
$sourceFile = "$PSScriptRoot.\..\src\modules\cl%25s%3CCo%3Ccat.cls"

GetObjectNameFormFileName $sourceFile

#default source dir
$sourceFile = "$PSScriptRoot.\..\src\modules\cl%22s%3CCo%3Ccat.cls"

GetObjectNameFormFileName $sourceFile
