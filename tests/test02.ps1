# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$sourceDir = "$PSScriptRoot.\..\src\modules"

[string]$sourceDir = (Get-Item $sourceDir).FullName

#no wildcard default*.*
$files1 = GetFilePathsInFolder $sourceDir

$files1.length

#with specified wild cards
$files2 = GetFilePathsInFolder $sourceDir "*.bas"
$files2 += GetFilePathsInFolder $sourceDir "*.cls"

$files2.length

#with layer above
$sourceDir = "$PSScriptRoot.\..\src"

$files = GetFilesList_DbModule $sourceDir
$files.length

