# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$sourceDir = "$PSScriptRoot.\..\src\modules"

[string]$sourceDir = (Get-Item $sourceDir).FullName

Write-Information "FilePath with without wildcards set" -InformationAction Continue
#no wildcard default*.*
$files1 = GetFilePathsInFolder $sourceDir

$files1.length

Write-Information "FilePath with wildcards set" -InformationAction Continue
#with specified wild cards appending
$files2 = GetFilePathsInFolder $sourceDir "*.bas"
$files2 += GetFilePathsInFolder $sourceDir "*.cls"

$files2.length

Write-Information "general dbModule" -InformationAction Continue
#with layer above
$sourceDir = "$PSScriptRoot.\..\src"

$files = GetFilesList $sourceDir "modules" @("*.bas")
# $files
$files.length


Write-Information "GetSecondarySourceFolder Only" -InformationAction Continue
$secondaryFolders = @("..\\..\\Access_Ext_LibA_Test\\src", "..\\..\\Access_Ext_LibB_Test\\src")

$files = GetFilePathsInSecondaryFolder $sourceDir $secondaryFolders "modules" @("*.bas")
# $files
$files.count


Write-Information "FullBlownModulewithSecondary" -InformationAction Continue
$files = GetFilesList $sourceDir "modules" @("*.bas") $secondaryFolders
# $files
$files.count
