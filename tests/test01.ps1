# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
$sourceDir = "$PSScriptRoot.\..\src"

[string]$sourceDir = (Get-Item $sourceDir).FullName

FolderHasVcsOptionsFile($sourceDir)

#empty
$sourceDir = "$PSScriptRoot.\..\"

[string]$sourceDir = (Get-Item $sourceDir).FullName

FolderHasVcsOptionsFile($sourceDir)
