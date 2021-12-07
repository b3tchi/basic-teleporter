# "$PSScriptRoot.\..\lib_spotter.ps1"

Import-Module "$PSScriptRoot.\..\lib_spotter.ps1"

#default source dir
# $sourceDir = "$PSScriptRoot.\..\src"
$sourceDir = "$PSScriptRoot.\..\..\Access_VCSTesting\src"

$options = GetVCSOptions $sourceDir

# $options.AddinVersion
$options.Options.SecondaryExportFolders
