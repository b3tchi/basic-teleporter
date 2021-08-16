Import-Module $PSScriptRoot/lib_spotter.ps1

$workerPath="C:\Users\czJaBeck\Repositories\AccessVCS\Version Control.accda"
# $workerPath="C:\Users\czJaBeck\Repositories\AccessVCS\Version Control_Test.accda"

function export(
  $appPath
  ){

  #file exits ?
  $appfile = Get-Item $appPath

  # $appfile.Name
  Write-Information "exporting file "$appfile.Name -InformationAction Continue

  $app = CreateAccess
  $app.Visible = $true

  $app.OpenCurrentDatabase($appPath)

  RemoveReference $app "MSAccessVCS"

  if ($workerPath -eq $appPath){
    Write-Information "exporting worker mode ... " -InformationAction Continue
  }else{
    $ref = $app.References.AddFromFile($workerFinalPath)
  }

  #export options
  $localTableWc = "t_*"
  $printVars = $false

  #export execution
  $app.Run("Export_Cli", [ref]$true, [ref]$printVars, [ref]$localTableWc)

  #exportlog
  Get-Content "$appPath.src/Export.log"

  RemoveReference $app "MSAccessVCS"

  $app.CloseCurrentDatabase()
  $app.Quit()

}

function build(
  $sourcePath
  ){

  $app = CreateAccess
  $app.Visible = $true

  # TODO create backup
  if ($sourcePath -like "*.accdb.src"){
    $suffix = ".accdb"
  }elseif ($sourcePath -like "*.accda.src"){
    $suffix = ".accda"
  }

  $sourceFile = $sourcePath -replace "$suffix.src", $suffix

  # Write-Information "building file form source ... $sourcePath" -InformationAction Continue
  # Write-Information "building file form source ... $sourceFile" -InformationAction Continue
  # Write-Information "building file form source ... $suffix" -InformationAction Continue
  # Write-Information "building file form source ... $workerPath" -InformationAction Continue

  # $sourceDir = Split-Path -Path $sourcePath

  #building worker script itself
  if($sourceFile -eq $workerPath){
    $sourceFile = $sourcePath -replace "$suffix.src", "_build$suffix"

    Write-Information "Building addin file under ... $sourceFile" -InformationAction Continue
  }

  #backup file if already exists
  if(Test-Path -Path $sourceFile -PathType Leaf){

    $src = (Get-Item "$sourceFile").Name
    $timestamp = Get-Date -format "yyMMdd-HHmmss"
    $backupFile = $src -replace "$suffix", "_$timestamp$suffix"

    Rename-Item $sourceFile $backupFile
    Write-Information "saving original file under ... $backupFile" -InformationAction Continue
  }


  $app.NewCurrentDatabase($sourceFile)

  $ref = $app.References.AddFromFile($workerPath)

  $app.Run("Build_Cli",[ref]$sourcePath,[ref]$true)

  RemoveReference $app "MSAccessVCS"

  #read log
  Get-Content "$sourcePath/Build.log"

  $app.CloseCurrentDatabase()

  $app.Quit()

}

function RemoveReference(
  $app
  ,[string]$refName
  ){

  $refs = $app.References

  foreach ($ref in $refs) {

    if ($ref.Name -eq $refName){
      Write-Information "refrence removed $refName" -InformationAction Continue
      $refs.Remove($ref)
    }

  }

}

function RenameWorker(
  $appPath
  ){
  if ($workerPath -eq $appPath){

    $workerName = (Get-Item $workerPath).Name
    $workerDir = Split-Path -Path $workerPath

    # Write-Information "rename project ... $workerName ... $workerDir" -InformationAction Continue

    $workerFinalPath = Join-Path -Path $workerDir -ChildPath "Temp_$workerName"

    Copy-Item $workerPath -Destination $workerFinalPath

    # Write-Information "need to rename project ... $workerFinalPath" -InformationAction Continue
    # Get-Item $workerFinalPath

    # Remove-Item $workerFinalPath
    # $proj = GetProject $app
    # $proj.Name = "MSAccessVCS-lib"
  }
  else {
    $workerFinalPath = $workerPath
  }
}

export "C:\Users\czJaBeck\OneDrive\DevProjects\AccessKanban\DragAndDropMassacre_Test.accdb"
# export $workerPath

build "C:\Users\czJaBeck\OneDrive\DevProjects\AccessKanban\DragAndDropMassacre_Test.accdb.src"
# build "$workerPath.src"

