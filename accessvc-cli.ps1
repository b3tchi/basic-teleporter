Import-Module $PSScriptRoot/lib_spotter.ps1

$workerPath="C:\Users\czJaBeck\Repositories\AccessVCS\Version Control.accda"
# $workerPath="C:\Users\czJaBeck\Repositories\AccessVCS\Version Control_Test.accda"

function export(
  $appPath
  ,$sourceDir
  ){

  #file exits ?
  $appfile = Get-Item $appPath

  if ($null -eq $sourceDir){
    $sourceDir = getSourceDir $appPath
  }

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
  $sourceDir
  ){

  #prepare file names
  $targetName = ReadJsonConfig $sourceDir

  if ($targetName -like "*.accdb"){
    $suffix = ".accdb"
  }elseif ($targetName -like "*.accda"){
    $suffix = ".accda"
  }

  $buildName = $targetName -replace "$suffix", "_build$suffix"

  #create build dir
  $projectDir = Split-Path -Path $sourceDir
  $buildDir = Join-Path $projectDir 'build'
  New-Item -ItemType directory $buildDir -Force

  #remove build file if exists
  $targetFile = Join-Path $buildDir $targetName
  $buildFile = Join-Path $buildDir $buildName

  if(Test-Path -Path $buildFile -PathType Leaf){
    Remove-Item $buildFile
  }

  #start build process
  $app = CreateAccess
  $app.Visible = $true

  $app.NewCurrentDatabase($buildFile)

  $ref = $app.References.AddFromFile($workerPath)

  #run build
  $app.Run("Build_Cli",[ref]$sourceDir,[ref]$true)

  #read & print log
  Get-Content (Join-Path $sourceDir "Build.log")

  RemoveReference $app "MSAccessVCS"

  $app.CloseCurrentDatabase()

  $app.Quit()

  #arcrhive file if needed
  if(Test-Path -Path $targetFile -PathType Leaf){

    # Build archive dir
    $archiveDir = Join-Path $buildDir 'archive'
    New-Item -ItemType directory $archiveDir -Force

    #Archive old build file
    $timestamp = Get-Date -format "yyMMdd-HHmmss"
    $archiveName = $targetName -replace "$suffix", "_$timestamp$suffix"

    Rename-Item $targetFile $archiveName
    Move-Item (Join-Path $buildDir $archiveName) $archiveDir

    Write-Information "saving original file under ... $backupFile" -InformationAction Continue
  }

  #rename build as target file
  Rename-Item $buildFile $targetName

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

function ReadJsonConfig($sourceFile){

  $projectPath = Join-Path $sourceFile "vbe-project.json"

  $json = (Get-Content "$projectPath" -Raw) | ConvertFrom-Json | Select Items

  return $json.Items.FileName

}

function getSourceDir(
  $appPath
  ){

  $appDir = Split-Path -Path $appPath

  #build
  if ((Split-Path -Path $appDir -Leaf) -eq 'build'){
    $projDir = Split-Path -Path $appDir
  }

  #root
  if($null -eq $projDir){
    $projDir = $appDir
  }

  #find project files within project folder
  #NTH if multiple items check project by json options
  $sourceProjFile = (Get-ChildItem -Path $projDir -Recurse -File -Filter "vbe-project.json")[0].FullName

  #if there is not any project file yet use app dir as project dir
  if($null -eq $sourceProjFile){
    $sourceDir = Join-Path $projDir "src"
  }
  else{
    $sourceDir = (Split-Path -Path $sourceProjFile)
  }

  return $sourceDir

}

# export "C:\Users\czJaBeck\OneDrive\DevProjects\AccessKanban\DragAndDropMassacre_Test.accdb"
# export $workerPath

# build "C:\Users\czJaBeck\OneDrive\DevProjects\AccessKanban\DragAndDropMassacre_Test.accdb.src"
# build "$workerPath.src"

#additional tests
# ReadJsonConfig "C:\Users\czJaBeck\OneDrive\DevProjects\AccessKanban\DragAndDropMassacre_Test.accdb.src"
# getSourceDir "C:\Users\czJaBeck\OneDrive\DevProjects\AccessKanban\DragAndDropMassacre_Test.accdb"
# getSourceDir "C:\Users\czJaBeck\OneDrive\DevProjects\AccessKanban\build\DragAndDropMassacre_Test.accdb"

