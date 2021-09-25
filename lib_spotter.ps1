# Write-Information 'lib loaded' -InformationAction Continue

$workerPath="$PSScriptRoot\Version Control.accda"
# $workerPath=sitories\AccessVCS\Version Control_Test.accda"
Write-Information "lib loaded $workerPath" -InformationAction Continue

function GetExcel($scriptPath) {

  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

  Write-Information "Trying to attach to - $scriptPath" -InformationAction Continue

  $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath)

  return $TargetApp.Application

}

function CreateAccess(){

  $appAccess =  New-Object -COMObject Access.Application

  return $appAccess

}

function GetAccess($scriptPath) {

  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

  $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath)

  return $TargetApp

}

function GetProject($officeApp){

    $appName = $officeApp.Name
    $appName = $appName.split(" ")[1]

    # $officeApp
    # Write-Debug $appName

    if ($appName -eq "Access"){
        $vbproj = $officeApp.VBE.VBProjects(1)
    }elseif ($appName -eq "Excel"){
        $vbproj = $officeApp.workbooks(1).vbProject
    }

    return $vbproj
}

function GetCodeModule($vbproj, $moduleName) {

    $codeModule = $vbproj.VBComponents($moduleName).CodeModule

    return $codeModule

}

function GetCode($codeModule){

    [string]$code = $codeModule.lines(1,$codeModule.CountOfLines)
    return $code

}

function RemoveCode($codeModule){

    return $codeModule.DeleteLines(1,$codeModule.CountOfLines)

}

function ExportCode($codeModule, $path){

  $COMPONENT_TYPE_MODULE = 1
  $COMPONENT_TYPE_CLASS = 2
  $COMPONENT_TYPE_FORM = 3
  $COMPONENT_TYPE_SPECIAL = 100

  switch($codeModule.Parent.Type){
      $COMPONENT_TYPE_FORM {$suffix = '.frm'}
      $COMPONENT_TYPE_CLASS {$suffix = '.cls'}
      $COMPONENT_TYPE_MODULE {$suffix = '.bas'}
      $COMPONENT_TYPE_SPECIAL {$suffix = '.cls'}
      default{1}
  }

  $moduleFilename = $codeModule.Name + $suffix

  $moduleDestination = [IO.Path]::Combine($path, $moduleFilename)

  $codeModule.Parent.Export($moduleDestination)

}

function RemoveCodeModule($vbProj,$codeModule){

    $vbProj.VBComponents.Remove($codeModule.Parent)
}

function ImportCode($vbProj, $path){
  $COMPONENT_TYPE_MODULE = 1
  $COMPONENT_TYPE_CLASS = 2
  $COMPONENT_TYPE_FORM = 3
  $COMPONENT_TYPE_SPECIAL = 100

    $moduleName = (Get-Item $path).Basename
    Write-Debug $moduleName

    #check if component exists
    $component = $null
    $componentType = -1
    try{
        $component = $vbProj.VBComponents($moduleName)
        switch($component.Type){
            $COMPONENT_TYPE_FORM {$componentType = 1}
            $COMPONENT_TYPE_CLASS {$componentType = 1}
            $COMPONENT_TYPE_MODULE {$componentType = 1}
            $COMPONENT_TYPE_SPECIAL {$componentType = 2}
            default{1}
        }
    }catch{

    }
    #special modules like sheets,workbooks,accessforms

    #exists normal - remove old
    if ($componentType -eq 1){
        RemoveCodeModule $vbProj $component.CodeModule
    }

    #import code into
    $newComponent = $vbProj.VBComponents.Import($path)

    #exists special
    if ($componentType -eq 2){
        $curModule = $component.CodeModule
        $newModule = $newComponent.CodeModule

        $newCode = GetCode $newModule
        # $vbProj.VBComponents.Remove($newComponent)


        RemoveCodeModule $vbProj $newModule
        $newModule = $curModule

        RemoveCode $curModule
        $curModule.AddFromString($newCode)
    }

    return $newModule

}

function ModulesToHashtable($proj){

  [Hashtable]$modules= @{}

  foreach($component in $proj.VBComponents){
    $name = $component.Name
    $code = GetCode $component.CodeModule
    $modules += @{$name=$code}
  }

  return $modules
}

function mergehashtables($htold, $htnew) {
    $keys = $htold.getenumerator() | foreach-object {$_.key}
    $keys | foreach-object {
        $key = $_
        if ($htnew.containskey($key))
        {
            $htold.remove($key)
        }
    }
    $htnew = $htold + $htnew
    return $htnew
}
#just for single level hashtable
function Get-DeepClone_Single {
    # [cmdletbinding()]
    param(
        $InputObject,
        $filter
    )
    process {
      $clone = @{}

      # if ($filter){
      foreach($key in $InputObject.keys) {
          $clone[$key] = $InputObject[$key]
      }

      return $clone
    }
}

#support of multilevel nested hashtable
function Get-DeepClone_Multi {
    [cmdletbinding()]
    param(
        $InputObject
    )
    process
    {
        if($InputObject -is [hashtable]) {
            $clone = @{}
            foreach($key in $InputObject.keys)
            {
                $clone[$key] = Get-DeepClone $InputObject[$key]
            }
            return $clone
        } else {
            return $InputObject
        }
    }
}

function CompareHashtableKeys($sourceht, $targetht){

  foreach($item in $sourceht.keys){
    if(-Not $targetht.ContainsKey($item)){
      $item
    }
  }

}

function CompareHashtableValues($sourceht, $targetht){

  # Get-TypeData $newht.keys

  # Compare-Object $sourceht $targetht -Property Keys

  foreach($item in $sourceht.keys){
    if($targetht.ContainsKey($item)){
      if($sourceht[$item] -ne $targetht[$item]){
        $item
      }
    }
  }

}

function HashToFolder($shadowRepo, $htchanged,$htadded,$htremoved){
  foreach ($key in $htchanged.keys) {
    # Add-Content $shadowRepo$key $htchanged[$key]
    Set-Content $shadowRepo$key $htchanged[$key]
  }

  foreach ($key in $htadded.keys) {
    Add-Content $shadowRepo$key $htadded[$key]
  }

  foreach ($key in $htremoved.keys) {
    Remove-Item $shadowRepo$key
  }

}

function FilterHash($hashTable, $keys){

}
function HashFromFolder($shadowRepo){

  $filesAll=Get-ChildItem -Path "${shadowRepo}*"

  # Write-Information "hff $shadowRepo" -InformationAction Continue

  [Hashtable]$modules= @{}

  $filesAll | ForEach-Object {
    $code = $_ | Get-Content -Raw
    $code = $code -Replace "\r\n$"
    $name = $_.Name

    # Write-Information "files $name" -InformationAction Continue

    $modules += @{$name=$code}

  }

  return $modules

}

function ChangesInVBE($excelFile, $cached){
  $app = GetExcel $excelFile

  $proj = GetProject $app

  $codes = ModulesToHashtable $proj

}


function RepoChanged($dbFile,$ExportLocation,$dteChange) {

    Write-Information "repo changed $dteChange" -InformationAction Continue

    $filesAll=Get-ChildItem -Path "${ExportLocation}*.*"
    $filesChanged = $filesAll | Where-Object {$_.LastWriteTime -gt $dteChange}

    # $filesChanged
    $m = $filesChanged | measure
    $m = $filesAll | measure

    Write-Information "repo changed $filesChanged.Count" -InformationAction Continue
    # Write-Host "RepoChanged"

    #loop all changed files
    $filesChanged | ForEach-Object {
        $code = $_ | Get-Content
        $name = $_.Name

        Write-Information "repo changed $name - $code" -InformationAction Continue

    }

    $accessRun = GetApp $dbFile

    if(!$accessRun) {
        #     $modules.add($module.Name, $modDate)
        Write-Information 'file closed' -InformationAction Continue
    }else{
        #     if($modDate -eq $modLog){
        #
        Write-Information "$accessRun is running" -InformationAction Continue

        # $accessRun.LoadFromText(5, "Testing", $ExportLocation+"Testing.txt")
        #
        #     }else{
        #         # Write-Information "$modDate is newer will be update" -InformationAction Continue
        #
        #     }
    }

}

function export {
  param(
    [Parameter(Mandatory=$true)]$appPath
    ,$sourceDir
    ,[Nullable[boolean]]$doFullExport
    ,[Nullable[boolean]]$sanitizeQuery
  )
  process{

    #parameter defaults
    if ([string]::IsNullOrEmpty($doFullExport)) {$doFullExport = $true}
    if ([string]::IsNullOrEmpty($sanitizeQuery)) {$sanitizeQuery = $true}

    #file exits ?
    $appfile = Get-Item $appPath

    $appPath = $appfile.FullName

    if ($null -eq $sourceDir){
      $sourceDir = getSourceDir $appPath
    }

    # $appfile.Name
    Write-Information "app file $appfile" -InformationAction Continue
    Write-Information "app path $appPath" -InformationAction Continue
    Write-Information "source dir  $sourceDir" -InformationAction Continue
    Write-Information "export $doFullExport" -InformationAction Continue

    $app = CreateAccess
    $app.Visible = $true

    $app.OpenCurrentDatabase($appPath)

    RemoveReference $app "MSAccessVCS"

    if ($workerPath -eq $appPath){
      Write-Information "exporting worker mode ... " -InformationAction Continue
    }else{
      $ref = $app.References.AddFromFile($workerPath)
    }

    #export options
    [boolean]$fullExport = $doFullExport
    [boolean]$printVars = $false
    [string]$localTableWc = "t_*"
    [string]$srcPath = $sourceDir
    [boolean]$buildFromSql = $true
    [boolean]$optSanitizeQuery = $true

    #export execution
    $app.Run("Export_Cli", [ref]$fullExport, [ref]$printVarS, [ref]$localTableWc, [ref]$srcPath, [ref]$buildFromSql, [ref]$optSanitizeQuery)

    #exportlog
    Get-Content (Join-Path $sourceDir "Export.log")

    RemoveReference $app "MSAccessVCS"

    $app.CloseCurrentDatabase()
    $app.Quit()

  }
}
function build(
  $sourceDir
  ){

  [string]$sourceDir = (Get-Item $sourceDir).FullName

  Write-Information "source dir $sourceDir" -InformationAction Continue

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

  # Write-Verbose "$projDir"
  # Write-Information "refrence removed $projDir" -InformationAction Continue

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

  # Write-Information "get project dir $sourceDir" -InformationAction Continue
  return $sourceDir

}
