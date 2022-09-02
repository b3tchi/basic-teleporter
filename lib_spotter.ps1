# Write-Information 'lib loaded' -InformationAction Continue

$workerPath="$PSScriptRoot\Version Control.accda"
$ds = [IO.Path]::DirectorySeparatorChar
$ls = [Environment]::NewLine
$sencd = [Text.Encoding]::Default
# $workerPath=sitories\AccessVCS\Version Control_Test.accda"
Write-Information "lib loaded $workerPath" -InformationAction Continue

function GetExcel($scriptPath)
{

  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

  Write-Information "Trying to attach to - $scriptPath" -InformationAction Continue

  $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath)

  return $TargetApp.Application

}

function CreateAccess()
{

  $appAccess =  New-Object -COMObject Access.Application

  return $appAccess

}

function GetAccess($scriptPath)
{

  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

  $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath)

  return $TargetApp

}

function GetProject($officeApp)
{

  $appName = $officeApp.Name
  $appName = $appName.split(" ")[1]

  # $officeApp
  # Write-Debug $appName

  if ($appName -eq "Access")
  {
    $vbproj = $officeApp.VBE.VBProjects(1)
  } elseif ($appName -eq "Excel")
  {
    $vbproj = $officeApp.workbooks(1).vbProject
  }

  return $vbproj
}

function GetCodeModule($vbproj, $moduleName)
{

  $codeModule = $vbproj.VBComponents($moduleName).CodeModule

  return $codeModule

}

function GetCode($codeModule)
{

  [string]$code = $codeModule.lines(1,$codeModule.CountOfLines)
  return $code

}

function RemoveCode($codeModule)
{

  return $codeModule.DeleteLines(1,$codeModule.CountOfLines)

}

function ExportCode($codeModule, $path)
{

  $COMPONENT_TYPE_MODULE = 1
  $COMPONENT_TYPE_CLASS = 2
  $COMPONENT_TYPE_FORM = 3
  $COMPONENT_TYPE_SPECIAL = 100

  switch($codeModule.Parent.Type)
  {
    $COMPONENT_TYPE_FORM
    {$suffix = '.frm'
    }
    $COMPONENT_TYPE_CLASS
    {$suffix = '.cls'
    }
    $COMPONENT_TYPE_MODULE
    {$suffix = '.bas'
    }
    $COMPONENT_TYPE_SPECIAL
    {$suffix = '.cls'
    }
    default
    {1
    }
  }

  $moduleFilename = $codeModule.Name + $suffix

  $moduleDestination = [IO.Path]::Combine($path, $moduleFilename)

  $codeModule.Parent.Export($moduleDestination)

}

function RemoveCodeModule($vbProj,$codeModule)
{

  $vbProj.VBComponents.Remove($codeModule.Parent)
}

function ImportCode($vbProj, $path, $moduleName)
{
  $COMPONENT_TYPE_MODULE = 1
  $COMPONENT_TYPE_CLASS = 2
  $COMPONENT_TYPE_FORM = 3
  $COMPONENT_TYPE_SPECIAL = 100

  # $moduleName = (Get-Item $path).Basename
  # Write-Debug $moduleName

  #check if component exists
  $component = $null
  $componentType = -1
  try {
    $component = $vbProj.VBComponents($moduleName)
    switch($component.Type) {
      $COMPONENT_TYPE_FORM {$componentType = 1 }
      $COMPONENT_TYPE_CLASS {$componentType = 1 }
      $COMPONENT_TYPE_MODULE {$componentType = 1 }
      $COMPONENT_TYPE_SPECIAL {$componentType = 2 }
      default {1 }
    }
  } catch {
  }
  #special modules like sheets,workbooks,accessforms

  #exists normal - remove old
  if ($componentType -eq 1) {
    RemoveCodeModule $vbProj $component.CodeModule
  }

  #import code into
  $newComponent = $vbProj.VBComponents.Import($path)

  #exists special
  if ($componentType -eq 2) {
    $curModule = $component.CodeModule
    $newModule = $newComponent.CodeModule

    $newCode = GetCode $newModule

    RemoveCodeModule $vbProj $newModule
    $newModule = $curModule

    RemoveCode $curModule
    $curModule.AddFromString($newCode)
  }

  return $newModule

}

function ModulesToHashtable($proj)
{

  [Hashtable]$modules= @{}

  foreach($component in $proj.VBComponents)
  {
    $name = $component.Name
    $code = GetCode $component.CodeModule
    $modules += @{$name=$code}
  }

  return $modules
}

function mergehashtables($htold, $htnew)
{
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
function Get-DeepClone_Single
{
  # [cmdletbinding()]
  param(
    $InputObject,
    $filter
  )
  process
  {
    $clone = @{}

    # if ($filter){
    foreach($key in $InputObject.keys)
    {
      $clone[$key] = $InputObject[$key]
    }

    return $clone
  }
}

#support of multilevel nested hashtable
function Get-DeepClone_Multi
{
  [cmdletbinding()]
  param(
    $InputObject
  )
  process
  {
    if($InputObject -is [hashtable])
    {
      $clone = @{}
      foreach($key in $InputObject.keys)
      {
        $clone[$key] = Get-DeepClone $InputObject[$key]
      }
      return $clone
    } else
    {
      return $InputObject
    }
  }
}

function CompareHashtableKeys($sourceht, $targetht)
{

  foreach($item in $sourceht.keys)
  {
    if(-Not $targetht.ContainsKey($item))
    {
      $item
    }
  }

}

function CompareHashtableValues($sourceht, $targetht)
{

  # Get-TypeData $newht.keys

  # Compare-Object $sourceht $targetht -Property Keys

  foreach($item in $sourceht.keys)
  {
    if($targetht.ContainsKey($item))
    {
      if($sourceht[$item] -ne $targetht[$item])
      {
        $item
      }
    }
  }

}

function HashToFolder($shadowRepo, $htchanged,$htadded,$htremoved)
{
  foreach ($key in $htchanged.keys)
  {
    # Add-Content $shadowRepo$key $htchanged[$key]
    Set-Content $shadowRepo$key $htchanged[$key]
  }

  foreach ($key in $htadded.keys)
  {
    Add-Content $shadowRepo$key $htadded[$key]
  }

  foreach ($key in $htremoved.keys)
  {
    Remove-Item $shadowRepo$key
  }

}

function FilterHash($hashTable, $keys)
{

}
function HashFromFolder($shadowRepo)
{

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

function ChangesInVBE($excelFile, $cached)
{
  $app = GetExcel $excelFile

  $proj = GetProject $app

  $codes = ModulesToHashtable $proj

}


function RepoChanged($dbFile,$ExportLocation,$dteChange)
{

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

  if(!$accessRun)
  {
    #     $modules.add($module.Name, $modDate)
    Write-Information 'file closed' -InformationAction Continue
  } else
  {
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

function export
{
  param(
    [Parameter(Mandatory=$true)]$appPath
    ,$sourceDir
    ,[Nullable[boolean]]$doFullExport
    ,[Nullable[boolean]]$sanitizeQuery
    ,[string]$secondarySourceDir
  )
  process
  {

    #parameter defaults
    if ([string]::IsNullOrEmpty($doFullExport))
    {$doFullExport = $true
    }
    if ([string]::IsNullOrEmpty($sanitizeQuery))
    {$sanitizeQuery = $true
    }
    if ([string]::IsNullOrEmpty($secondarySourceDir))
    {$secondarySourceDir = ""
    }

    #file exits ?
    $appfile = Get-Item $appPath

    $appPath = $appfile.FullName

    if ($null -eq $sourceDir)
    {
      $sourceDir = getSourceDir $appPath
    }

    # if ($null -ne $secondarySourceDir){
    #   $secondarySourceDir = Get-AbsolutePath $secondarySourceDir
    # }

    # $appfile.Name
    Write-Information "app file $appfile" -InformationAction Continue
    Write-Information "app path $appPath" -InformationAction Continue
    Write-Information "worker path $workerPath" -InformationAction Continue
    Write-Information "source dir  $sourceDir" -InformationAction Continue
    Write-Information "full export $doFullExport" -InformationAction Continue
    Write-Information "full export $secondarySourceDir" -InformationAction Continue

    $app = CreateAccess
    $app.Visible = $true

    $app.OpenCurrentDatabase($appPath)

    RemoveReference $app "MSAccessVCS"

    if ($workerPath -eq $appPath)
    {
      Write-Information "exporting worker mode ... " -InformationAction Continue
    } else
    {
      $ref = $app.References.AddFromFile($workerPath)
    }

    #export options
    [boolean]$fullExport = $doFullExport
    [boolean]$printVars = $false
    [string]$localTableWc = "t_*"
    [string]$srcPath = $sourceDir
    [boolean]$buildFromSql = $true
    [boolean]$optSanitizeQuery = $true
    [string]$SecondarySrcPath = $secondarySourceDir

    #export execution
    $app.Run("Export_Cli", [ref]$fullExport, [ref]$printVarS, [ref]$localTableWc, [ref]$srcPath, [ref]$buildFromSql, [ref]$optSanitizeQuery, [ref]$SecondarySrcPath)

    #exportlog
    Get-Content (Join-Path $sourceDir "Export.log")

    RemoveReference $app "MSAccessVCS"

    $app.CloseCurrentDatabase()
    $app.Quit()

  }
}
function build(
  $sourceDir
  ,[Nullable[boolean]]$devWorker
  # ,[string]$secondarySourceDir
)
{

  if ([string]::IsNullOrEmpty($devWorker)) {$devWorker = $false }
  # if ([string]::IsNullOrEmpty($secondarySourceDir)) {$secondarySourceDir = ""}

  [string]$sourceDir = (Get-Item $sourceDir).FullName

  Write-Information "source dir $sourceDir" -InformationAction Continue

  #prepare file names
  $targetName = ReadJsonConfig $sourceDir

  if ($targetName -like "*.accdb")
  {
    $suffix = ".accdb"
  } elseif ($targetName -like "*.accda")
  {
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

  if(Test-Path -Path $buildFile -PathType Leaf) {
    Remove-Item $buildFile
  }

  #start build process
  $app = CreateAccess
  $app.Visible = $true

  $app.NewCurrentDatabase($buildFile)

  #rename worker
  # if($devWorker -eq $true){
  #
  #   $workerName = (Get-Item $workerPath).Name
  #   $workerDir = Split-Path -Path $workerPath
  #
  #   $workerTempPath = Join-Path -Path $workerDir -ChildPath "Temp_$workerName"
  #
  #   if(Test-Path -Path $workerTempPath -PathType Leaf){
  #     Remove-Item $workerTempPath
  #   }
  #
  #   Copy-Item $workerPath -Destination $workerTempPath
  #
  #   $workerFinalPath = $workerTempPath
  #
  # }else{
  $workerFinalPath = $workerPath
  # }

  $ref = $app.References.AddFromFile($workerFinalPath)

  #run build
  if($devWorker -eq $true)
  {
    Build_Cli0 $sourceDir $app
    exit
  } else
  {
    $app.Run("Build_Cli", [ref]$sourceDir)
  }


  #read & print log
  Get-Content (Join-Path $sourceDir "Build.log")

  RemoveReference $app "MSAccessVCS"

  $app.CloseCurrentDatabase()

  $app.Quit()

  #arcrhive file if needed
  if(Test-Path -Path $targetFile -PathType Leaf)
  {

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

  #cleanup if worker path if devworker
  if($devWorker -eq $true)
  {
    if(Test-Path -Path $workerTempPath -PathType Leaf)
    {
      Remove-Item $workerTempPath
    }
  }

}

function Build_Cli0 (
  $sourceDir
  ,$app
)
{

  Write-Information "going over cli0 function" -InformationAction Continue

  $optionsFile = FolderHasVcsOptionsFile $sourceDir
  if ($false -eq $optionsFile)
  {
    Write-Information "vcs-options.json not found in source dir exiting" -InformationAction Continue
    exit
  }

  $options = $app.Run("GetOptions")

  $options.LoadOptionsFromFile([string]$optionsFile)
  $options.ExportFolder = [string]$sourceDir

  $vcsOptions = GetVCSOptions $sourceDir
  # $vcsOptions.Info.AddinVersion

  # $options

  $buildFile = $app.CurrentProject.FullName()

  Write-Information "building at file $buildFile" -InformationAction Continue

  # $options
  Write-Information "removing non builtin references" -InformationAction Continue
  # $app.Run("RemoveNonBuiltInReferences")

  #TODO Run Before Build
  if($null -ne $options.RunBeforeBuild)
  {
    Write-Information "will run before build script" -InformationAction Continue
  }

  $containers = $app.Run("GetAllContainersPS")

  $secondaryFolders = $vcsOptions.Options.SecondaryExportFolders
  # $secondaryFolders

  # Write-Information "initializeForms "$secondaryFolders -InformationAction Continue
  # exit


  $docmd = $app.Docmd
  $proj = GetProject $app

  'modules'
  $files = GetFilesList $sourceDir "modules" @("*.bas","*.cls") $secondaryFolders

  foreach($file in $files) {
    DbModule_Import $file $proj $docmd
  }

  'macros'
  $files = GetFilesList $sourceDir "macros" @("*.bas") $secondaryFolders

  foreach($file in $files) {
    DbMacro_Import $file $app
  }

  'queries'
  $files = GetFilesList $sourceDir "queries" @("*.sql") $secondaryFolders

  foreach($file in $files) {
    DbQuery_Import $file $app
  }

  'project settings'
  $file = GetFileAsList $sourceDir "project.json"
  DbProject_Import $file $app

  'project propert settings'
  $file = GetFileAsList $sourceDir "proj-properties.json"
  DbProject_Import $file $app

  'db properties'
  $file = GetFileAsList $sourceDir "dbs-properties.json"
  DbProperties_Import $file $app

  'vbe properties'
  $file = GetFileAsList $sourceDir "vbe-project.json"
  DbVbeProject_Import $file $app

  #Import Container
  foreach ($container in $containers) {
    $container.SourceFolder

    $files = $container.GetFileList()

      foreach($file in $files) {
        $file
        $container.Import([string]$file)
      }
    # }

  }

  #Initialize Forms
  Write-Information "initializeForms" -InformationAction Continue
  $app.Run("InitializeForms")

  #TODO Run After Build
  if($null -ne $options.RunAfterBuild) {
    Write-Information "will run after build script" -InformationAction Continue
  }

  #Save Index Value
  Write-Information "saving data to version contrindex " -InformationAction Continue
  $vcsIndex = $app.Run("GetVcsIndex")
  $now = get-Date
  $vcsIndex.FullBuildDate = [datetime]$now
  $vcsIndex.Save([string]$sourceDir)
  $app.Run("CloseVcsIndex")
}

function SaveVcsOptions(){
  Write-Information "!!TBD!!" -InformationAction Continue
  #$myJson | ConvertTo-Json -Depth 4 | Out-File .\test.json
}

function DefaultVcsOptions(){
  Write-Information "!!TBD!!" -InformationAction Continue
}

function GetVCSOptions(
  [string]$sourceDir
  ){

  $optionsFile = (Get-ChildItem -Path $sourceDir -File "vcs-options.json").FullName
  $optionsRaw = Get-Content $optionsFile -Raw

  $optionsRaw | ConvertFrom-Json
}

function DbTableDef_Import(
  [string]$fileName
  ,$app
  ,$tdata
  ,$fdata
  ){

  $json = (Get-Content $fileName | ConvertFrom-Json)

  $items=$json.Items

  $db=$app.CurrentDb()

  if($null -eq $items.Connect) {

    $supportName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + ".xml"
    $supportPath = Split-Path -Path $fileName

    # $supportName
    [int]$acStructureOnly=0 #constant

    $app.ImportXML([string](Join-Path $supportPath $supportName), $acStructureOnly)

    $tbl=$db.TableDefs($items.Name)
  }else{

    $tbl=$db.CreateTableDef($items.Name)
    $tbl.Connect=$items.Connect
    $tbl.SourceTableName=$items.SourceTableName
    $tbl.Attributes=$items.Attributes

    try{
      $db.TableDefs.Append($tbl)
    }catch{

    }

    $tbl.RefreshLink()
    ($db.TableDefs).Refresh()

    #TODO Create Unique Index
    #if not access linked table create pseudo index
    if( -not $items.Connect -like ";DATABASE=*"){

      try{
        $pkName = $items.PrimaryKey
      }catch{
        $pkName = ''
      }

      if ( ($pkName -ne '') -and (TableDefs_IndexAvailable $tbl) ) {
        $tblName = $items.SourceTableName
        $sql = "CREATE UNIQUE INDEX __uniqueindex ON [$tblName] ($pkName) WITH PRIMARY"
        $dbFailOnError = 128
        $db.Execute($sql,$dbFailOnError)
        ($db.TableDefs).Refresh()
      }
    }


  }

  $keys=$json.Items.Properties.psobject.properties.Name

  #table properties
  foreach ($key in $keys){

    $value=$items.Properties.$key
    $type=$tdata.$key.Type

    AcProperty_Create $tbl $key $type $value

  }

  #field properties
  foreach ($fld in $json.Items.Fields) {
    # $fld.Name
    # $fldo.Name

    $fldo=$tbl.Fields($fld.Name)
    $fkeys=$fld.Properties.psobject.properties.Name

    foreach($fkey in $fkeys){

      $ftype=$fdata.$fkey.Type
      $fvalue=$fld.Properties.$fkey

      AcProperty_Create $fldo $fkey $ftype $fvalue

    }

  }

}


function DbTableDef_HasUniqueIndex(
    $tdf
    ){

    if(DbTableDef_IndexAvailable $tdf){
      $idxs = $tdf.Indexes
      foreach($idx in $idxs){
        if($idx.Unique -eq $true){
          return $true
        }

      }

    }

}

function DbTableDef_IndexAvailable(
    $tdf
    ){

    $test = 0

    try{
      $test = $tdf.Indexes.Count
    }catch{

    }

    return $test -eq 0
}

function DbRelation_Import(
  [string]$fileName
  ,$app
  ){

  #exit if path not exists
  if (!(Test-Path -Path $fileName)) {
    return
  }

  $db = $app.CurrentDb()

  $json = (Get-Content $filename | ConvertFrom-Json)

  $items = $json.Items

  $relation = $db.CreateRelation($items.Name, $items.Table, $items.ForeignTable)

  $relation.Attributes = $items.Attributes

  foreach($fld in ($items.Fields)){
    $field = $relation.CreateField($fld.Name)
    $field.ForeignName = $fld.ForeignName
    $relation.Fields.Append($field)
  }

  try{
    $db.TableDefs($items.Table).Indexes.Delete($items.Name)
    $db.TableDefs($items.ForeignTable).Indexes.Delete($items.Name)
    $db.Relations.Delete($items.Name)
  }
  catch{

  }

  $db.Relations.Append($relation)

}

function DbQuery_Import(
  [string]$fileName
  ,$app
  ){

  # $tempFileName = [System.IO.Path]::GetTempFileName()
  [string]$objectName = GetObjectNameFormFileName $fileName
  $content = (Get-Content $fileName)

  $db=$app.CurrentDb()
  $qry=$db.CreateQueryDef([string]$objectName, [string]$content)

  # $data=$app.CurrentData
  # $qry=$data.AllQueries([string]$objectName)

  $qry=$db.QueryDefs([string]$objectName)
  $qry.SQL = [string]$content
  # $qry

  $queries=$db.QueryDefs
  $queries.Refresh()

  # $module
  $fileName

  # get support file name
  $supportName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + ".json"

}

function DbMacro_Import(
  [string]$fileName
  ,$app
  ){

  [string]$objectName = GetObjectNameFormFileName $fileName

  $tempFileName = [System.IO.Path]::GetTempFileName()
  $content = (Get-Content $fileName).Split($ls)

  [System.IO.File]::WriteAllLines($tempFileName, $content, $sencd)

  [int]$macroTypeAc=4

  LoadComponentFromText $tempFileName $objectName $macroTypeAc $app

  # $module
  $fileName

}

function LoadComponentFromText(
    [string]$fileName
    ,[string]$objectName
    ,[int]$fileType
    ,$app
    ){

    $app.LoadFromText($fileType,$objectName,$fileName)

}

function DbProject_Import(
  [string]$fileName
  ,$acc
  ){

  $proj = $acc.CurrentProject

  $json = (Get-Content $fileName -Raw) | ConvertFrom-Json

  $proj.RemovePersonalInformation=$json.Items.RemovePersonalInformation

  #TODO Version Control
}

function DbVbeReferences_Import(
  [string]$fileName
  ,$acc
  ){

  $json = (Get-Content $fileName -Raw | ConvertFrom-Json) # -AsHashtable
  $items=$json.Items
  $vbproj = GetProject $acc
  $refs = $vbproj.References

  $existing_refs=@{}

  foreach ($ref in $refs){
     $existing_refs[$ref.GUID]=$ref.Name
  }

  $keys=$json.Items.psobject.properties.Name

  foreach ($key in $keys){

    $guid=$items.$key.GUID
    $version=[string]($items.$key.Version) -split '.'
    $maj=$version[0]
    $min=$version[1]

    #TODO Add ref from filename
    if( -not $existing_refs.ContainsKey($guid)){
      $newref=$refs.AddFromGuid([string]$guid, [int]$maj, [int]$min)
      # Write-Information "adding ref $key" -InformationAction Continue
    }else{
      # Write-Information "exists ref $key" -InformationAction Continue

    }


  }

  #read-only - properties
  #FileName,Mode,Protection,Type

  #TODO Version Control
}


function DbVbeProject_Import(
  [string]$fileName
  ,$acc
  ){

  $json = (Get-Content $fileName -Raw | ConvertFrom-Json) # -AsHashtable
  $items=$json.Items
  $vbproj = GetProject $acc

  try{

    $vbproj.Name=$items.Name
    $vbproj.Description=$items.Name

    $acc.SetOption("Conditional Compilation Arguments",$items.ConditionalCompilationArguments)
  }catch{
    Write-Information "failed to import vbeproject" -InformationAction Continue
  }

  try{

    $vbproj.HelpContextId=$items.HelpContextId
    $vbproj.HelpFile=$items.HelpFile

  }catch{
    Write-Information "failed to import help" -InformationAction Continue
  }

  #read-only - properties
  #FileName,Mode,Protection,Type

  #TODO Version Control
}

function DbListProperties(
  $objs
  ){

  $exp=@{}

  foreach( $obj in $objs ){
    # 'loop'
    # $obj
    $props=$obj.Properties
    # $props
    foreach($prop in $props){
      # $prop.Name
      if ( -not $exp.ContainsKey($prop.Name)){
        $exp[$prop.Name]+=@{'Type'=$prop.Type}
      }
    }
  }

  $exp

}

function DbListPropertiesLevelBellow(
  $objs
  ,[string]$l2name
  ){

  $exp=@{}

  foreach( $obj0 in $objs ){
    # 'loop'
    # $obj
    $l2items=$obj0.$l2name
    foreach($obj in $l2items){
      $props=$obj.Properties
      foreach($prop in $props){
        if ( -not $exp.ContainsKey($prop.Name)){
          $exp[$prop.Name]+=@{'Type'=$prop.Type}
        }
      }
    }
  }

  $exp

}

#custom function for merging two object on property level
function Merge-Object ($target, $source) {
    $source.psobject.Properties | % {
        if ($_.TypeNameOfValue -eq 'System.Management.Automation.PSCustomObject' -and $target."$($_.Name)" ) {
            Merge-Object $target."$($_.Name)" $_.Value
        }
        else {
            $target | Add-Member -MemberType $_.MemberType -Name $_.Name -Value $_.Value -Force
        }
    }
}

function DbProjectProperties_Import(
  [string]$fileName
  ,$acc
  ){

  $proj=$acc.CurrentProject
  $props=$proj.Properties
  $existing_props=@{}

  foreach ($prop in $props){

    if ($prop.Name -ne 'Connection') {
     $existing_props[$prop.Name]=$prop.Value
    }
  }

  # $json = (Get-Content $fileName -Raw) | ConvertFrom-Json

  foreach ($jprops in $json.Items){

    $key=$jprops.psobject.properties.Name

    if( ($key -ne "Name") -and ($key -ne "Connection") ){

      $value=$jprops.psobject.properties.Value

      #check if property exists
      if($existing_props.ContainsKey($key)){
        $props.item($key).Value=$value
      }else{
        $props.Add($key,$value)
      }

    }

  }

  #TODO Version Control
}

function DbProperties_Import(
  [string]$fileName
  ,$acc
  ){

  $db=$acc.CurrentDb()

  $json = (Get-Content $fileName -Raw | ConvertFrom-Json) # -AsHashtable

  $items=$json.Items
  $keys=$json.Items.psobject.properties.Name

  $exclude=@("Version","Connection","Name","CollatingOrder","Updatable","Transactions","RecordsAffected","DesignMasterID","ReplicaID","ShowDocumentTabs")

  # exit
  foreach ($key in $keys){

    if( -not $exclude.Contains($key) ){

      $value=$items.$key.Value
      $type=$items.$key.Type

      AcProperty_Create $db $key $type $value

    }

  }

  #TODO Version Control
}
function AcProperty_Create(
    $obj
    ,$key
    ,$type
    ,$value
    ){

  $props=$obj.Properties

    # switch($type){
    #   1 {$value=[boolean]$value;break}
    #   2 {$value=[string]$value;break}
    #   3 {$value=[string]$value;break}
    #   4 {$value=[string]$value;break}
    #   10 {$value=[string]$value;break}
    #   12 {$value=[string]$value;break}
    #   15 {$value=[string]$value;break}
    # }

  $exists=(AcProperty_Check $obj $key $type $value)
  # Write-Information "property set $key check: $exists" -InformationAction Continue

    switch($exists){
      0 {
        $prop=$obj.CreateProperty([string]$key, [int]$type, [string]$value)
          $props.Append($prop)
          # Write-Information "creating $key" -InformationAction Continue
      }
      1 {
        $props.Item([string]$key).Value=[string]$value
        # Write-Information "creating $key" -InformationAction Continue
      }
    }



# #check if property exists
# if($true -eq $exists){
#   $props.item([string]$key).Value=$value
# }else{
#   $xprop=$db.CreateProperty([string]$key, [int]$type, $value)
#   $props.Append($xprop)
# }
}

# Formats JSON in a nicer format than the built-in ConvertTo-Json does.
function Format-Json([Parameter(Mandatory, ValueFromPipeline)][String] $json) {
  $indent = 0;
  ($json -Split '\n' |
    % {
      if ($_ -match '[\}\]]') {
        # This line contains  ] or }, decrement the indentation level
        $indent--
      }
      $line = (' ' * $indent * 2) + $_.TrimStart().Replace(':  ', ': ')
      if ($_ -match '[\{\[]') {
        # This line contains [ or {, increment the indentation level
        $indent++
      }
      $line
  }) -Join "`n"
}

function AcProperty_Check(
    $obj
    ,$key
    ,$type
    ,$value
    ){

  $props=$obj.Properties

  try{
    $propType=$props.item($key).Type
  }catch{
    $propType=-1
  }

  #not exists
  if ($propType -eq -1){
    return 0
  }

  #exist but wrong type
  if ($propType -ne $type){
    $props.Delete($key)
    return 0
  }

  #nothing to do
  if ($props.item($key).Value -eq $value){
    return 2 #nothing
  }else{
    return 1 #change
  }

}

function DbModule_Import(
  [string]$fileName
  ,$proj
  ,$docmd
  ){

  [string]$objectName = GetObjectNameFormFileName $fileName
  $content = DbModule_ParseSourceFile $fileName $objectName

  $tempFileName = [System.IO.Path]::GetTempFileName()

  [System.IO.File]::WriteAllLines($tempFileName, $content, $sencd)

  $module = ImportCode $proj $tempFileName $objectName
  [int]$moduleTypeAc = 5

  $docmd.Save($moduleTypeAc, $objectName)

  # $module
  $fileName

}

# function DbModule_Save(
#     $app
#    , $moduleName
#     ){
#
#
#
# }


function DbModule_ParseSourceFile(
    [string]$fileName
    ,[string]$objectName
    ){

    $content = (Get-Content $fileName).Split($ls)
    [bool]$haveHeader = $false
    [bool]$isClass = $false

    # $fileName

    #loop throug header
    foreach ($ln in (0..8)){
      if($content[$ln].StartsWith("VERSION 1.0 CLASS")){
        $haveHeader = $true
        $isClass = $true
        break
      }

      if($content[$ln].StartsWith("Attribute VB_Name = """)){
        $haveHeader = $true
        break
      }

      if($content[$ln].StartsWith("Attribute VB_GlobalNameSpace = ")){
        $isClass = $true
        break
      }

    }

    $contentF = @()

    if ($true -ne $haveHeader){
      if ($true -eq $isClass){
        $contentF += "VERSION 1.0 CLASS"
        $contentF += "BEGIN"
        $contentF += "  MultiUse = -1  'True"
        $contentF += "END"
      }
      $contentF += "Attribute VB_Name = ""$objectName"""
    }

    # $content

    #module body
    foreach ($line in $content){
      $contentF += $line
    }

    #remove trailing space
    # $contentF = $contentF.Remove(2)

    #return final to string
    $contentF -join $ls

}

function GetObjectNameFormFileName(
    [string]$fileName
    ){

    $objectName = Split-Path -Path $fileName -Leaf
    $objectName =[System.IO.Path]::GetFileNameWithoutExtension($objectName)

    $fileCodes = @(
       "%3C"
      ,"%3E"
      ,"%3A"
      ,"%22"
      ,"%2F"
      ,"%5C"
      ,"%7C"
      ,"%3F"
      ,"%2A"
      ,"%25"
      )

    $forbiddenChars = @(
       "<"
      ,">"
      ,":"
      ,""""
      ,"/"
      ,"\"
      ,"|"
      ,"?"
      ,"*"
      ,"%"
      )

    foreach($x in (0..9)){
      $objectName = $objectName.Replace($fileCodes[$x],$forbiddenChars[$x])
    }

    [string]$objectName

}

function GetFileAsList(
  [string]$sourceDir
  ,$fileName
  ){

  $files = @()

  $optionsFile = (Get-ChildItem -Path $sourceDir -File $fileName).FullName

  $files += $optionsFile

  $files

}

function GetFilesList(
  [string]$sourceDir
  ,$module
  ,$extensions
  ,$secondaryFolders
  ){

  # $module = "modules"
  # $extensions = @("*.bas","*.cls")

  $files = @()
  $localPath = (Join-Path $sourceDir $module)

  foreach($ext in $extensions){
    $files += GetFilePathsInFolder $localPath $ext
  }

  #check if there are any secondary source folders
  if($PSBoundParameters.ContainsKey('secondaryFolders')){
    $files += GetFilePathsInSecondaryFolder $sourceDir $secondaryFolders $module $extensions
  }

  $files

}

function GetFilePathsInSecondaryFolder(
  [string]$sourceDir
  ,$secordarySourceDirs
  ,$sourceFolder
  ,$extensions
){

  if($secordarySourceDirs.length -gt 0){

    $files = @()

    #loop all paths
    foreach($relativePath in $secordarySourceDirs){
      $sourcePath = $sourceDir + $ds + $relativePath + $ds + $sourceFolder

      if(Test-Path -Path $sourcePath){
        foreach($ext in $extensions) {
          $files += GetFilePathsInFolder $sourcePath $ext
        }
      }
    }

    #return final array
    $files
  }
}


function GetFilePathsInFolder(
  [string]$folder
  ,[string]$filePattern
){

  if ([string]::IsNullOrEmpty($filePattern)) {$filePattern = "*.*" }

  Get-ChildItem -Path $folder -Filter $filePattern -Recurse | Select-Object -ExpandProperty FullName
}

function FolderHasVcsOptionsFile (
  $sourceDir
){

  $optionsFile = Join-Path $sourceDir "vcs-options.json"

  if(Test-Path -Path $optionsFile -PathType Leaf){
    $optionsFile
  } else {
    $false
  }

}


function RemoveReference(
  $app
  ,[string]$refName
){

  $refs = $app.References

  foreach ($ref in $refs){

    if ($ref.Name -eq $refName){
      Write-Information "refrence removed $refName" -InformationAction Continue
      $refs.Remove($ref)
    }

  }

}

function RenameWorker(
  $appPath
)
{
  if ($workerPath -eq $appPath)
  {

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
  } else
  {
    $workerFinalPath = $workerPath
  }
}

function ReadJsonConfig($sourceFile)
{

  $projectPath = Join-Path $sourceFile "vbe-project.json"

  $json = (Get-Content "$projectPath" -Raw) | ConvertFrom-Json | Select-Object Items

  return $json.Items.FileName

}
Function Get-AbsolutePath
{
  param([string]$Path)
  [System.IO.Path]::GetFullPath([System.IO.Path]::Combine((Get-Location).ProviderPath, $Path));
}

function getSourceDir(
  $appPath
)
{

  #get full path from relative path
  $appDir = Split-Path -Path $appPath

  #build
  if ((Split-Path -Path $appDir -Leaf) -eq 'build')
  {
    $projDir = Split-Path -Path $appDir
  }

  #root
  if($null -eq $projDir)
  {
    $projDir = $appDir
  }

  # Write-Verbose "$projDir"
  # Write-Information "refrence removed $projDir" -InformationAction Continue

  #find project files within project folder
  #NTH if multiple items check project by json options
  $sourceProjFile = (Get-ChildItem -Path $projDir -Recurse -File -Filter "vbe-project.json")[0].FullName

  #if there is not any project file yet use app dir as project dir
  if($null -eq $sourceProjFile)
  {
    $sourceDir = Join-Path $projDir "src"
  } else
  {
    $sourceDir = (Split-Path -Path $sourceProjFile)
  }

  # Write-Information "get project dir $sourceDir" -InformationAction Continue
  return $sourceDir

}
