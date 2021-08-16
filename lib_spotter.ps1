# Write-Information 'lib loaded' -InformationAction Continue

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

