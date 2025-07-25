#Hello, This Script will not function in it's current form please read the discription to update the script acordingly or contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281

function Folder-browser{
Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowser.Description = 'Select the folder containing the data'
$FolderBrowser.RootFolder = 'MyComputer'
$result = $FolderBrowser.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
if ($result -eq [Windows.Forms.DialogResult]::OK){
    return $FolderBrowser.SelectedPath
}
else {
    Write-warning "Please select the parent folder for the Script to search"
    Start-Sleep -Seconds 5
    exit
}
}
if(!(Get-InstalledModule -Name powershell-yaml))
    {
    Write-host "Installing package powershell-yaml This is one time process" -ForegroundColor Red -BackgroundColor Cyan
    Expand-Archive -Path "$PWD\powershell-yaml.0.4.3.zip"
    New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\powershell-yaml" -ItemType Directory
    Move-Item -Path "$PWD\powershell-yaml.0.4.3" -Destination "$PWD\0.4.3"
    Move-Item -Path ".\0.4.3" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\powershell-yaml"
    }
Import-Module powershell-yaml
Write-host "Hi, Welcome to YAML data Scraper Tool `nPlease select the Root Folder which contains YML Files" -ForegroundColor Green
$parent=Folder-browser
$yaml_files=Get-ChildItem -Path "$parent\*.yml" -Recurse

$dat=get-date -UFormat "%d-%b-%H.%M"
if(Test-Path "$parent\File1.txt" ){Rename-Item -Path "$parent\File1.txt" -NewName "File1_$dat.txt"}
if(Test-Path "$parent\File2.txt" ){Rename-Item -Path "$parent\File2.txt" -NewName "File2_$dat.txt"}
if(Test-Path "$parent\File3.txt" ){Rename-Item -Path "$parent\File3.txt" -NewName "File3_$dat.txt"}
if(Test-Path "$parent\File4.txt" ){Rename-Item -Path "$parent\File4.txt" -NewName "File4_$dat.txt"}

foreach ($file in $yaml_files)
    {
        try{$map = ConvertFrom-Yaml -Yaml (Get-Content -Path $file.fullname -Raw)}
        catch{Write-warning "$($file.name) is not a Valid YAML File"
        continue}
        Write-host "Parsing file $($file.name)"
        $FileLocName=($file.name).Split('-') | select -Index 0
        $FileObjName=($file.name).Split('-') | select -Index 1
        $FileObjName=$FileObjName -replace '\.yml',''
        $ObjectName=$map.steps.Keys
        $operation=$map.steps."$ObjectName".operation


        "$FileLocName~$FileObjName~$($ObjectName.split('-')[0])~$($ObjectName.split('-')[1])~~~$($operation.config.postSQL)~$($operation.config.preSQL)~$($operation.config.testsEnabled)~$($operation.database)~$($operation.dependencies)~$($operation.deployEnabled)~$($operation.description)~$($operation.isDataVault)~$($operation.ismultisource)~$($operation.locationID)~$($operation.locationNAME)~$($operation.materializationType)~" | out-file "$parent\File1.txt" -Append
        
        $metadata=$map.steps."$ObjectName".operation.metadata
        $Columns=$map.steps."$ObjectName".operation.metadata.columns
        foreach ($column in $columns)
        {
            if($Column.config -eq $null){
            "$($column.columnReference.columnCounter)~$($column.columnReference.stepCounter)~$($column.dataType)~$($column.description)~$($column.hashColumns)~$($column.hashDetails)~$($column.name)~$($column.nullable)~~~$($column.sourceColumnReferences.columnReferences.columnCounter)~$($column.sourceColumnReferences.columnReferences.stepCounter)~$($column.sourceColumnReferences.transform)" | out-file "$parent\File2.txt" -Append
            }
        }

        foreach ($column in $columns)
        {
            if($column.isSystemcreateDate -eq $null){$isdate=$column.isSystemUpdateDate}
            if($column.isSystemUpdateDate -eq $null){$isdate=$column.isSystemCreateDate}
            if($Column.config -ne $null){
            "$($column.acceptedValues.strictMatch)~$($column.acceptedValues.'values')~$($column.appliedColumnTests.values)~~$($column.columnReference.columnCounter)~$($column.columnReference.stepCounter)~$($column.config.values)~$($column.datatype)~$($column.defaultValue)~$($column.description)~$($column.hashColumns)~$isdate~$($column.name)~$($column.nullable)~~$($column.sourceColumnReferences.columnReferences)~$($column.sourceColumnReferences.transform)" | out-file "$parent\File3.txt" -Append
            }
        }
        
        "$FileLocName~$FileObjName~$($metadata.cteString)~$($metadata.enabledColumnTestIDs)~~~$($metadata.sourceMapping.customsql.customSQL)~~$($metadata.sourceMapping.dependencies.locationName)~$($metadata.sourceMapping.dependencies.nodeName)~~$($metadata.sourceMapping.join.joinCondition)~$($metadata.sourceMapping.name)~$($metadata.sourceMapping.noLinkRefs)~$($operation.name)~$($operation.overrideSQL)~$($operation.schema)~$($operation.sqlType)~$($operation.Type)~$($map.steps."$ObjectName".stepcounter)" | out-file "$parent\File4.txt" -Append
        
    }

Write-Host "Execution Completed" -ForegroundColor Gray
pause
