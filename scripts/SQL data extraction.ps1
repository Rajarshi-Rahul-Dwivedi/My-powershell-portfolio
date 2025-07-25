#Hello, This Script will not function in it's current form please read the discription to update the script acordingly or contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281
#reserved []().\^$|?*+{}
function read-sql
{
    [CmdletBinding()]
    param 
    (
    [Object[]]$file_content,
    [string]$foldername,
    [string]$filename
    )
    $Table=[pscustomobject]@()
    $count=0
    $skip=$true

    foreach($line in $file_content)
    {
           if($skip -eq $true){
            if($line -match "create or replace table\s(.+)\s\(")
            {
                $table_name=$matches[1]
                $skip=$false
                $Table+=[pscustomobject]@{'FolderName'=$foldername;'FileName'=$filename;'TableName'=$table_name;'ColumName'=$null;'DataType'=$null;'Nullability'=$null;'OtherProperties'=$null}
                continue
            }
           }
            
            if($skip -eq $false)
        {
            if($line -match ";"){$skip=$true}
            $matches=$null
            $column_name=$null
            $Data_Type=$null
            $nullability=$null
            $other_properties=$null

            $line -match "^\s*(\S+)\s+(\S+),$" | Out-Null
            if($matches -ne $null){
            $column_name=$matches[1]
            $Data_Type=$matches[2]
            $nullability=$null
            $other_properties=$null
            }
         

            if($matches -eq $null){
            $line -match "^\s*(\S+)\s+(\S+)\s+(.+null)\s?(.*),$" | Out-Null
            if($matches -ne $null){
            $column_name=$matches[1]
            $Data_Type=$matches[2]
            $nullability=$matches[3]
            $other_properties=$matches[4]
            }
            }


            if($matches -eq $null){
            $line -match "^\s*(\S+)\s+(\S+)\s+\s(.+),$" | Out-Null
            if($matches -ne $null){
            $column_name=$matches[1]
            $Data_Type=$matches[2]
            $nullability=$null
            $other_properties=$matches[3]
            }
            }
            
            if($matches -eq $null){
            $line -match "constraint (\S+)\s+(.*)\s+\((.+)\)" | Out-Null
            if($matches -ne $null){
            $column_name="constraint"
            $Data_Type=$matches[1]
            $nullability=$matches[2].trim()
            $other_properties=$matches[3]
            }
            }

            if($matches -eq $null){
            $line -match "^\s*(\S+)\s+(\S+)\s+(\S+)\s?(.*),$" | Out-Null
            if($matches -ne $null){
            $column_name=$matches[1]
            $Data_Type=$matches[2]
            $nullability=$matches[3]
            $other_properties=$matches[4]
            }
            }

            if($column_name -eq $null){$column_name="complex"}
            if($column_name -eq 'complex' -and $line -match ";"){$column_name=$null}
            $Table+=[pscustomobject]@{'FolderName'=$foldername;'FileName'=$filename;'TableName'=$table_name;'ColumName'=$column_name;'DataType'=$Data_Type;'Nullability'=$nullability;'OtherProperties'=$other_properties}
        }
    }
    return $table
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

$parent=Get-Content "$pwd/File_location.txt"
$folder_01=Get-ChildItem "$parent\01"
$folder_03=Get-ChildItem "$parent\03"
$table_original=@()
foreach($file in ($folder_01+$folder_03))
{
    $file_content=$file.fullname | foreach {(get-content $_) -replace "--.*$",""}
    $file_content= $file_content |  ? {$_.trim() -ne "" }
    $foldername=$file.directory | Select-String "0[13]$" | select -ExpandProperty matches
    $table_original+=read-sql -file_content $file_content  -foldername $foldername.value -filename $file.name
}

$output=$table_original | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
$output | Out-File "$parent\Create_query_details.csv"

import-module powershell-yaml
$output_yaml=ConvertTo-Yaml $table_original
$output_yaml | out-file "$parent\Create_query_details.yaml"
Write-Host "Script Execution Completed `n`n" -ForegroundColor green
Write-Host "Please check output file in $parent" -ForegroundColor red
Start-Sleep -Seconds 14
# ? is alias for Where-Object
#evaluating content of $folder_01


