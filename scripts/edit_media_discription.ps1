#Hello, This Script will not function in it's current form please read the discription to update the script acordingly or contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281
Write-Host "Hi, Please select Parent folder of your Movie Library. Prompt might be in the background of this window" -ForegroundColor green
Function Get-Folder($initialDirectory) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.RootFolder = 'MyComputer'
    if ($initialDirectory) { $FolderBrowserDialog.SelectedPath = $initialDirectory }
    [void] $FolderBrowserDialog.ShowDialog()
    return $FolderBrowserDialog.SelectedPath
}
#Unblock-File -Path .\Start-ActivityTracker.ps1
#Set-ExecutionPolicy Unrestricted
$parent=Get-Folder
$table=@()
echo $PWD.Path
if(!(test-path -Path "$($PWD.Path)\TagLibSharp.dll"))
{
Write-Warning "TagLibSharp.dll Not found please make sure dll file is present in same location as this Script"
pause
Exit
}
[System.Reflection.Assembly]::LoadFrom((Resolve-Path "TagLibSharp.dll"))

#searching for media files

$files1=@(Get-ChildItem -Path "$parent\*.mp4" -Recurse)
$files2=@(Get-ChildItem -Path "$parent\*.mkv" -Recurse)
$files3=@(Get-ChildItem -Path "$parent\*.avi" -Recurse)
$files4=@(Get-ChildItem -Path "$parent\*.mpeg" -Recurse)
$files5=@(Get-ChildItem -Path "$parent\*.WEBM" -Recurse)
$files6=@(Get-ChildItem -Path "$parent\*.flv" -Recurse)
$files7=@(Get-ChildItem -Path "$parent\*.mov" -Recurse)
$files=$files1+$files2+$files3+$files4+$files5+$files6+$files7

#iterating through the found media files
foreach($movie in $files)
 {
    if($movie.fullname -eq $null){continue}
    $video = $video = [TagLib.File]::Create($movie.fullname)
    $previous_title=$video.Tag.Title
    $name=$movie.name
    try{$name=[regex]::Match($movie.name,"(.*)(\(.*\))([.mp4]*[.mkv]*[.avi]*)").captures.groups[1].value} catch{$name=[regex]::Match($movie.name,"(.*)(\.\w+)").captures.groups[1].value}
    $name=$name.Trim(" $")
    if(!($video.Tag.Title -eq $name))
        {
            $video.Tag.Title = $name
            $video.Save()
            $Table+=[pscustomobject]@{'Media Name'=$movie.name;'File Location' =$movie.fullname;'Old Discription title'=$previous_title;'Updated Discription Title'=$name;}
        }
    echo "Checked file $($movie.name) `n"
    $video=$null
 }
#Creating Log File
if(!(Test-Path "$($PWD.Path)\Logs")){mkdir "$($PWD.Path)\Logs"}
$dat=get-date -UFormat "%M-%d-%b"
$table
$output=$table | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
$output | Out-File "$($PWD.Path)\Logs\Updates_$dat.csv"
pause

