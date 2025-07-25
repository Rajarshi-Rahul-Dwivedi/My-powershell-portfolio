#Hello, This Script will not function in it's current form please read the discription or feel free to contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281

start-service -ServiceName  "OracleServiceSampleDB"
start-service -ServiceName  "OracleTNSListenerV1" -ErrorAction SilentlyContinue
start-service -ServiceName  "OracleTNSListenerV2" -ErrorAction SilentlyContinue

function Training_user {
    [CmdletBinding()]
    param (
        $str1
    )

    $tr_user = Get-Content "$results\train_user.txt"
    if (($tr_user -match "no rows selected") -or ($tr_user -eq $null) -or ([String]::IsNullOrWhiteSpace($tr_user))) {
        return "No"                    
        } else {
            return "Yes"
        }
    }


function C20-SYSAUXValidation {
[CmdletBinding()]
    param (
        $str1
    )    
    $retentionContent = Get-Content "$results\retention.txt" | Where-Object {$_ -match "AUTO_STATS_ADVISOR_TASK"}    
    $retentionValue = ($retentionContent -split "\s{2,}" | Select-Object -Last 1).Trim()
    
    $indexStatusContent = Get-Content "$results\index_status.txt" | Where-Object {$_ -match "WRI"}
    $indexStatus = $indexStatusContent | ForEach-Object {
        ($_ -split "\s{2,}" | Select-Object -Last 1).Trim()
    }

    # Check if retention parameter is 10 and all indexes are valid
    if ($retentionValue -eq "10" -and $indexStatus -contains "VALID") {
        return "Valid"
    } else {
        return "Missing"
    }
}


function acl{
[CmdletBinding()]
    param (
        $str1
    )
$acl=Get-Content "$results\acl.txt"
$grantedCount = ($acl -split "`n" | Where-Object { $_.Trim() -eq "Granted" }).Count
if ($grantedCount -eq 3) {
    "valid"
} else {
    "missing"
}
}	
	
function OPI-check{
[CmdletBinding()]
    param (
        $str1
    )
    $is_opi=Get-Content "$results\OPI_check.txt"
    if($is_opi -match "OPI"){
        Write-Host "This property uses OPI (Oracle Payment Interface) `nPlease enter the" -ForegroundColor Cyan -NoNewline
        Write-Host " OPI version " -ForegroundColor Yellow -NoNewline;Write-Host "below, check OPI server for version." -ForegroundColor Cyan -NoNewline
        $opi_version=Read-Host "`n"
        Write-Host "Please enter the" -ForegroundColor Cyan -NoNewline
        Write-Host " MySql version " -ForegroundColor Yellow -NoNewline;Write-Host "below, check OPI server for version." -ForegroundColor Cyan -NoNewline
        $sql_version=Read-Host "`n"
        }else{$opi_version=$sql_version="Not Applicable"}
   return $opi_version,$sql_version
}

function Validate-JavaBackup {
    [CmdletBinding()]
    param (
        $str1
    )

$path1 = "C:\\Program_Files_X86"
$path2 = "C:\\Program_Files"
$javaPathX86 = "C:\\Program_Files_X86\Java"
$javaPath = "C:\\Program_Files\Java"
$jdkPath = "D:\\Sample\\JavaBackup"

$javabackup = Get-ChildItem -Path $path1, $path2 -Directory | Where-Object { $_.Name -match "^JAVA" -and $_.Name -ne "Java" } | Select-Object -ExpandProperty FullName

$javaSubFoldersX86 = Get-ChildItem -Path $javaPathX86 -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "conf" }
$invalidJavaSubFoldersX86 = if ($javaSubFoldersX86.Count -gt 1) { $javaPathX86 }

$javaSubFolders = Get-ChildItem -Path $javaPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "conf" }
$invalidJavaSubFolders = if ($javaSubFolders.Count -gt 0) { $javaPath }

$invalidJDKPreUpgrade = Get-ChildItem -Path $jdkPath -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName

$nonCompliantFolders = @($javabackup + $invalidJavaSubFoldersX86 + $invalidJavaSubFolders + $invalidJDKPreUpgrade) | Where-Object { $_ }

if ($nonCompliantFolders) {
    return "Missing"
} else {
    return "Valid"
}
}

function Validate-httpd {
    [CmdletBinding()]
    param (
        $str1
    )
    
    $paths = @(
        "D:\ora\mwfr\12cappr2\ohs\bin",
        "D:\ora\mwfr\12cappr2\.patch_storage"
    )

    $fileName = "httpd.exe"

    foreach ($path in $paths) {
        if (Test-Path -Path $path) {
            $file = Get-ChildItem -Path $path -Recurse -Filter $fileName -ErrorAction SilentlyContinue
            if ($file) {
                return "Missing"
            }
        }
    }
    
    return "Valid"
}

function Validate-log4j{
[CmdletBinding()]
    param (
        $str1
    )
    if((Test-Path -Path "D:\Oracle\1970\suptools\tfa\release\tfa_home\jlib\log4j-core-2.9.1.jar") -or (Test-Path -Path "D:\Oracle\1970\suptools\tfa\release\tfa_home\jlib\tfa.war"))
    {return "MISSING"} 
    else{return "Valid"} 
}

function log4j-evidence {
    [CmdletBinding()]
    param (
        [string]$Weblogic_server
    )

    if ($Weblogic_server -match "12.2.1.4.0") {
        if ((Get-Content -Path "$scripts\log4j_evidence.txt") -match (Get-Content -Path "$results\host_marsha.txt")) {
            return "VALID"
        } else {
            if (Get-ChildItem -Path "D:\log4j_evidence\*.pdf" -ErrorAction SilentlyContinue) {
                return "Valid"
            } else {
                return "Missing"
            }
        }
    } else {
        return "Not Applicable"
    }
}


function Check-JDK_version{
[CmdletBinding()]
    param (
        $str1
    )
    $valid_count=0
    cmd /c "set JAVA_HOME=D:\ORA\JDK"
    cmd /c "cd /D D:\ORA\JDK\jre\bin"
    $JDK_version= & cmd /c "java -version 2>&1"
    if($JDK_version -match "1.8.0_441"){$valid_count+=1}
    #for JDK OHS
	$wl_service=Get-Service -name 'Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12214ohs_wlserver)' -ErrorAction SilentlyContinue
   if($wl_service.DisplayName -ne $null){$oracle_H="set ORACLE_HOME=D:\ORA\12214ohs"}else{$oracle_H="set ORACLE_HOME=D:\ORA\12213ohs"}
    cmd /c $oracle_H
    cmd /c "set JAVA_HOME=D:\ORA\JDK\oracle_common\jdk\jre"
    cmd /c "cd /d %JAVA_HOME%"
    $JDK_version_ohs= & cmd /c "java -version 2>&1"
    if($JDK_version_ohs -match "1.8.0_441"){$valid_count+=1}

    #for MWFR
    cmd /c "set ORACLE_HOME=D:\ORA\MWFR\12cappr2"
    cmd /c "set JDK_HOME=D:\ORA\JDK"
    cmd /c "set JAVA_HOME=D:\ORA\MWFR\12cappr2\oracle_common\jdk"
    cmd /c "cd /d %JAVA_HOME%"
    $JDK_version_mwfr= & cmd /c "java -version 2>&1"
    if($JDK_version_mwfr -match "1.8.0_441"){$valid_count+=1}

    if($valid_count -eq 3){return "VALID"}else{return "MISSING"} 
}


function Validate-C14{
[CmdletBinding()]
    param (
        $str1
    )
   $Check_flag=$true
   
    foreach($line in Get-Content "$results\C14_Business_Block_notes.txt")
    {
     if($line -match "^\w+\s+(?:Y N|N N)\s*")
      {
       if($line -match "FACTS")
        {
            if(!($line -match "FACTS\s+N N\S*")){$Check_flag=$false}
        }else{if(!($line -match "^\w+\s+(?:Y N|N)\s*$")){$Check_flag=$false}}         
       }
    }
    if($Check_flag -eq $true){return "VALID"}else{return "MISSING"}
}

function Validate-C12{
[CmdletBinding()]
    param (
        $str1
    )
    if((Get-Content "$results\C12_RitzCarltonRewards.txt") -match "RitzCarltonRewards\s*Y\s+N\s*"){return "VALID"}else{Return "MISSING"}
    #(Get-Content "D:\test.txt") | % {$_ -match "RitzCarltonRewards\s*Y\s+N\s*"}
}

function Validate-C19{
[CmdletBinding()]
    param (
        $str1
    )
	if((Get-Content -Path "$scripts\MGP_list.txt") -match $hostname){return "VALID"}
	
	if((Test-Path -Path "d:\Discovery_exports\output.log") -and (Test-Path -Path "D:\DISCOVERY_EXPORTS\Discovery_Exports_$hostname.zip"))
		{
        if((Get-Content "d:\Discovery_exports\output.log") -match "Error"){return "MISSING"} 
        if((Get-Content "d:\Discovery_exports\output.log") -match "Consulting Discovery Scripts"){return "VALID"}else{return "MISSING"}
		}else{return "MISSING"} 
}
function validate-c18_tablespace{
[CmdletBinding()]
    param (
        $str1
    )
    $tablespace=get-content "$results\C18_Table_Space_Verification.txt"
    $temseg_verification=get-content "$results\C18_Tempseg.txt"
    $check_list="OPERA_DAT", "OPERA_INDX", "OXI_DAT", "OXI_INDX", "FINDATA", "FININDX", "LOGDATA", "LOGINDX", "SYSAUX"
    $valid="VALID"
    $return_output=@()
    $return_output+="Tablespace free space status, "
    foreach($line in $tablespace)
    {
        foreach($tablespace_name in $check_list)
            {
                if($line -match $tablespace_name)
                {
                    $percentage_free=$line | Select-String -pattern "(\d*\.\d+|\d\d)\s*$" | select -ExpandProperty matches
                    $percentage_free=[decimal]($percentage_free.Groups[1].value)
                    $return_output+="$tablespace_name : $percentage_free%, "
                    if($percentage_free -lt 20)
                    {
                        $return_output+="Low space in $tablespace_name only $percentage_free % remaining, "
                        $valid="MISSING"
                    }

                }
            }

    }
    foreach($line in $temseg_verification)
    {
        if($line -match "TEMPSEG")
        {
            $free_space=$line | Select-String -pattern "\s(\d+)\s*$" | select -ExpandProperty matches
            $free_space=[decimal]($free_space.Groups[1].value)
            if($free_space -lt 2023)
                    {
                        $return_output+="Tempseg only $free_space MB of free space, "
                        $valid="MISSING"
                    }
        }
    }
    return $valid,$return_output 
}
function validate-c16{
[CmdletBinding()]
    param (
    [string]$property)
    if($c16_value -eq $null -or $c16_value -match "no rows selected" -or $c16_value -match 'null'){return "VALID"}
    if($c16_value -match "$property\s*\t*CRISIS_EXPORT.*\s+\t*Y" -or $c16_value -match "$property\s+\t*CRISIS_EXPORT_DAY.*\s+\t*Y")
    {
        if($c16_value -match "$property\s*\t*CRISIS_EXPORT.*\s+\t*N" -or $c16_value -match "$property\s*\t*CRISIS_EXPORT_DAY.*\s+\t*N"){return "MISSING"}

            if(!($c16_value -match "RMICRISISMGT_")){return "VALID"}else{return "MISSING"}

    }else{return "MISSING"}
}
function Btr_check{
[CmdletBinding()]
    param (
    [string]$property)
    $US_managed=Get-Content "$scripts\US_MANAGED.txt"
    if($US_managed -match $property){
        $value=Get-Content "$results\BTR_CHECK.txt" | select -Index 0
        $counter=Get-Content "$results\BTR_CHECK.txt" | select -Index 1
        try{$counter=[int]$counter}catch{$btr_validation="MISSING"}
        if($counter -ge 1){$btr_validation="VALID"}
            else{$btr_validation="MISSING"}
    }else{$value="Not Applicable"
          $btr_validation="Not Applicable"  
            }
    if($value -eq $null){$value="NULL"}
    return $value,$btr_validation       
}
function HTML_Multi_validation{
[CmdletBinding()]
    param (
        $str1
    )
    if($P_count -gt 1){$marsha_first=" : $(Get-Content -Path "$results\marsha.txt" | select -Index 0)"}
    $Block_rolling=Get-Content -Path "$results\Block rolling date.txt" | select -index 0 
                if($Block_rolling -match "715"){$Block_rolling_val="VALID"}
                else{$Block_rolling_val="MISSING$marsha_first"}
    $mimpg_xml=Get-Content -Path "$results\mimpg_xml_version.txt" | select -Index 0
                if($mimpg_xml -match "V3"){$mimpg_xml_val="VALID"}
                else{$mimpg_xml_val="MISSING$marsha_first"}
    $Oxi_comm=Get-Content -Path "$results\OXI Comm Method_ERS Rate Opera.txt" | select -index 0
                if($Oxi_comm -match 'https://marshaprod.marriott.com:5617/reservationgpms'){$Oxi_comm_val="VALID"}
                else{$Oxi_comm_val="MISSING$marsha_first"}
    $Rate_Codes_val_output=check-Rate_Codes -P_count 0
    $c16_val=validate-c16 $Marsha_excel
               

 if($P_count -gt 1){
    for($i=0;$i -lt $P_count;$i++)
            {   
                $Marsha_excel=Get-Content -Path "$results\marsha.txt" | select -Index $i

                $Block_rolling=Get-Content -Path "$results\Block rolling date.txt" | select -index $i
                if(!([String]::IsNullOrWhiteSpace($Block_rolling))){
                    if($Block_rolling -match "715"){if($Block_rolling_val -notmatch "MISSING"){$Block_rolling_val="VALID"}}
                    else{
                    $Marsha_block=$Marsha_block+" "+$Marsha_excel
                    $Block_rolling_val="MISSING : $Marsha_block"}
                }

                $mimpg_xml=Get-Content -Path "$results\mimpg_xml_version.txt" | select -Index $i
                if(!([String]::IsNullOrWhiteSpace($mimpg_xml))){
                    if($mimpg_xml -match "V3"){if($mimpg_xml_val -notmatch "MISSING"){$mimpg_xml_val="VALID"}}
                    else{
                    $Marsha_mimpg=$Marsha_mimpg+" "+$Marsha_excel
                    $mimpg_xml_val="MISSING : $Marsha_mimpg"}
                }

                $Oxi_comm=Get-Content -Path "$results\OXI Comm Method_ERS Rate Opera.txt" | select -index $i
                if(!([String]::IsNullOrWhiteSpace($Oxi_comm))){
                    if($Oxi_comm -match 'https://marshaprod.marriott.com:5617/reservationgpms'){if($Oxi_comm_val -notmatch "MISSING"){$Oxi_comm_val="VALID"}}
                    else{
                    $Marsha_comm=$Marsha_comm+" "+$Marsha_excel
                    $Oxi_comm_val="MISSING : $Marsha_comm"}
                }


                $rate_codes=Get-Content -Path "$results\c13_rate_codes.txt" | select -Index $i
                if(!([String]::IsNullOrWhiteSpace($rate_codes))){
                    $Rate_Codes_val=check-Rate_Codes -P_count $i
                    if($Rate_Codes_val -match "Conflict"){
                    $Marsha_rate2=$Marsha_rate2+" "+$Marsha_excel
                    $Rate_Codes_val1="Conflict for $Marsha_rate2"}

                    if($Rate_Codes_val -match "MISSING"){
                    $Marsha_rate1=$Marsha_rate1+" "+$Marsha_excel
                    $Rate_Codes_val2="MISSING : $Marsha_rate1"}

                    $Rate_Codes_val_output=$Rate_Codes_val2+" "+$Rate_Codes_val1
                    if($Rate_Codes_val -match "VALID" -and $Rate_Codes_val_output -eq " "){$Rate_Codes_val_output="VALID"}
                }

                    $C16=validate-c16 $Marsha_excel
                    if($C16 -match 'VALID'){if($c16_val -notmatch "MISSING"){$c16_val="VALID"}}
                    else{
                    $Marsha_c16=$Marsha_c16+" "+$Marsha_excel
                    $c16_val="MISSING : $Marsha_c16"}

            }
      }
return $Oxi_comm_val,$Block_rolling_val,$mimpg_xml_val,$Rate_Codes_val_output,$c16_val
}

function get-country{
[CmdletBinding()]
    param (
        $P_count
    )
    $country=Get-Content "$results\country.txt" | select -index 0
    if(($country -match "no rows selected") -or ($country -eq $null) -or ([String]::IsNullOrWhiteSpace($country)))
        {
            $a=(Get-TimeZone).standardname
            $ac=Get-TimeZone
            $country=[regex]::Match($a,"(.*)(Standard Time)").captures.groups[1].value
            if($ac -match "Mexico" -or $ac -match "US `& Canada" -or $ac -match "Russia"){$country=$Matches[0]}
        }
    $city=Get-Content "$results\city.txt" | select -index $P_count
    if(($city -match "no rows selected") -or ($city -eq $null) -or ([String]::IsNullOrWhiteSpace($city)))
    {
        $b=(Get-TimeZone).displayname
        $city=[regex]::Match($b,"(\(UTC.*\) )(.*)").captures.groups[2].value
    }
    
    return $country,$city
    
}

function SnC_check{
[CmdletBinding()]
    param (
    [string]$property)
    if((Get-Content "$results\SnC_check.txt") -like 'Y'){
        if(($property -match "Ritz.*Carlton") -or ($property -match "Bulgari")){return "EVO settings are applicable please check manually"}
        else{return "ODC Settings are applicable please check manually"}
        }else{return "OPERA Sales & Catering license is not active, hence this correction is not applicable" }
}

Function NLS-uifont{
[CmdletBinding()]
    param (
        $str1
    )
    $hostname=hostname
    $path="D:\ORA\user_projects\domains\OperaDomain\config\fmwconfig\components\ReportsToolsComponent\reptools`_$hostname\tools\COMMON"
    $validation="VALID"
    $Reason="VALID"
    if(!(Test-Path -Path "D:\ORA\MWFR\12cappr2\tools\common\uifont.ali" -PathType Leaf))
    {
        $validation= "MISSING"
        $Reason= "Uifont.ali does not exist in D:\ORA\MWFR\12cappr2\tools\common"
    }
    $fileA=get-content "$path\uifont.ali"
    $fileB=get-content "D:\ORA\MWFR\12cappr2\tools\common\uifont.ali"
    if(Compare-Object $fileA $fileB)
    {
        $validation= "MISSING"
        $Reason="Instance of uifont at 12cappr2\tools\common do not match with original"
    }
    
    $value=$validation,$Reason
    return $value

}
function Validate_Opera_tools{
    [CmdletBinding()]
    param (
        $str1
    )
    $data =  @{}
   $oapp_ver=Get-Content "$results\OAppConf_ver.txt"
   $a = $oapp_ver -join "`n"
   $a -match "(file version\:\s*)(.+)" | Out-Null
   $data.Add("oapp",$matches.2)

   $SMT_ver=Get-Content "$results\OPERA_SMT_ver.txt"
   $b = $SMT_ver -join "`n"
   $b -match "(Original Name\:\s*)(.+)" | Out-Null
   $data.Add("SMT",$matches.2)

   $oxi_interface_version=Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Opera XChange Interface Processor" | Select-Object DisplayVersion
   $oxi_interface_version.DisplayVersion -match "(.+)( \(.+)" | Out-Null
   $data.Add("oxi",$matches.1)
   
   if($oapp_ver -match '5.4.3.70'){$data.Add("oapp_validation","VALID")}
   else{$data.Add("oapp_validation","MISSING")}
   if($SMT_ver -match "original name:.*OperaSMT 4\.0 \(30\/[ag]\)"){$data.Add("smt_validation","VALID")}
   else{$data.Add("smt_validation","MISSING")}
   if($oxi_interface_version.DisplayVersion -match "5.6.15.12" -or $oxi_interface_version.DisplayVersion -match "5.6.25.0"){$data.Add("oxi_validation","VALID")}
   else{$data.Add("oxi_validation","MISSING")}
   return $data
}


function C3-Validation {
    [CmdletBinding()]
    param (
    [string]$marsha)
    $data =  @{}
	$wl_service=Get-Service -name 'Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12214ohs_wlserver)' -ErrorAction SilentlyContinue
    if($wl_service.DisplayName -ne $null)
	{$b=sc.exe qc "Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12214ohs_wlserver)" | Select-String "START_TYPE" | ForEach-Object { ($_ -replace '\s+', ' ').trim().Split(" ") | Select-Object -Last 1 }}
	else
	{$b=sc.exe qc "Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12213ohs_wlserver)" | Select-String "START_TYPE" | ForEach-Object { ($_ -replace '\s+', ' ').trim().Split(" ") | Select-Object -Last 1 }}
    $a=sc.exe qc "Oracle Weblogic OperaDomain NodeManager (D_ORA_MWFR_12cappr2_wlserver)" | Select-String "START_TYPE" | ForEach-Object { ($_ -replace '\s+', ' ').trim().Split(" ") | Select-Object -Last 1 }
    $data.Add('b_v',"$a , $b")
    if(($a -match "Auto_start") -and ($b -match "Delayed")){$data.Add('b',"VALID")}
        else{$data.Add('b',"MISSING")}

    $c3c=Get-Content "D:\ORA\user_projects\domains\OperaDomain\config\jdbc\operaoperads-jdbc.xml"
    $c3c=$c3c -replace ' ',''
    $c_v=$c3c -match "<url>.*</url>"
    $data.Add('c_v',"$c_v")
    $hostname=hostname
    if($c3c -like "<url>jdbc:oracle:thin:@(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$hostname)(PORT=1521))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=OPERA)))</url>"){$data.Add('c',"VALID")}
    else{$data.Add('c',"MISSING")}

        
    $ipv6=Get-NetAdapterBinding -ComponentID ms_tcpip6 | where -Property enabled -eq $true
    #get-wmiobject win32_networkadapter -filter "netconnectionstatus = 2" | select netconnectionid, name, InterfaceIndex, netconnectionstatus
    $enable_adapters=Get-NetAdapter | Where-Object {$_.status -eq 'up' -or $_.status -eq 'Disconnected'}
    $networkConfig = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "ipenabled = 'true'"
    $required_adapter=$null
    foreach($e1 in $ipv6.name){
    foreach($e2 in $enable_adapters.name){
            if ($e1 -eq $e2){$required_adapter=$required_adapter+$e1+' + '}
        }
    }
    $data.Add('d_v',"Adapter where Ipv6 is enabled : $required_adapter and DNSDomain Suffix is $($networkConfig.DNSDomain)")
   
    if(($required_adapter -eq $null) -and ($networkConfig.DNSDomain -eq $null)){$data.Add('d',"VALID")}
    else{$data.Add('d',"MISSING")}
     
    $ip=Get-Content -Path "$results\IP_Address.txt" | Select -Index 0
        $hosts=Get-Content "C:\Windows\System32\drivers\etc\hosts"
        $e_v=Get-Content "C:\Windows\System32\drivers\etc\hosts" | select -skip 21
        $e_v2=$e_v -join ' '
        $e_v2=$e_v2.trim()
        $e_v2=$e_v2 -replace '	 ',' '

    $data.Add('e_v',"$e_v2")
    if(($hosts -match "$ip\s*$hostname") -and ($hosts -match "$ip\s*\w*\W*$marsha.marriott.com")){$data.Add('e',"VALID")}
    elseif(($hosts -match "$ip\t*$hostname") -and ($hosts -match "$ip\t*\w*\W*$marsha.marriott.com")){$data.Add('e',"VALID")}
        else{$data.Add('e',"MISSING")}

    $heap_size=get-Itemproperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\SubSystems" | select Windows
    $data.Add('f_v',$heap_size.windows)
    if($heap_size.windows -match "4096,20480,4096"){$data.Add('f',"VALID")}
         else{$data.Add('f',"MISSING")}

    $c3g1=Get-Content "D:\ORA\user_projects\domains\OperaOHSDomain\config\fmwconfig\components\OHS\instances\ohs1\moduleconf\forms.conf"
    $c3g2=Get-Content "D:\ORA\user_projects\domains\OperaOHSDomain\config\fmwconfig\components\OHS\ohs1\moduleconf\forms.conf"
    $g_v=($c3g1 -match "WebLogicCluster\s*$ip`:6001")
    $g_v=$g_v+($c3g2 -match "WebLogicCluster\s*$ip`:6001")
    $data.Add('g_v',$g_v)
    if(($c3g1 -match "WebLogicCluster\s*$ip`:6001") -and ($c3g2 -match "WebLogicCluster\s*$ip`:6001")){$data.Add('g',"VALID")}
        else{$data.Add('g',"MISSING")}

    return $data

}

function check-Rate_Codes {
    [CmdletBinding()]
    param (
    $P_count)
    $resort=Get-Content -Path "$results\marsha.txt" | select -Index $P_count
     if([String]::IsNullOrWhiteSpace($resort)){return $null}
    $rate_codes=Get-Content -Path "$results\c13_rate_codes.txt" | select -Index $P_count
		if(($rate_codes -match "FIXME") -or ($rate_codes -match "NORATE")){return "MISSING"}
	if(!((Get-Content -Path "$scripts\ERS_Hotel_list.txt") -match $resort))
        {
        if(($rate_codes -match "10XDRZ") -and ($rate_codes -match "12XDRZ")){return "Conflict : Property Not Found In ACTIVE ERS Database, But ERS Rate Codes Are Configured. Please Check For Updated List"}

    }
    if((Get-Content -Path "$scripts\ERS_Hotel_list.txt") -match $resort)
    {
        if(($rate_codes -match "10XDRZ") -and ($rate_codes -match "12XDRZ")){return "VALID"}
        else{return "MISSING"}
    }
    else{
    if(($rate_codes -match "no rows selected") -or ($rate_codes -eq $null) -or ([String]::IsNullOrWhiteSpace($rate_codes))){return "VALID"}
    else{return "MISSING"}
    }
}

function check-freedomfix{
    [CmdletBinding()]
    param (
        $str1
    )
    $freedomfix=get-content "$results\Hotfix.txt"
    $count=0
    $patchlist="freedom_fix_ff25"
    foreach ($element in $patchlist)
    {
        if($freedomfix -match $element){
        $count=$count+1
        "$element,"| Out-File "$results\applied_freedomfixes.txt" -append
        }
    }
    $balance_date=gci "D:\MICROS\opera\production\runtimes\balance.fmx" | select LastWritetime
	$balance_date=$balance_date.LastWriteTime.ToString('MM dd yy')
    #$b_date_sql=Get-Content "$results\Freedom_fix_date.txt"
    #$b_date_sql=[string]$b_date_sql
    #$b_date_sql=[datetime]::parseexact($b_date_sql, 'dd-MMM-yy', $null)
        if (($count -eq 1) -and ($balance_date -eq "11 15 23")) {
        return "VALID"}
        else{return "MISSING"}
}
function match-withObject {
    [CmdletBinding()]
    param (
    $object,
    [string]$match_string)
    if([string]::IsNullOrWhiteSpace($object)){return "MISSING"}
    if($object -eq $null){return "MISSING"}
    foreach($element in $object){
    if($element -match $match_string){return "VALID"}
    }
    if($object -notmatch $match_string){return "MISSING"}
}
function check-WLOpatch {
    [CmdletBinding()]
    param (
        $WL_version
    )
    $count=0
 $wl_service=Get-Service -name 'Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12214ohs_wlserver)' -ErrorAction SilentlyContinue
 if($wl_service.DisplayName -ne $null){
        foreach ($element in $WL_version)
    {
        if($element -match "13.9.4.2.18"){
        $count=$count+1
        }
        if($element -match "12.2.0.1.45"){
        $count=$count+1
        }
    }
        if ($count -eq 3){return "VALID"}
        else{return "MISSING"}
 }
 else{
    foreach ($element in $WL_version)
    {
        if($element -match "13.9.4.2.1[2-9]"){
        $count=$count+1
        }
    }
        if ($count -eq 2){return "VALID"}
        else{return "MISSING"}
  }
}

function check-Weblogicpatches {
    [CmdletBinding()]
    param (
        $WL_version
    )
    $missing_patches=@()
    $count=0
    $Valid_group=0

    $wl_service=Get-Service -name 'Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12214ohs_wlserver)' -ErrorAction SilentlyContinue
    $WL_MWFR=Get-Content -Path "$results\Weblogic_MWFR.txt"
    $WL_OHS=Get-Content -Path "$results\Weblogic_OHS.txt"

    #checking for 12214
    if($wl_service.DisplayName -ne $null)
        {

            #for MWFR
            $missing_patches+="Missing in D:\ORA\MWFR\12cappr2`n"
            foreach ($element in Get-Content -Path "$scripts\Weblogic_patch_MWFR_12214.txt")
            {
                if($WL_MWFR -match $element){$count=$count+1
                }else{$missing_patches+=$element}
            }
            if($count -eq 33){$Valid_group+=1}
            $count=0

            #for OHS
            $missing_patches+="Missing in D:\ORA\12214ohs`n"
            foreach ($element in Get-Content -Path "$scripts\Weblogic_patch_OHS_12214.txt")
            {
                if($WL_OHS -match $element){$count=$count+1
                }else{$missing_patches+=$element}
            }
            if($count -eq 17){$Valid_group+=1}
            $count=0
         $missing_patches | out-file "$results\Weblogic_lspatches_missing.txt" -append
        
            if($Valid_group -eq 2){return "VALID"}else{return "MISSING"}
        }
    #checking for 12213
    else{

        #for MWFR
    $missing_patches+="Missing in D:\ORA\MWFR\12cappr2`n"
    foreach ($element in Get-Content -Path "$scripts\Weblogic_patch_MWFR.txt")
    {
        if($WL_MWFR -match $element){$count=$count+1
        }else{$missing_patches+=$element}
    }
    if($count -eq 38){$Valid_group+=1}
    $count=0

    #for OHS
    $missing_patches+="Missing in D:\ORA\12213ohs`n"
    foreach ($element in Get-Content -Path "$scripts\Weblogic_patch_OHS.txt")
    {
        if($WL_OHS -match $element){$count=$count+1
        }else{$missing_patches+=$element}
    }
    if($count -eq 18){$Valid_group+=1}
    $count=0
         
    $missing_patches | out-file "$results\Weblogic_lspatches_missing.txt" -append
    
        if($Valid_group -eq 2){return "VALID"}else{return "MISSING"}
    }
        
        
}

function check-Weblogicclient {
    [CmdletBinding()]
    param (
        $WL_version
    )
    $count=0
	$missing_patches=@()
    if($DB_service -eq "19c"){$WL_Client=Get-Content -Path "$results\Weblogic_19cClient.txt"}
    else{$WL_Client=Get-Content -Path "$results\Weblogic_1221Client.txt"}
    #for 1221Client and 19c client
    $missing_patches+="Missing in D:\ORA\1221Client or D:\ORA\19c\19cClient `n"
	if($DB_service -eq "19c"){
    foreach ($element in Get-Content -Path "$scripts\Weblogic_patch_19cClient.txt")
    {
        if($WL_Client -match $element){$count=$count+1
        }else{$missing_patches+=$element}
    }
    }else{
    foreach ($element in Get-Content -Path "$scripts\Weblogic_patch_1221Client.txt")
    {
        if($WL_Client -match $element){$count=$count+1
        }else{$missing_patches+=$element}
    }
  }
    $missing_patches | out-file "$results\Weblogic_lspatches_missing.txt" -append
	if($count -eq 2){return "VALID"}else{return "MISSING"}
}


function Get-EAR_version {
    [CmdletBinding()]
    param (
        [string]$str3)
    (New-Object System.Net.WebClient).DownloadString("http://$str3.marriott.com:9004/OPERASVC/integ/IntegrationGenericServices") > "$results\Oiw_servlet.html"
    Start-Sleep -Seconds 2
    $match=select-string -Path "$results\Oiw_servlet.html" -pattern "Build Version:"
    if($match -match "(.+Version: )(.+)(.\(Git)")
    {
    $ver=$matches.2
    $ver | Set-Content "$results\EAR_version.txt"
    return $ver
    }else{return "Unable to fetch EAR Version"}
}

function check-dbpatches_19c {
    [CmdletBinding()]
    param (
        $DB_lspatches
    )
    $count=0
    $patchlist=Get-Content -Path "$scripts\DB_patch_list_19c.txt"
    foreach ($element in $patchlist)
    {
        if($DB_lspatches -match $element){
        $count=$count+1
        }
    }
        $verbose_check=Get-Content -Path "$results\DB_verbose_check_19c.txt"
       
    if(($verbose_check -match "37486199") -and ($verbose_check -match "37102264")){$count=$count+1}

        
    if ($count -eq 4){return "VALID"}
        else{return "MISSING"}
}

function check-db_timezone {
    [CmdletBinding()]
    param (
        $str1
    )
     if((get-content "$results\DB_Time_zone.txt") -match "timezlrg_43.dat ")    
    {return "VALID"}
    else{return "MISSING"}
}

function check-dbpatches {
    [CmdletBinding()]
    param (
        $DB_lspatches
    )
    $count=0
    $patchlist=Get-Content -Path "$scripts\DB_patch_list.txt"
    foreach ($element in $patchlist)
    {
        if($DB_lspatches -match $element){
        $count=$count+1
        }
    }
        $verbose_check=Get-Content -Path "$results\DB_verbose_check.txt"
       
        if(($verbose_check -match "33577550") -and ($verbose_check -match "33488333")){$count=$count+1}
    
    if ($count -eq 5){return "VALID"}
        else{return "MISSING"}
}
function get-menu {
    [CmdletBinding()]
    param (
        [string]$str3
    )
    Clear-Host
    Write-Host "================ WELCOME To OPERA DATA COLLECTION SCRIPT ================" -ForegroundColor Cyan
    Write-Host "Please choose the activity" -ForegroundColor Cyan
    Write-Host "Press '1' for Go-Back `(FULL OPERA Upgrade and Security Patches)." -ForegroundColor Cyan
    Write-Host "Press '2' for Go-Back `(Part-1 OPERA Upgrade`)." -ForegroundColor Cyan
    Write-Host "Press '3' for Go-Back `(Part-2 Security Patches`)." -ForegroundColor Cyan
    Write-Host "Press '4' for Upgrade." -ForegroundColor Cyan
    Write-Host "Press '5' for Hypercare." -ForegroundColor Cyan
    $input=Read-Host "`n"
    switch ($input)
{
    1 { return "Go-Back-FULL" }
    2 { return "Go-Back-1" }
    3 { return "Go-Back-2" }
    4 { return "Upgrade" }
    5 { return "Hypercare" }
    default{Write-Host "Please select a Valid Responce" -ForegroundColor Darkred
    Start-Sleep -Seconds 2.5
    $input=get-menu
    return $input}
}

} 


function get-Owner {
    [CmdletBinding()]
    param (
        [string]$str3
    )
    Write-Host "Please Select Ownership Type" -ForegroundColor Cyan
    Write-Host "Press '1' for Managed" -ForegroundColor Cyan
    Write-Host "Press '2' for Franchised" -ForegroundColor Cyan
    $input=Read-Host "`n"
    switch ($input)
{
    1 { return "M" }
    2 { return "F" }
    default{Write-Host "Please select a Valid Responce" -ForegroundColor Darkred
    Start-Sleep -Seconds 2.5
    $input=get-Owner
    return $input}
}

}

function get-region {
    [CmdletBinding()]
    param (
        [string]$str3
    )
    Write-Host "Please Choose the Property Region" -ForegroundColor Cyan
    Write-Host "Press '1' for US/CAN" -ForegroundColor Cyan
    Write-Host "Press '2' for CALA" -ForegroundColor Cyan
    Write-Host "Press '3' for EMEA" -ForegroundColor Cyan
    Write-Host "Press '4' for APAC." -ForegroundColor Cyan
    $input=Read-Host "`n"
    switch ($input)
{
    1 { return "US/CAN" }
    2 { return "CALA" }
    3 { return "EMEA" }
    4 { return "APAC" }
    default{Write-Host "Please select a Valid Responce" -ForegroundColor Darkred
    Start-Sleep -Seconds 2.5
    $input=get-region
    return $input}
}

}

function check-files {
    [CmdletBinding()]
    param (
        [string]$str3
    )
    $count=0
    $files="Opera_Central_Systems_Analyzer.rep","Opera_DB_Analyzer.rep","Opera_EOD_Analyzer.rep","Opera_Health_check.rep","Opera_Interface_Analyzer.rep","Opera_OXIHUB_Analyzer.rep","Opera_OXI_Analyzer.rep","Opera_Printing_Analyzer.rep"

    foreach($element in $files){
    
    if(Test-Path -Path "d:\micros\opera\production\runtimes\$element" -PathType Leaf)
    {
        $count=$count+1
    }
    else
    {
         "File $element does not exist " | Out-File "$results\c10_runtimes_files.txt" -Append
         $missing_files=$missing_files+' '+$element
    }
    
    }
if($Count -eq 8 -and ($Report_analyzer[1] -match "Opera Analyzer"))
    {
     return "VALID"
    }
else{return "MISSING Runtime files: $missing_files"}
}

function Check-Hotfix {
    [CmdletBinding()]
    param (
        [string]$hotfix_var
    )
	
$SEL = get-content "$results\Hotfix.txt"


if($SEL -imatch "36442317")
{
   $res="VALID"
}
else{   
   $res="Not Applicable"
	}	
  return $res
}
function execute-sql {
    [CmdletBinding()]
    param (
        [string]$sys_password
    )
    "conn sys/$sys_password@opera as sysdba 
    set newpage 0; 
    set echo off;
    SET LINESIZE 32767;
    SET TRIMSPOOL ON; 
    set feedback off;
    SET WRAP OFF; 
    set pagesize 0 ;
    set heading off;
    @""c:\Temp\OASIS\scripts\Country.sql""
    @""c:\Temp\OASIS\scripts\City.sql""
    @""c:\Temp\OASIS\scripts\property_name.sql""
    @""c:\Temp\OASIS\scripts\role.sql""
    @""c:\Temp\OASIS\scripts\DB_Version.sql""
    @""c:\Temp\OASIS\scripts\New Installed Date.sql""
    @""c:\Temp\OASIS\scripts\New PMS Version.sql""
    @""c:\Temp\OASIS\scripts\New E-Patch Level.sql""
    @""c:\Temp\OASIS\scripts\OPERA_Schema_size.sql""
    @""c:\Temp\OASIS\scripts\OXI_Schema_size.sql""
    @""c:\Temp\OASIS\scripts\Opera_OIW_Schema_size.sql""
    @""c:\Temp\OASIS\scripts\Archive Logs Status.sql""
    @""c:\Temp\OASIS\scripts\Session timeout.sql""
    set feedback on;
    @""c:\Temp\OASIS\scripts\Disable_audit_policy.sql""
    @""c:\Temp\OASIS\scripts\Receive Broadcast.sql""
    @""c:\Temp\OASIS\scripts\Authentication Provider.sql""
    set feedback off;
    @""c:\Temp\OASIS\scripts\Screen Painter_OpportunityID.sql""
    @""c:\Temp\OASIS\scripts\OXI Comm Method_ERS Rate Opera.sql""
    @""c:\Temp\OASIS\scripts\Block rolling date.sql""
    @""c:\Temp\OASIS\scripts\Report Analyzer.sql""
    @""c:\Temp\OASIS\scripts\Hotfix.sql""
    @""c:\Temp\OASIS\scripts\Freedom_fix_date.sql""
    @""c:\Temp\OASIS\scripts\DB_verbose_check.sql""
    @""c:\Temp\OASIS\scripts\DB_verbose_check_19c.sql""
    @""c:\Temp\OASIS\scripts\mimpg_xml_version.sql""
    @""c:\Temp\OASIS\scripts\SnC_check.sql""
    @""c:\Temp\OASIS\scripts\c13_rate_codes.sql""
    @""c:\Temp\OASIS\scripts\Schema_version.sql""
	@""c:\Temp\OASIS\scripts\IP_Address.sql""
    @""c:\Temp\OASIS\scripts\C15_WS_Alive.sql""
    @""c:\Temp\OASIS\scripts\BTR_CHECK.sql""
    @""c:\Temp\OASIS\scripts\C16_disable_crisis_export.sql""
    @""c:\Temp\OASIS\scripts\C18_Table_Space_Verification.sql""
    @""c:\Temp\OASIS\scripts\C18_Tempseg.sql""
    @""c:\Temp\OASIS\scripts\C17-Loyalty_Stay_Export_Delivery_Config.sql""
    @""c:\Temp\OASIS\scripts\C12_RitzCarltonRewards.sql""
    @""c:\Temp\OASIS\scripts\C14_Business_Block_notes.sql""
    @""c:\Temp\OASIS\scripts\DB_Time_zone.sql""
    @""c:\Temp\OASIS\scripts\OPI_check.sql""
	@""c:\Temp\OASIS\scripts\acl.sql""
    @""c:\Temp\OASIS\scripts\IFC_active.sql""
    @""c:\Temp\OASIS\scripts\IFC_machine.sql""
	@""c:\Temp\OASIS\scripts\IFC_version.sql""
	@""c:\Temp\OASIS\scripts\retention.sql""
	@""c:\Temp\OASIS\scripts\index_status.sql""
	@""c:\Temp\OASIS\scripts\train_user.sql""
	exit" | sqlplus  /nolog

}

function Is-MultProperty {
    [CmdletBinding()]
    param (
        [string]$sys_password
    )
    "conn sys/$sys_password@opera as sysdba 
    set newpage 0; 
    set echo off;
    SET LINESIZE 32767;
    SET TRIMSPOOL ON; 
    set feedback off;
    SET WRAP OFF; 
    set pagesize 0 ;
    set heading off;
    @""c:\Temp\OASIS\scripts\Marsha.sql""
    exit" | sqlplus  /nolog

}

#set-script
$global:results="c:\Temp\OASIS\results"
$global:scripts="c:\Temp\OASIS\scripts"
$global:logs="C:\temp\OASIS\Logs"

Start-Transcript -Path "$logs\OASIS.OPERA.v1.0.Activity.log" -Append

$script_release="6.2"
Write-Output "Script Release Number :$script_release"
Write-Output "Released on 05-Sep-2024 14:20 IST"

$DB_Release="Jan-2025"
$DBOpatch_release="12.2.0.1.45"
$Weblogic_Release="Jan-25"
$WLOpatch_release="13.9.4.2.18"
$JDK_release="441"
$wlclient_release="Jan-25"
$Opera_release="5.6.15.20/5.6.25.5"
$EAR_release="1.13.0.15/1.25.5.0"
$FF_release="freedom_fix_ff25"
$Oiw_release="21.1.2.14/5.6.24.0"
$OAPP_release="5.4.3.70"
$smt_release="4.0 (30a/30g)"
$OXIProcessor_release="5.6.15.12/5.6.25.0"


#if((Get-Date) -gt ([datetime]::parseexact('16-10-22', 'dd-MM-yy', $null))){Write-Warning 'Expired Script'
#Start-Sleep -Seconds 10
#exit}'
if(!(Test-Path "$scripts")){mkdir $scripts}
if(!(Test-Path "$results")){mkdir $results}
$dat=get-date -UFormat "%d-%B-%y"
if(!(Test-Path "D:\Patching_$dat")){mkdir D:\Patching_$dat}
#robocopy "scripts" "$scripts" /e
Remove-Item "$results\*.txt" -Recurse -Confirm:$false
Clear-Host
Write-Host "Welcome to Opera Information collection Script`n" -ForegroundColor Green
$Activity = get-menu
if(Test-Path "$scripts\manual_entry.txt"){
$region=Get-Content "$scripts\manual_entry.txt" | select -Index 0
$ownership=Get-Content "$scripts\manual_entry.txt" | select -Index 1
}
else{
$region=get-region
$ownership=get-owner
$region | Out-File "$scripts\manual_entry.txt"
$ownership | Out-File "$scripts\manual_entry.txt" -Append
}
#$password = Read-Host "Please Enter Sys Schema Password" -AsSecureString
#$Newpass = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))

if ( Test-Path -Path "$scripts\*" -include *.sql )
{
        #do-nothing    
}
else
{
    robocopy "$((get-location).path)\scripts" "$scripts" /e

}

#MULTI-Property Check
Is-MultProperty -sys_password "Newpass"
$Multi_property_check=Get-Content -Path "$results\marsha.txt" | select -Index 1
$global:P_count=1
if(!([String]::IsNullOrWhiteSpace($Multi_property_check))){
$marsha=Read-host "Hi Script has noticed that this is a Multi-property Server, Please Enter The Primary Marsha (OPERA URL MARSHA)`n"
$property_type="Multi-Property"

For(;;){
	if(!($marsha -cmatch "^[A-Z]{5}$")){$marsha=Read-host "INVALID Marsha response, please Re-Enter Marsha `n"}
else{break}
}
$Marsha | out-file "$results\host_marsha.txt"
    For(;;){
        $Multi_property_check=Get-Content -Path "$results\marsha.txt" | select -Index $P_count
        if(!([String]::IsNullOrWhiteSpace($Multi_property_check))){$P_count++}
        else{break}
    }
}else{$Marsha=Get-Content -Path "$results\marsha.txt" | select -Index 0
$property_type="Stand-alone"
$Marsha | out-file "$results\host_marsha.txt"
}

$Marsha_multi=Get-Content -Path "$results\marsha.txt"
$Marsha_multi=[system.String]::Join("/", $Marsha_multi)
$Marsha_excel=Get-Content -Path "$results\marsha.txt" | select -Index 0
#END of MULTI-Property Check
if($property_type -ne "Stand-alone"){
if($Marsha_excel -eq $marsha -and $P_count -gt 1){$property_type="Multi-Property"}else{$property_type="Part of Multi-Property"}}
$Multi_Property_Marshas=$Marsha_multi -replace "(^$marsha/)?(/$marsha)?",$null
if($P_count -eq 1){$Multi_Property_Marshas=$null}
#end of multiproperty variables set
execute-sql -sys_password "Newpass"
#checking Db version
$DB_service=Get-Service -Name 'OracleTNSListenerV2' -ErrorAction SilentlyContinue
if($DB_service -eq $null){$DB_service="12c"}else{$DB_service="19c"}
Write-Output "Database Version is $DB_service`n"
$opi_versions=OPI-check

#collectdata
Write-Host "`nPlease wait Collecting Server Information`n" -ForegroundColor yellow
#$ip=Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $(Get-NetConnectionProfile | Select-Object -ExpandProperty InterfaceIndex) | Select-Object -ExpandProperty IPAddress
$cpu_details=Get-WmiObject -class win32_processor -Property  ”name”, “numberOfCores” , "maxclockspeed" | Select-Object -Property ”name”, “numberOfCores” , "maxclockspeed"
$cpu_details="$(($cpu_details.name).trim()) $($cpu_details.maxclockspeed) $($cpu_details.numberOfCores) Cores"
#$Total_RAM=(systeminfo | Select-String 'Total Physical Memory:').ToString() -replace '\D', '' | ForEach-Object { ([math]::Round($_ / 1KB, 2)).Tostring() + " GB" }
$Total_RAM=((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb).ToString() + " GB"
$total_HDD=(Get-CimInstance Win32_LogicalDisk | Measure-Object -Property Size -Sum | Select-Object -ExpandProperty Sum | % {[math]::Round($_ / 1GB,2)}).Tostring() + " GB"
$hdd_sizes=Get-CimInstance Win32_LogicalDisk | Select-Object DeviceID , size,freespace
$total_HDD="C:$([math]::Round($hdd_sizes[0].size / 1GB,2)) GB free-$([math]::Round($hdd_sizes[0].freespace / 1GB,2)) GB ,D:$([math]::Round($hdd_sizes[1].size / 1GB,2)) GB free-$([math]::Round($hdd_sizes[1].freespace / 1GB,2))GB Total Space-$($total_HDD)"

$Report_date=get-date -UFormat "%d-%B-%y %r"
$hostname=hostname
$ip_address=Get-Content -Path "$results\IP_Address.txt"
$runtime_version=Get-Content -Path "d:\micros\opera\production\runtimes\opera_pms.ins"
cmd /c "$scripts\Weblogic_version.bat"
if($DB_service -eq "19c"){$nls1=Get-ItemProperty -path "HKLM:\SOFTWARE\ORACLE\KEY_OraDB19Home1" | select nls_lang}
else{$nls1=Get-ItemProperty -path "HKLM:\SOFTWARE\ORACLE\KEY_ORA12201" | select nls_lang}
$nls2=Get-ItemProperty -path "HKLM:\SOFTWARE\ORACLE\KEY_OracleHome1" | select nls_lang
$nls3=Get-ItemProperty -path "HKLM:\SOFTWARE\ORACLE\KEY_OracleHome2" | select nls_lang
$nls4=Get-ItemProperty -path "HKLM:\SOFTWARE\ORACLE\KEY_OracleHome3" | select nls_lang
$nls5=Get-ItemProperty -path "HKLM:\SOFTWARE\WOW6432Node\ORACLE\KEY_OraClient19cHome1_32bit" | select nls_lang
$NLS_lang="$($nls1.NLS_LANG) $($nls2.NLS_LANG) $($nls3.NLS_LANG) $($nls4.NLS_LANG) $($nls5.NLS_LANG)"
write-host "`n`nPlease Wait collecting Security Patches information`n`n" -ForegroundColor yellow
cmd /c "set JAVA_HOME=D:\ORA\JDK"
cmd /c "cd /D D:\ORA\JDK\jre\bin"
$JDK_version= & cmd /c "java -version 2>&1"
$JDK_version=$JDK_version[0] -replace """" ,''

if($DB_service -eq "19c"){cmd /c "$scripts\Weblogic_check_19c.bat"
cmd /c "$scripts\DB_check_19c.bat"}
else{cmd /c "$scripts\Weblogic_check.bat"
cmd /c "$scripts\DB_check.bat"}

cmd /c "$scripts\Opera_tools_version.bat" | Out-Null

$result_list = Get-ChildItem $results | select name
foreach($element in $result_list ) {(get-content "$results\$($element.name)") | ? {$_.trim() -ne "" } | set-content "$results\$($element.name)"}
foreach($element in $result_list ) {(get-content "$results\$($element.name)") -replace ";","," | set-content "$results\$($element.name)"}
foreach($element in $result_list ) {(get-content "$results\$($element.name)") -replace "no rows selected","null" | set-content "$results\$($element.name)"}
$Weblogic_server=Get-Content "$results\Weblogic_version.txt" | Select -Index 0

$wl_service=Get-Service -name 'Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12214ohs_wlserver)' -ErrorAction SilentlyContinue
if($wl_service.DisplayName -ne $null)
{
$Delayed_auto_start=Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Services\Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12214ohs_wlserver)" | select DelayedAutostart
$Auto_start_Delay=Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Services\Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12214ohs_wlserver)" | select AutoStartDelay
}
else
{
$Delayed_auto_start=Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Services\Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12213ohs_wlserver)" | select DelayedAutostart
$Auto_start_Delay=Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Services\Oracle Weblogic OperaOHSDomain NodeManager (D_ORA_12213ohs_wlserver)" | select AutoStartDelay
}
write-host "`n`nPlease Wait Validating Collected Information`n`n" -ForegroundColor yellow
#SQL_Data_collection

$property=Get-Content -Path "$results\property_name.txt" | select -Index 0
$Country_0_city_1=get-country -P_count 0
$role=Get-Content -Path "$results\role.txt"
$Version=Get-Content -Path "$results\DB_Version.txt" | select -Index 1
$NInstall_Date=Get-Content -Path "$results\New Installed Date.txt"
$NPMS_Version=Get-Content -Path "$results\New PMS Version.txt"
#Script Created By Rajarshi Rahul Dwivedi
$NEpatch=Get-Content -Path "$results\New E-Patch Level.txt"
$Opera_size=Get-Content -Path "$results\OPERA_Schema_size.txt"
$Oxi_size=Get-Content -Path "$results\OXI_Schema_size.txt"
$Opera_OIW_size=Get-Content -Path "$results\Opera_OIW_Schema_size.txt"
$Opera_size=$Opera_size -replace " ",""
$Oxi_size=$Oxi_size -replace " ",""
$Opera_OIW_size=$Opera_OIW_size -replace " ",""
$Archive_log=Get-Content -Path "$results\Archive Logs Status.txt"
$STimeout=Get-Content -Path "$results\Session timeout.txt"
(Get-Content -Path "$results\Receive Broadcast.txt" -raw) -replace "1 row selected\.","Null" | Set-Content "$results\Receive Broadcast.txt"
(Get-Content -Path "$results\Receive Broadcast.txt" -raw) -replace "2 rows selected\.","Null" | Set-Content "$results\Receive Broadcast.txt"
(Get-Content -Path "$results\Authentication Provider.txt" -raw) -replace "1 row selected\.","Null" | Set-Content "$results\Authentication Provider.txt"

$Rbrodcast=Get-Content -Path "$results\Receive Broadcast.txt"
$AProvider=Get-Content -Path "$results\Authentication Provider.txt"
$Screen_PO=Get-Content -Path "$results\Screen Painter_OpportunityID.txt"
$Oxi_comm=Get-Content -Path "$results\OXI Comm Method_ERS Rate Opera.txt" | select -index 0
$Block_rolling=Get-Content -Path "$results\Block rolling date.txt" | select -index 0
$Report_analyzer=Get-Content -Path "$results\Report Analyzer.txt"
$DB_lspatches=Get-Content -Path "$results\DB_lspatches.txt" 
$DB_version=Get-Content -Path "$results\DB_Opatch_Version.txt" | select -First 4
$WL_lspatches=(Get-Content -Path "$results\Weblogic_MWFR.txt")+(Get-Content -Path "$results\Weblogic_OHS.txt")+(Get-Content -Path "$results\Weblogic_1221Client.txt")+(Get-Content -Path "$results\Weblogic_SPB.txt")
$WL_version=Get-Content -Path "$results\WLS_Opatch_Version.txt"
$audit_policy=Get-Content -Path "$results\Disable_audit_policy.txt"
$audit_policy=$audit_policy -replace "no rows selected.*","Null"
$freedom_fix_Validation=check-freedomfix
$C5_ODC=SnC_check $property
if((Get-Content "$results\SnC_check.txt") -like 'Y')
{$C7c="C7-c Correction is applicable, please check manually"}
else{$C7c="OPERA Sales & Catering license is not active, hence this correction is not applicable"}
$mimpg_xml=Get-Content -Path "$results\mimpg_xml_version.txt" | select -Index 0
$schema_version=Get-Content -Path "$results\Schema_version.txt"
$Ear_version=Get-EAR_version $Marsha
$c15_ws_alive=Get-Content "$results\C15_WS_Alive.txt"
$btr_value=Btr_check $marsha
$c16_value=Get-Content "$results\C16_disable_crisis_export.txt"
$c16_validation=validate-c16 $Marsha_excel
$c17_file=(Get-Content "$results\C17-Loyalty_Stay_Export_Delivery_Config.txt").trim()
(validate-c18_tablespace)[1].split(',')
Write-Host "`n `nPlease wait Creating csv file" -ForegroundColor yellow

$c2_Validation="$STimeout; $($STimeout -like "15");$Rbrodcast;$(($Rbrodcast -match " Y " ) -or ($Rbrodcast -match "Null"));$AProvider;$(($AProvider -match "NULL") -or $($AProvider -notmatch "OPERA_PORTAL"))"
$C3_Validation_a="$(($Auto_start_Delay.AutoStartDelay -like "180") -and ($Delayed_auto_start.DelayedAutostart -like "1"))"
$Archive_log_Validation=match-withObject $Archive_log "No Archive Mode"
$Archive_log_Validation=$Archive_log_Validation -replace 'VALID' , 'Disabled' -replace 'MISSING','Enabled'
$c3=C3-Validation $Marsha
$opera_tools=Validate_Opera_tools
$time_zone=(Get-TimeZone | select displayname).displayname
$time=get-date -UFormat "%I:%M %p"
$time_2=get-date -UFormat "%I.%M %p"
if($DB_service -eq "19c"){
$Db_patches_call=check-dbpatches_19c -DB_lspatches $DB_lspatches
$db_Opatch_check=match-withObject $DB_version "12.2.0.1.45"
}
else{
$Db_patches_call=check-dbpatches -DB_lspatches $DB_lspatches
$db_Opatch_check=match-withObject $DB_version "12.2.0.1.3[7-9]"
}
if (check-hotfix -eq "VALID") {$hotfix="36442317"} else {$hotfix=" "}
$ifc_machine_count=Get-Content -Path "$results\IFC_machine.txt" | select -index 0
$ifc_machine_count=[string]$ifc_machine_count.trim()
$ifc_active_names=Get-Content -Path "$results\IFC_active.txt"
$ifc_active_names=[string]::Join(",",$ifc_active_names)
$ifc_active_names=$ifc_active_names.trim()
$IFC_version=Get-Content -Path "$results\IFC_version.txt"
$IFC_version=[string]::Join(",",$IFC_version)
$IFC_version=$IFC_version.trim()
#writedata to csv



$PSDefaultParameterValues['out-file:width'] = 2000 
"Marsha;Property Name;Ownership;Region;Country;City;Activity;Report Capture Date;Property Type;Host Marsha;Multi Property Marsha(s);Server Hostname;IP Address;RAM;CPU;Storage;DB Role;DB Version;BTR Applied (US Managed Properties);BTR Validation;Opera Schema Version;OXI Schema Version;OIW Schema Version;Runtimes Version;Upgrade Date;OPERA Schema Size (GB);OXI Schema Size (GB);OIW Schema Size (GB);Oapp Conf Tool Version;Opera SMT Tool Version;OXI Processor Shell Version;EAR Version;WebLogic Version;WebLogic Installed Patches;WebLogic Patches Validation;Weblogic Opatch Validation;Weblogic Client Patches Validation;DB Installed Patches;Database Patches Validation;DB Opatch Validation;Archive Logs Status;Archive Log Result;Installed HotFixes;HotFix Validation;Freedom Fix;Freedom Fix Validation;Disable SYSAUX policy;Disable SYSAUX policy Validation;C1-a (NLS Lang UiFont validation);C1-b (NLS Lang Registry value);C1-b (NLS Lang Registry value Validation);C2-a (Opera session timeout`);C2-a Validation ;C2-b (Receive Broadcast`);C2-b Validation;C2-c (Authentication Provider`);C2-c Validation;C3-a (OPERA Login After Reboot Registry Value);C3-a Validation;C3-b (Auto Start Delay);C3-b Validation;C3-c (WLS JDBC Connection String Format);C3-c Validation;C3-d (Network Adaptor Settings);C3-d validation;C3-e (Host File);C3-e Validation;C3-f (Heap Size);C3-f Validation;C3-g (Forms.conf file);C3-g validation;C5 (ODC,EVO Activation Settings);C6 (Scrren Painter_OpportunityID);C6 Validation;C7-a (OXI Comm Method);C7-a Validation;C7-b (UPDATE BUSINESS EVENT CONFIGURATION);C7-c (S&C Owner Override);C8 (Block rolling date);C8 Validation;C9 (KB Article for error INT-90002,396);C10 (Report Analyzer);C10 Validation;C11 MIMPG XML version;C11 Validation;C12 Ritz Carlton Rewards;C13 MARSHA Proxy Rate Codes;C13 Validation;C14 Business Block notes;C15 Check MARSHA OXI Global Parameters: WS Alive;C15 Validation;C16 Crisis Export Decommission;C16 Validation;C-17 Stay Export Value;C-17 Stay Export validation;C18 Tablespace;C18 Tablespace Validation;C19 Discovery Export Validation;OPI Version;MySql Version;Interface Names;Interface Machine Count;C20-SYSAUX Validation;Training User;IFC Version" | Out-File "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv" -Encoding ascii
"$Marsha_excel;$property;$ownership;$region;$($Country_0_city_1[0]);$($Country_0_city_1[1]);$Activity;$Report_date;$property_type;$marsha;$Multi_Property_Marshas;$hostname;$ip_address;$Total_RAM;$cpu_details;$total_HDD;$role;$Version;$($btr_value[0]);$($btr_value[1]);$($schema_version[0]);$($schema_version[1]);$($schema_version[2]);$runtime_version;$NInstall_Date;$Opera_size;$Oxi_size;$Opera_OIW_size;$($opera_tools['oapp']);$($opera_tools['smt']);$($opera_tools['oxi']);$Ear_version;$Weblogic_server;$WL_version $JDK_version, $WL_lspatches;$(check-Weblogicpatches) $Weblogic_Release;$(check-WLOpatch $WL_version);$(check-Weblogicclient);$DB_version ,$DB_lspatches;$Db_patches_call $DB_Release;$($db_Opatch_check);$Archive_log;$Archive_log_Validation;$Hotfix;$(check-hotfix);$(Get-Content "$results\applied_freedomfixes.txt");$freedom_fix_Validation;$audit_policy;$(match-withObject $audit_policy "Null");NA;$NLS_lang;NA;$c2_Validation;$($Auto_start_Delay.AutoStartDelay) $($Delayed_auto_start.DelayedAutostart);$C3_Validation_a;$($c3['b_v']);$($c3['b']);$($c3['c_v']);$($c3['c']);$($c3['d_v']);$($c3['d']);$($c3['e_v']);$($c3['e']);$($c3['f_v']);$($c3['f']);$($c3['g_v']);$($c3['g']);$C5_ODC;$Screen_PO;$(match-withObject $Screen_PO ""UDFC14"");$Oxi_comm;$((HTML_Multi_validation)[0]);Manual Check;$C7c;$Block_rolling;$($Block_rolling -match "715");Manual Check;$Report_analyzer;$(check-files);$mimpg_xml;$($mimpg_xml -match "V3");$(Validate-C12);$(Get-Content -Path "$results\c13_rate_codes.txt" | select -Index 0);$(check-Rate_Codes -P_count 0);$(Validate-C14);$c15_ws_alive;$(($c15_ws_alive[0] -match "https://marshaprod.marriott.com:5617/reservationgpms") -and ($c15_ws_alive[1] -match "Y"));$c16_value;$c16_validation;$c17_file;$(match-withObject $c17_file  "wmbprd.marriott.com");$((validate-c18_tablespace)[1]);$((validate-c18_tablespace)[0]);$(Validate-C19);$($opi_versions[0]);$($opi_versions[1]);$ifc_active_names;$ifc_machine_count;$(C20-SYSAUXValidation);$(Training_user);$IFC_version" | Out-File "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv" -Append -Encoding ascii
(Get-Content -Path "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv" -raw) -replace "True","VALID" -replace "False","MISSING" | Set-Content "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv"

$csv_input = Get-Content "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv"
$csv_data = ConvertFrom-Csv -Delimiter ";" -InputObject $csv_input
$csv_data
Start-Sleep -Seconds 3
$csv_data | Out-File "$logs\Simple_$marsha`_Compliance-report.txt"
$heap=@{}
$jason_output="{"
$csv_data.PSObject.Properties | ForEach-Object {

    $heap.add("$($_.Name)","$($_.Value)")
	$jason_output=$jason_output+"""$($_.Name)"":""$($_.Value)"","
}
$jason_output=$jason_output+"}"
$jason_output=$jason_output -replace ",}","}" -replace '\\','\\'
$jason_output | out-file "$logs\$Marsha_excel`_Compliance-report_$dat.json" -Encoding ascii

#Start of HTML Creation
Copy-Item -path  "$scripts\*.pdf" -Destination "$results" -Recurse
Copy-Item -Path "$scripts\Validation_Table.html" -Destination "d:\Patching_$dat\OPERA_Compliance_Status.html"
if((Test-Path "d:\Patching_$dat\$marsha`_OPERA_Compliance_Status_*.html")){
$ab=Get-ChildItem "d:\Patching_$dat\$marsha`_OPERA_Compliance_Status_*.html"
ren "$($ab.fullname)" "OLD_$($ab.name)"
}
$row_count=0
#Validation html file print
Write-Host "`n `nPlease wait Creating HTML validation file" -ForegroundColor yellow
"<body>
<div class=""container"">
<h1>$Marsha_multi - OPERA Compliance Status Report ($hostname)</h1>
<h3 class=pink>$property</h3>
<h3 class=pink>$dat`</h3>
<h3 class=pink>$time, $time_zone</h3>
<h5 class=pink>Completion Status: NA</h5>
<div class=grey>Security patch validation: $DB_release Release</div>
<div class=grey>Database Version $Version</div>
<div class=grey>$Weblogic_server</div>
<div class=grey>OPERA Version $($schema_version[0])</div>

<table class=""table table-hover"">
  <thead class=""thead-dark"">
    <tr>
      <th scope=""col"">#</th>
      <th scope=""col"">Activity</th>
      <th scope=""col"">Status</th>
      <th scope=""col"">Reference</th>
    </tr>
  </thead>
  <tbody>
     <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Opera Runtimes Version <span class=""grey"">(Expected Version - $Opera_release) </span></td>
      <td>$(($runtime_version[0] -match ""5.0.06.[12]5"") -and ($runtime_version[1] -match ""E000[02][05]""))</td>
      <td> </td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>OPERA Schema Version <span class=""grey"">(Expected Version - $Opera_release)</span></td>
      <td>$($schema_version[0] -match ""5.0.06.[12]5E000[02][05]"")</td>
      <td> </td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>OXI Schema Version <span class=""grey"">(Expected Version - $Opera_release)</span></td>
      <td>$($schema_version[1] -match ""5.0.06.[12]5E000[02][05]"")</td>
      <td> </td>
    </tr>
    

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>OIW Schema Version <span class=""grey"">(Expected Version - $Oiw_release)</span></td>
      <td>$($schema_version[2] -match ""21.1.2.14"" -or $schema_version[2] -match ""5.6.24.0"" )</td>
      <td> </td>
    </tr>
    

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td> OAppConf <span class=""grey"">(Expected Version - $OAPP_release)</span></td>
      <td>$($opera_tools['oapp_validation'])</td>
      <td> </td>
    </tr>
    

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>OPERA_SMT <span class=""grey"">(Expected Version - $smt_release )</span></td>
      <td>$($opera_tools['smt_validation'])</td>
      <td> </td>
    </tr>
    

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>OXI Processor Shell <span class=""grey"">(Expected Version - $OXIProcessor_release)</span></td>
      <td>$($opera_tools['oxi_validation'])</td>
      <td> </td>
    </tr>
    

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>EAR Version <span class=""grey"">(Expected Version - $EAR_release)</span></td>
      <td>$($Ear_version -match ""1.13.0.15_MAR"" -or $Ear_version -match ""1.25.5.0_MAR"" )</td>
      <td> </td>
    </tr>


	<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>HotFix Validation</td>
      <td>$(check-hotfix)</td>
      <td>Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Freedom Fix Validation <span class=""grey"">($FF_release)</span> </td>	  
      <td>$freedom_fix_Validation</td>
      <td>Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>

     <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Weblogic OPatch <span class=""grey"">(Expected Version - $WLOpatch_release)</span> </td>
      <td>$(check-WLOpatch $WL_version)</td>
      <td><a href=""$results\09b_MARRIOTT_OPERA_5.6_Weblogic_SecurityPatching.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D15F80BD57634925056FA62D14C66628DB76A94487A0?allowInterrupt=1"">Link</td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Weblogic Patches Validation <span class=""grey"">(Release: $Weblogic_Release)</span> </td>
      <td>$(check-Weblogicpatches)</td>
      <td><a href=""$results\09b_MARRIOTT_OPERA_5.6_Weblogic_SecurityPatching.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D15F80BD57634925056FA62D14C66628DB76A94487A0?allowInterrupt=1"">Link</td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Weblogic Client Patches Validation <span class=""grey"">(Release: $wlclient_release)</span> </td>
      <td>$(check-Weblogicclient)</td>
      <td><a href=""$results\09b_MARRIOTT_OPERA_5.6_Weblogic_SecurityPatching.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D15F80BD57634925056FA62D14C66628DB76A94487A0?allowInterrupt=1"">Link</td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>JDK Version validation<span class=""grey"">(Expected Version - $JDK_release)</span> </td>
      <td>$(Check-JDK_version)</td>
      <td><a href=""$results\09b_MARRIOTT_OPERA_5.6_Weblogic_SecurityPatching.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D15F80BD57634925056FA62D14C66628DB76A94487A0?allowInterrupt=1"">Link</td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>DB OPatch <span class=""grey"">(Expected Version - $DBOpatch_release)</span> </td>
      <td>$($db_Opatch_check)</td>
      <td><a href=""$results\09a_MARRIOTT_OPERA_5.6_DB_SecurityPatching.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D6732ED728F4ABAA9731520C897D2E702AAA17926C6F?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>DB Patches Validation <span class=""grey"">(Release: $DB_Release)</span> </td>
      <td>$Db_patches_call</td>
      <td><a href=""$results\09a_MARRIOTT_OPERA_5.6_DB_SecurityPatching.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D6732ED728F4ABAA9731520C897D2E702AAA17926C6F?allowInterrupt=1"">Link</a></td>
    </tr>

<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>DB Timezone Validation <span class=""grey"">(Release: $DB_Release)</span> </td>
      <td>$(check-db_timezone)</td>
      <td><a href=""$results\09a_MARRIOTT_OPERA_5.6_DB_SecurityPatching.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D6732ED728F4ABAA9731520C897D2E702AAA17926C6F?allowInterrupt=1"">Link</a></td>
    </tr>
    
    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Disable SYSAUX Policy Validation</td>
      <td>$(match-withObject $audit_policy ""Null"")</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_Disable_UnifiedAuditPolicy.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D68B964691DBAB78F3435DC93DD3AABA6E7350724052?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>BTR Applied (US Managed Properties)</td>
      <td>$($btr_value[1])</td>
      <td><a href=""$results\BTR Issue script.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp -</td>
    </tr>
    
    
    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-2a (Opera Session Timeout)</td>
      <td>$($STimeout -like ""15"")</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_SessionTimeout.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-2b (Receive Broadcast)</td>
      <td>$(($Rbrodcast -match "" Y "" ) -or ($Rbrodcast -match ""Null""))</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_SessionTimeout.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-2c (Authentication Provider)</td>
      <td>$(($AProvider -match ""NULL"") -or $($AProvider -notmatch ""OPERA_PORTAL""))</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_SessionTimeout.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

        <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-3a (OPERA Login After Reboot Registry Value Validation)</td>
      <td>$C3_Validation_a</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_JobAid_OPERALogin_AfterReboot.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-3b (Auto Start Delay Validation)</td>
      <td>$($c3['b'])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_JobAid_OPERALogin_AfterReboot.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-3c (WLS JDBC Connection String Format)</td>
      <td>$($c3['c'])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_JobAid_OPERALogin_AfterReboot.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-3d (Network Adaptor Settings Validation)</td>
      <td>$($c3['d'])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_JobAid_OPERALogin_AfterReboot.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-3e (Host File Validation)</td>
      <td>$($c3['e'])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_JobAid_OPERALogin_AfterReboot.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

 	<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-3f (Heap Size validation)</td>
      <td>$($c3['f'])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_JobAid_OPERALogin_AfterReboot.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

 	<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-3g (Forms.conf File Validation)</td>
      <td>$($c3['g'])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_JobAid_OPERALogin_AfterReboot.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DC03D77D289FF128BA17D562C89A3D582D0E4B374708?allowInterrupt=1"">Link</a></td>
    </tr>

   
    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-5 (ODC,EVO Activation Settings)</td>
      <td class=grey>$C5_ODC</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_ODC Export.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/fileview/DC9B346216161B88C0D5EFFF1B7F983A10991FF469CA"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-6 (Scrren Painter Opportunity-ID)</td>
      <td>$(match-withObject $Screen_PO ""UDFC14"")</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_GoBackDocument_BlockScreenpainting_OpportunityID.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DFF2A5D4F0991EB4DB84655F2F3B3CBE37D4F24E2318?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-7a (OXI Comm Method and ERS Rate Opera)</td>
      <td>$((HTML_Multi_validation)[0])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_GoBackDocument_OXI Comm Method.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DF64AF8879D95A4CE8A566DDCDCF5ECF463F8F5A9A9F?allowInterrupt=1"">Link</a></td>
    </tr>

     <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-7b (Update Business Event Configuration)</td>
      <td class=grey>Please Check Manually</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_GoBackDocument_OXI Comm Method.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DF64AF8879D95A4CE8A566DDCDCF5ECF463F8F5A9A9F?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-7c (S&C Owner Override)</td>
      <td class=grey>$C7c</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_GoBackDocument_OXI Comm Method.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DF64AF8879D95A4CE8A566DDCDCF5ECF463F8F5A9A9F?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-8 (Block Rolling End Date)</td>
      <td>$((HTML_Multi_validation)[1])</td>
      <td><a href=""$results\Block rolling End date.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp -</td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-9 (KB Article for Error INT-90002,396 / Invalid or Missing Name) </td>
     <td class=green>VALID <span class=grey>- Will be shared with property as a part of activity completion email</span></td>
    <td></a><a href=""$results\C9_KB_Article.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://marriott.service-now.com/msp?class=kb_article&sys_class=22c3e1b31b72e4d017e1ece66e4bcb5d"">Link</td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-10 (Report Analyzer)</td>
      <td>$(($Report_analyzer[1] -match ""Opera Analyzer"") -and ($(check-files) -match ""VALID""))</td>
    <td><a href=""$results\MARRIOTT_OPERA_5.6_GoBackDocument_OPERA_Analyzer_Bundle.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D92E1C56D707A71681B84154864105A315641E01CB8C?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-11 (MIMPG XML Version)</td>
      <td>$((HTML_Multi_validation)[2])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_JobAid_MIMPG_InventoryXML.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D017A61AF286A560E44FF7101151E506F60FC6025979?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-12 Ritz Carlton Rewards</td>
      <td>$(Validate-C12)</td>
      <td><a href=""$results\C12_MARRIOTT_OPERA_5.6_JobAid_OXIConversion_RitzCarltonRewards.pdf"" target=""_explorer.exe"">Job Aid</a></a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D007CE20D915AB9E44AE28AC2F8359B47FBF747367EC?allowInterrupt=1"">Link</a></td>
    </tr>

     <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-13 MARSHA Proxy Rate Codes</td>
      <td>$((HTML_Multi_validation)[3])</td>
      <td><a href=""$results\MARRIOTT_OPERA_5.6_GoBackDocument_OXIMARSHA_ProxyRateCodes.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D88D1F5DA459D275EB972FC03EB55B80325DC6FDEE42?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-14 Business Block notes Configuration in Opera</td>
      <td>$(Validate-C14)</td>
      <td><a href=""$results\C14_286_operav56upgradebookingnoteinternalflag.pdf"" target=""_explorer.exe"">Job Aid</a></a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/D8E6A96DBA5F8EBB0D38AEF12D976EE13E53B4FF655D?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-15 Check Marsha OXI Global Parameters: WS Alive</td>
      <td>$(($c15_ws_alive[0] -match ""https://marshaprod.marriott.com:5617/reservationgpms"") -and ($c15_ws_alive[1] -match ""Y""))</td>
      <td><a href=""$results\C15_MARRIOTT_OPERA_5.6_JobAid_OXIMARSHA_GlobalParameters_WSALIVE.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents/embed/v2/link/LF7B341C48BC97CF9D1845E90FF2E6AA0C019C57056A/file/DACF663EAFF0F770B8D4598A655FBD8E88029E0EA585?allowInterrupt=1"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-16 Crisis Export Decommission</td>
      <td>$((HTML_Multi_validation)[4])</td>
      <td><a href=""$results\C16-opera56crisisexportdecomission.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-17 Stay Export validation</td>
      <td>$(match-withObject $c17_file  "wmbprd.marriott.com")</td>
      <td><a href=""$results\C17_Opera56stayexportconfiguration.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-18 Table Space validation </td>
      <td>$((validate-c18_tablespace)[0])</td>
      <td><a href=""$results\C18_Tablespace verification jobaid.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-19 Discovery Export validation</td>
      <td>$(Validate-C19)</td>
      <td><a href=""$results\C18_Tablespace verification jobaid.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>
	
	 <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>C-20 SYSAUX validation</td>
      <td>$(C20-SYSAUXValidation)</td>
      <td><a href=""$results\C20-SYSAUX Tablespace Grows Rapidly After Upgrading Database.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>

    <tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>OPI Validation</td>
      <td>$(if($opi_versions[0] -eq "Not Applicable"){"Not Applicable"}elseif([string]::IsNullOrWhiteSpace($opi_versions[0])){"MISSING"}else{"VALID"})</td>
      <td>Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>
	
	
<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>check DB log4j</td>
      <td>$(Validate-log4j)</td>
      <td><a href=""$results\Log4j-Remediation-Post-19C-Database-Upgrade.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
</tr>

<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Check Java Backup</td>
      <td>$(Validate-JavaBackup)</td>
      <td>Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
</tr>


<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>log4j evidence</td>
      <td>$(log4j-evidence $Weblogic_server)</td>
      <td>Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
</tr>
<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Check httpd</td>
      <td>$(Validate-httpd)</td>
      <td>Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
</tr>
<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>ACL</td>
      <td>$(acl)</td>
      <td><a href=""$results\ACL.pdf"" target=""_explorer.exe"">Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
    </tr>

<tr>
      <th scope=""row"">$($row_count++)$row_count</th>
      <td>Training User</td>
      <td>$(Training_user)</td>
      <td>Job Aid</a>&nbsp&nbsp&nbsp&nbsp<a href=""https://securesites.oracle.com/documents"">Link</a></td>
</tr>


  </tbody>
</table>
<footer>
<br>
<br> 
This report is generated by LITMUS (A PMS compliance tracking tool), developed by ITC-Infotech on behalf of Marriott International
<br> 
Version: $script_release  
</footer>
</div>
</body>
</html>" | Out-File "d:\Patching_$dat\OPERA_Compliance_Status.html" -Append -Encoding ascii
(Get-Content -Path "d:\Patching_$dat\OPERA_Compliance_Status.html" -raw) -replace "<td>True","<td class=green>VALID" -replace "<td>VALID","<td class=green>VALID" -replace "<td>False","<td class=red>MISSING" -replace "<td>MISSING","<td class=red>MISSING" | Set-Content "d:\Patching_$dat\OPERA_Compliance_Status.html"
(Get-Content -Path "d:\Patching_$dat\OPERA_Compliance_Status.html" -raw) -replace "<td>Conflict : Property Not Found In ACTIVE ERS Database, But ERS Rate Codes Are Configured. Please Check For Updated List","<td class=red>Conflict <span class=grey>: Property Not Found In ACTIVE ERS Database,<br>But ERS Rate Codes Are Configured. Please Check For Updated List</span>" | Set-Content "d:\Patching_$dat\OPERA_Compliance_Status.html"
(Get-Content -Path "d:\Patching_$dat\OPERA_Compliance_Status.html" -raw) -replace "<td>Conflict for","<td class=red>Conflict for" | Set-Content "d:\Patching_$dat\OPERA_Compliance_Status.html"
(Get-Content -Path "d:\Patching_$dat\OPERA_Compliance_Status.html" -raw) -replace "<td>Not Applicable","<td class=grey>Not Applicable" | Set-Content "d:\Patching_$dat\OPERA_Compliance_Status.html"
#Completion status calculation
$Compli=(Get-Content -Path "d:\Patching_$dat\OPERA_Compliance_Status.html" -raw)
$Valid_count=$Compli -split ">VALID<" | Measure-Object | Select-Object -ExpandProperty count
$total_checks_count=50
if($btr_value[1] -eq "Not Applicable"){$total_checks_count--}
if((Check-Hotfix) -eq "Not Applicable"){$total_checks_count--}
if($opi_versions[0] -eq "Not Applicable"){$total_checks_count--}
if((log4j-evidence $Weblogic_server) -eq "Not Applicable"){$total_checks_count--}
$Valid_count=($Valid_count*100)/$total_checks_count
$Valid_count=[math]::round($Valid_count,0)
if($Valid_count -ge 70 -and $Valid_count -le 90){$percent="<h5 style=""color:gold;"">Completion Status: $Valid_count%</h5>"}
if($Valid_count -gt 90){$percent="<h5 class=green>Completion Status: $Valid_count%</h5>"}
if($Valid_count -lt 70){$percent="<h5 class=red>Completion Status: $Valid_count%</h5>"}
(Get-Content -Path "d:\Patching_$dat\OPERA_Compliance_Status.html" -raw) -replace "<h5 class=pink>Completion Status: NA</h5>","$percent" | Set-Content "d:\Patching_$dat\OPERA_Compliance_Status.html"
#End of HTML Creation

#Multi-property addition
if($P_count -gt 1)
    {
        for($i=1;$i -lt $P_count;$i++)
            {
                $Marsha_excel=Get-Content -Path "$results\marsha.txt" | select -Index $i

                $Block_rolling=Get-Content -Path "$results\Block rolling date.txt" | select -index $i

                $property=Get-Content -Path "$results\property_name.txt" | select -Index $i

                $mimpg_xml=Get-Content -Path "$results\mimpg_xml_version.txt" | select -Index $i

                $Country_0_city_1=get-country -P_count $i

                $Oxi_comm=Get-Content -Path "$results\OXI Comm Method_ERS Rate Opera.txt" | select -index $i

                $c16_value=Get-Content "$results\C16_disable_crisis_export.txt"
                $c16_validation=validate-c16 $Marsha_excel

                if($Marsha_excel -eq $marsha){$property_type="Multi-Property"}else{$property_type="Part of Multi-Property"}

                "$Marsha_excel;$property;$ownership;$region;$($Country_0_city_1[0]);$($Country_0_city_1[1]);$Activity;$Report_date;$property_type;$marsha; ;$hostname;$ip_address;$Total_RAM;$cpu_details;$total_HDD;$role;$Version;$($btr_value[0]);$($btr_value[1]);$($schema_version[0]);$($schema_version[1]);$($schema_version[2]);$runtime_version;$NInstall_Date;$Opera_size;$Oxi_size;$Opera_OIW_size;$($opera_tools['oapp']);$($opera_tools['smt']);$($opera_tools['oxi']);$Ear_version;$Weblogic_server;$WL_version $JDK_version, $WL_lspatches;$(check-Weblogicpatches) $Weblogic_release;$(check-WLOpatch $WL_version);$(check-Weblogicclient);$DB_version ,$DB_lspatches;$Db_patches_call $DB_Release;$($db_Opatch_check);$Archive_log;$Archive_log_Validation;$Hotfix;$(check-hotfix);$(Get-Content "$results\applied_freedomfixes.txt");$freedom_fix_Validation;$audit_policy;$(match-withObject $audit_policy "Null");NA;$NLS_lang;NA;$c2_Validation;$($Auto_start_Delay.AutoStartDelay) $($Delayed_auto_start.DelayedAutostart);$C3_Validation_a;$($c3['b_v']);$($c3['b']);$($c3['c_v']);$($c3['c']);$($c3['d_v']);$($c3['d']);$($c3['e_v']);$($c3['e']);$($c3['f_v']);$($c3['f']);$($c3['g_v']);$($c3['g']);$C5_ODC;$Screen_PO;$(match-withObject $Screen_PO ""UDFC14"");$Oxi_comm;$($Oxi_comm -match 'https://marshaprod.marriott.com:5617/reservationgpms');Manual Check;$C7c;$Block_rolling;$($Block_rolling -match "715");Manual Check;$Report_analyzer;$(check-files);$mimpg_xml;$($mimpg_xml -match "V3");$(Validate-C12);$(Get-Content -Path "$results\c13_rate_codes.txt" | select -Index $i);$(check-Rate_Codes -P_count $i);$(Validate-C14);$c15_ws_alive;$(($c15_ws_alive[0] -match "https://marshaprod.marriott.com:5617/reservationgpms") -and ($c15_ws_alive[1] -match "Y"));$c16_value;$c16_validation;$c17_file;$($c17_file -match "wmbprd.marriott.com");$((validate-c18_tablespace)[1]);$((validate-c18_tablespace)[0]);$(Validate-C19);$($opi_versions[0]);$($opi_versions[1]);$ifc_active_names;$ifc_machine_count;$(C20-SYSAUXValidation);$(Training_user);$IFC_version" | Out-File "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv" -Append -Encoding ascii
                (Get-Content -Path "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv" -raw) -replace "True","VALID" -replace "False","MISSING" | Set-Content "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv"
                
                $csv_input = Get-Content "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv"
                $csv_data = ConvertFrom-Csv -Delimiter ";" -InputObject $csv_input
                $csv_data[$i]
                "`n ========================================================================================================================================================================================================= `n" | Out-File "$logs\Simple_$marsha`_Compliance-report.txt" -append
                $csv_data[$i] | Out-File "$logs\Simple_$marsha`_Compliance-report.txt" -append

                
                $jason_output="{"
                $csv_data[$i].PSObject.Properties | ForEach-Object {
	            $jason_output=$jason_output+"""$($_.Name)"":""$($_.Value)"","
                }

                $jason_output=$jason_output+"}"
                $jason_output=$jason_output -replace ",}","}" -replace '\\','\\'
                $jason_output | out-file "$logs\$Marsha_excel`_Compliance-report_$dat.json" -Encoding ascii
            }
    }
#END of Multi-property addition
Copy-Item -Path "d:\Patching_$dat\$marsha`_Compliance-report_$dat.csv" -Destination "$logs\$marsha`_Compliance-report_$dat.csv" -Recurse
Write-Host "Script executed Please Find the created $marsha`_Compliance-report_$dat.csv file in D drive Patching_$dat folder`n `n" -ForegroundColor Green
Copy-Item -Path "d:\Patching_$dat\OPERA_Compliance_Status.html" -Destination "$logs\OPERA_Compliance_Status.html"

Copy-Item -Path "d:\Patching_$dat\OPERA_Compliance_Status.html" -Destination "$logs\Compliance_Report_$marsha.html"
(Get-Content -Path "$logs\OPERA_Compliance_Status.html" -raw) -replace "<table class=""table table-hover"">","<table class=""customers"">" | Set-Content "$logs\OPERA_Compliance_Status.html"

Copy-Item -Path "$logs\Simple_$marsha`_Compliance-report.txt" -Destination "d:\Patching_$dat\Simple_$marsha`_Compliance-report.txt"

ren "d:\Patching_$dat\OPERA_Compliance_Status.html" "$marsha`_OPERA_Compliance_Status_$time_2.html"
ii "d:\Patching_$dat\$marsha`_OPERA_Compliance_Status_$time_2.html"
Stop-Transcript
