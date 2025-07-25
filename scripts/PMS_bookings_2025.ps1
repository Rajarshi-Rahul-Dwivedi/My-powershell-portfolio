#Hello, This Script will not function in it's current form please read the discription or feel free to contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281
Write-Host "Please wait executing Change Tracker" -BackgroundColor DarkRed
$change_tracker="$PSScriptRoot\Meta Data\PMS_change_tracker.ps1"
.$change_tracker
Write-Host "Script Version is 8.9 - release 30 December 2024" -BackgroundColor Blue
start-sleep -Seconds 3
#$global:loc="D:\Bookings Tool\2024\PMS"
$loc=$PSScriptRoot
Start-Transcript -Path "$loc\fetch_booking_data_log.txt"

function Get-ValidEmails {
    param (
        [string]$InputString
    )

    # Define a regular expression pattern to match valid email addresses
    $EmailPattern = "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"

    # Use the Select-String cmdlet to find email addresses in the input string
    $Matches = $InputString | Select-String -Pattern $EmailPattern -AllMatches

    # Extract the matched email addresses and store them in an array
    $ValidEmails = @()
    foreach ($Match in $Matches.Matches) {
    try{$null=[mailaddress]$Match.Value}catch{write-warning "$($Match.Value) Not a valid Email";continue}
        $ValidEmails += $Match.Value
    }

    $ValidEmails=$ValidEmails -join ';'
    return $ValidEmails
}

#try and catch block to intialize all meta_data
try{
$dat=get-date -UFormat "%M-%d-%b"
$files=(Get-ChildItem "$loc\*opera*\*.tsv" | select fullname).fullname
$header="Date Time	Customer Name	Customer Email	Customer Phone	Customer Address	Staff	Staff Name	Staff Email	Service	Location	Duration (mins.)	Pricing Type	Price	Currency	Cc Attendees	Signed Up Attendees Count	Text Notifications Enabled	 Custom Fields	Event Type	Booking Id	Tracking Data"
Get-Content $files | Set-Content "$loc\temp_BookingsReportingData.tsv"
Get-Content "$loc\temp_BookingsReportingData.tsv" | Where-Object {$_ -notmatch 'Date Time	Customer Name'} | Set-Content "$loc\temp2_BookingsReportingData.tsv"
$header | Set-Content "$loc\header.tsv"
Get-Content "$loc\header.tsv","$loc\temp2_BookingsReportingData.tsv" | Set-Content "$loc\BookingsReportingData.tsv"
Remove-Item "$loc\temp2_BookingsReportingData.tsv","$loc\temp_BookingsReportingData.tsv","$loc\header.tsv"

$csv_input=Get-Content "$loc\BookingsReportingData.tsv"
$csv_input=$csv_input -replace "Host MARSHA \(Required\)" , 'MARSHA Code (Required)' -replace "Preferred Start Time for Activity \(If don't want to start at 11:00 PM Property Local Time\) \(optional\)",'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time) '
$csv_input=$csv_input -replace  "Preferred Start Time for Activity \(If don't want to start at 11:00 PM Property Local Time\) \(Note: If you are booking for Sunday, the earliest slot available is 7:00 PM Property Time\) " , 'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
$csv_input=$csv_input -replace  "Preferred Start Time for Activity \(If don't want to start at 11:00 PM Property Local Time\) \(Note: If you are booking for Sunday, the earliest slot available is 7:00 PM Property Time\)" , 'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
$csv_input=$csv_input -replace "Property Contact Email \(Required\)","Property Contact Email" -replace "Property Contact Phone Number\(Required\)","Property Contact Phone Number"
$csv_data = ConvertFrom-Csv -Delimiter "`t" -InputObject $csv_input

$meta_input=Get-Content "$loc\Meta data\PMS_meta_data.csv"
$meta_data=ConvertFrom-Csv -Delimiter "," -InputObject $meta_input

$smartsheet_input=Get-Content "$loc\Meta Data\PMS_smartsheet.csv"
$smartsheet_data=ConvertFrom-Csv -Delimiter "," -InputObject $smartsheet_input

$IATA_city_input=Get-Content "$loc\Meta Data\IATA_CIty.csv"
$IATA_city=ConvertFrom-Csv -Delimiter "," -InputObject $IATA_city_input

$airport_code=Import-Csv -Path "$loc\Meta Data\airports.csv"

$booked_date_csv=ConvertFrom-Csv -InputObject (gc "$loc\Meta Data\Booked_date.csv") -Delimiter ','

}catch{Write-Warning "Meta data initialization failed `n $_";pause;exit}

#---------------------------------------------------Setting variables
$booked_date=@()
$booked_date+=$booked_date_csv
$table=@()

$No_metaData=@()
$loopcounter=0
[decimal]$total_loop=($csv_data.count)*4+($meta_data.count)*5
$meta_data_d=@{}
$incorrect_time=@{}

#Setting dictonary with marsha as key and row as value
foreach ($row in $meta_data) {
    $meta_data_d[$row."Marsha"] = $row
}


#Q1

foreach($element in $csv_data)
{
    $loopcounter++
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow
    $skip_duplicate_check=$false
    if($element.'Service' -match 'q1')
    {
        $temp=$element.'Custom Fields' | ConvertFrom-Json
        $Marsha=$temp.'MARSHA Code (Required)'
        if([String]::IsNullOrWhiteSpace($marsha)){$marsha=$temp.'MARSHA Code (Required)'}
        if($Marsha -match "^[A-Z]{5}"){$Marsha=($Marsha | Select-String "^[A-Z]{5}" | select -ExpandProperty matches).value}
        
        $marsha=$Marsha.ToUpper()
        $marsha=$marsha.Trim()
        $find_booking_date=$booked_date | Where-Object {$_.Marsha -eq $marsha -and ![string]::IsNullOrEmpty($_.'Q1 Booking date')}
        
        $add_marsha=$temp.'Additional MARSHA Codes, if multi-property configuration'
        if([string]::IsNullOrWhiteSpace($add_marsha)){$add_marsha=$temp."Additional MARSHA Codes"}
        if($add_marsha -eq $Marsha){$add_marsha=$null}
        $date=$element.'Date Time'
        $date=[datetime]$date
        $q_date=$date.ToString("MM\/dd\/yyyy")
        
        #setting up booked_on date with schedule date in booked_date.csv

        if(!($find_booking_date.'Q1 Schedule' -match $q_date))
        {
            
            $booked_date+=[pscustomobject]@{'Marsha'=$Marsha;'Q1 Schedule'=$q_date;'Q1 Booking date'="$((get-date).ToString("MM\/dd\/yyyy"))";}
            $skip_duplicate_check=$true

            $table_object_ref=$table | Where-Object {$_.'MARSHA' -match $marsha}
            if($table_object_ref -ne $null){
            $table_object_ref.psobject.properties.name | foreach{$table_object_ref.psobject.properties.remove("$_")}}

        }

        #to avoid duplicate entry
        if($table.'marsha' -match $Marsha -and $skip_duplicate_check -eq $false)
            {
                #1.check if current element is on latest booked_on date
                #2.if it is remove previous entry of same marsha
                #3.else skip this element
                $is_latest_flag=$true
                $current_element_booked_on=$find_booking_date | Where-Object {$_.Marsha -eq $marsha -and $_.'Q1 Schedule' -eq $q_date}
                $current_element_booked_on=[datetime]$current_element_booked_on.'Q1 Booking date'
                foreach($row in $find_booking_date){
                    if($current_element_booked_on -lt [datetime]$row.'Q1 Booking date'){$is_latest_flag=$false}
                }

                 if($is_latest_flag -eq $true){
                 $table_object_ref=$table | Where-Object {$_.'MARSHA' -match $marsha}
                 $table_object_ref.psobject.properties.name | foreach{$table_object_ref.psobject.properties.remove("$_")} 
                 }else{continue}
        
             }
        
        
        $header_poc=$element.'Customer Name'
        $header_email=$element.'Customer Email'
        $header_phone=$element.'Customer Phone'


        $Body_Poc=$temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'
        
        $Body_Phone=$temp.'Property Contact Phone Number '
        $Body_email=$temp.'Property Contact Email'
        $q_Time=$temp.'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
        $original_time=$q_time
        
        if(!([String]::IsNullOrWhiteSpace($q_Time) -or $q_Time -eq 'null'))
        {
            if($q_Time -match '(?i)\d+[:.]*\d*\s*(?:AM|PM)?')
                {
                    if($q_Time -match'\d+.\d*\s*[ap]*[m]*'){$q_Time=$q_Time -replace '(?<=\d)\.(?=\d)',':' -replace '\.',''}
                    if($q_Time -match'(?i)\d+\s?:\s?\d+\s*(?:AM|PM)?'){$q_Time=$q_Time -replace '(?<=\d)\s+(?=:)','' -replace '(?<=:)\s+(?=\d)','' -replace '(?<=\d)\s*(?=A|P)',' '}
                    $q_Time=($q_Time | Select-String '(?i)\d+:*\d*\s*(?:AM|PM)?' | select -ExpandProperty matches).value
                    try{$q_time=[datetime]$q_time
                    $q_Time=$q_Time.ToString("hh:mm tt")}catch{$incorrect_time[$marsha]="Q1 : "+$original_time;$q_Time="11:00 PM"} 
                }else{$incorrect_time[$marsha]="Q1 : "+$original_time;$q_Time="11:00 PM"}
            }else{$q_Time="11:00 PM"}

        $Flag=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)? (Required)'
        if([string]::IsNullOrWhiteSpace($flag)){$flag=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)?'}
#querying meta data
                $meta_object=$meta_data_d[$Marsha]
                if(!($meta_data.'marsha' -match $marsha)){$No_metaData+=$Marsha}


                $Table+=[pscustomobject]@{'Marsha'=$Marsha;'Property Name'=$meta_object.'Property Name';'Q1 Schedule'=$q_Date+' '+$q_Time;'Q2 Schedule'=$null;'Q3 Schedule'=$null;'Q4 Schedule'=$null;
                'Q1 Upgrade Date'=$q_Date;'Q1 Upgrade Time'=$q_Time;'Q1 OPERA PMS Fiscal or Legal Requirements'=$Flag;
                'Q2 Upgrade Date'=$null;'Q2 Upgrade Time'=$null;'Q2 OPERA PMS fiscal or legal requirements'=$null;
                'Q3 Upgrade Date'=$null;'Q3 Upgrade Time'=$null;'Q3 OPERA PMS fiscal or legal requirements'=$null;
                'Q4 Upgrade Date'=$null;'Q4 Upgrade Time'=$null;'Q4 OPERA PMS fiscal or legal requirements'=$null;
       
                'Ownership'=$meta_object.'Ownership';'Property Type'=$null;'Multi-Property Marsha'=$add_marsha;
                'Continent'=$meta_object.'continent';'Country'=$meta_object.'country';'Time Zone'=$meta_object.'Time Zone';
                'GM'=$meta_object.'GM';'GM Email'=$meta_object.'GM Email';'POC Name'=$meta_object.'POC Name';'POC Email'=$meta_object.'POC Email';'POC Phone'=$meta_object.'POC Phone';'IT Email'=$meta_object.'IT Email';'IT Phone'=$meta_object.'IT Phone';
                'POC Name (Booking)'=$header_poc;'Email ID (Booking)'=$header_email;'Ph No (Booking)'=$header_phone;
                'POC Name (Details)'=$Body_Poc;'Email ID (Details)'=$Body_email;'Ph No (Details)'=$Body_Phone;
                'On MI Network'=$meta_object.'On MI Network';'Hosting'=$meta_object.'Hosting'}
    }
    
}

#q2
foreach($element in $csv_data)
{
    $loopcounter++
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow
    $skip_duplicate_check=$false
    if($element.'Service' -match 'Q2')
    {
        $table_object=$null
        $temp=$element.'Custom Fields' | ConvertFrom-Json
        $Marsha=$temp.'MARSHA Code (Required)'
        if([String]::IsNullOrWhiteSpace($marsha)){$marsha=$temp.'MARSHA Code (Required)'}
        if($Marsha -match "^[A-Z]{5}"){$Marsha=($Marsha | Select-String "^[A-Z]{5}" | select -ExpandProperty matches).value}
        $marsha=$Marsha.ToUpper()
        $marsha=$marsha.Trim()
        $date=$element.'Date Time'
        $date=[datetime]$date
        $q_date=$date.ToString("MM\/dd\/yyyy")

        $find_booking_date=$booked_date | Where-Object {$_.Marsha -eq $marsha -and ![string]::IsNullOrEmpty($_.'Q2 Booking date')}
        #setting up booked_on date with schedule date in booked_date.csv

        if(!($find_booking_date.'Q2 Schedule' -match $q_date))
        {
            
            $booked_date+=[pscustomobject]@{'Marsha'=$Marsha;'Q2 Schedule'=$q_date;'Q2 Booking date'="$((get-date).ToString("MM\/dd\/yyyy"))";}
            $skip_duplicate_check=$true
        }
        #check if marsha is present in table
        if($table.'MARSHA' -match $Marsha)
        {
            #Write-Host "Marsha matched $marsha Mathes : $($matches)"
            $table_object=$table | Where-Object {$_.'MARSHA' -match $Marsha}
            $add_marsha=$temp.'Additional MARSHA Codes, if multi-property configuration'
            if([string]::IsNullOrWhiteSpace($add_marsha)){$add_marsha=$temp."Additional MARSHA Codes"}
            if($add_marsha -eq $Marsha){$add_marsha=$null}
            
            #handle duplicate entry
        if(!([String]::IsNullOrWhiteSpace($table_object.'Q2 Schedule')))
            {
                 
        #to avoid duplicate entry
        if($table.'marsha' -match $Marsha -and $skip_duplicate_check -eq $false)
            {
                #1.check if current element is on latest booked_on date
                #2.if it is remove previous entry of same marsha
                #3.else skip this element
                $is_latest_flag=$true
                $current_element_booked_on=$find_booking_date | Where-Object {$_.Marsha -eq $marsha -and $_.'Q2 Schedule' -eq $q_date}
                $current_element_booked_on=[datetime]$current_element_booked_on.'Q2 Booking date'
                foreach($row in $find_booking_date){
                    if($current_element_booked_on -lt [datetime]$row.'Q2 Booking date'){$is_latest_flag=$false}
                }

                 if($is_latest_flag -eq $true){
                 $table_object=$table | Where-Object {$_.'MARSHA' -match $marsha}
                 $table_object.'Q2 Schedule'=$null
                 $table_object.'Q2 Upgrade Date'=$null
                 $table_object.'Q2 Upgrade Time'=$null
                 $table_object.'Q2 OPERA PMS fiscal or legal requirements'=$null
                 }else{continue}
        
            }       
    }
            
            $table_object.'Q2 Upgrade Date'=$q_date
            try{$table_object.'Q2 Upgrade Date'=$q_date}catch{pause}
            $q_Time=$temp.'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
            $original_time=$q_time
            

            if(!([String]::IsNullOrWhiteSpace($q_Time) -or $q_Time -eq 'null'))
                {
                if($q_Time -match '(?i)\d+[:.]*\d*\s*(?:AM|PM)?')
                {
                    if($q_Time -match'\d+.\d*\s*[ap]*[m]*'){$q_Time=$q_Time -replace '(?<=\d)\.(?=\d)',':' -replace '\.',''}
                    if($q_Time -match'(?i)\d+\s?:\s?\d+\s*(?:AM|PM)?'){$q_Time=$q_Time -replace '(?<=\d)\s+(?=:)','' -replace '(?<=:)\s+(?=\d)','' -replace '(?<=\d)\s*(?=A|P)',' '}
                    $q_Time=($q_Time | Select-String '(?i)\d+:*\d*\s*(?:AM|PM)?' | select -ExpandProperty matches).value
                    try{$q_time=[datetime]$q_time
                    $q_Time=$q_Time.ToString("hh:mm tt")}catch{$incorrect_time[$marsha]+=" Q2 : "+$original_time;$q_Time="11:00 PM"} 
                }else{$incorrect_time[$marsha]+=" Q2 : "+$q_Time;$q_Time="11:00 PM"}
            }else{$q_Time="11:00 PM"}

            $table_object.'Q2 Schedule'=$q_date+' '+$q_Time
            $table_object.'Q2 Upgrade Time'=$q_Time
            $table_object.'Q2 OPERA PMS Fiscal or Legal Requirements'=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)? (Required)'
            if([string]::IsNullOrWhiteSpace($table_object.'Q2 OPERA PMS Fiscal or Legal Requirements')){$table_object.'Q2 OPERA PMS Fiscal or Legal Requirements'=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)?'}
            
            if($table_object.'Property Type' -eq "Single-Property"){$table_object.'Property Type'=$null}
            if($table_object.'Multi-Property Marsha' -notmatch $add_marsha -and !([String]::IsNullOrWhiteSpace($add_marsha))){$table_object.'Multi-Property Marsha'=$table_object.'Multi-Property Marsha'+" "+$add_marsha}
            
            if(![String]::IsNullOrWhiteSpace($element.'Customer Name')){$table_object.'POC Name (Booking)'=$element.'Customer Name'}
            if(![String]::IsNullOrWhiteSpace($element.'Customer Email')){$table_object.'Email ID (Booking)'=$element.'Customer Email'}
            if($table_object.'Ph No (Booking)' -notmatch $element.'Customer Phone'){$table_object.'Ph No (Booking)'=$table_object.'Ph No (Booking)'+' ; '+$element.'Customer Phone'}
          

            if(($table_object.'POC Name (Details)' -notmatch $temp.'Full Name of Property Contact During Upgrade/Patching, if different than above') -and !([String]::IsNullOrWhiteSpace($temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'))){$table_object.'POC Name (Details)'=$table_object.'POC Name (Details)'+' ; '+$temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'}
            if(($table_object.'Email ID (Details)' -notmatch [Regex]::Escape($temp.'Property Contact Email')) -and !([String]::IsNullOrWhiteSpace($temp.'Property Contact Email'))){$table_object.'Email ID (Details)'=$table_object.'Email ID (Details)'+' ; '+$temp.'Property Contact Email'}
            if(($table_object.'Ph No (Details)' -notmatch [Regex]::Escape($temp.'Property Contact Phone Number ')) -and !([String]::IsNullOrWhiteSpace($temp.'Property Contact Phone Number '))){$table_object.'Ph No (Details)'=$table_object.'Ph No (Details)'+' ; '+$temp.'Property Contact Phone Number '}
        }

        #creating new entry for only q2
        if(!($table.'marsha' -match $Marsha))
        {
        $date=$element.'Date Time'
        $date=[datetime]$date
        $q_date=$date.ToString("MM\/dd\/yyyy")
        $header_poc=$element.'Customer Name'
        $header_email=$element.'Customer Email'
        $header_phone=$element.'Customer Phone'

        $add_marsha=$temp.'Additional MARSHA Codes, if multi-property configuration'
        if([string]::IsNullOrWhiteSpace($add_marsha)){$add_marsha=$temp."Additional MARSHA Codes"}
        if($add_marsha -eq $Marsha){$add_marsha=$null}
            


        $Body_Poc=$temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'
        
       
        
        $Body_Phone=$temp.'Property Contact Phone Number '
        $Body_email=$temp.'Property Contact Email'
         $q_Time=$temp.'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
         $original_time=$q_time
         
        if(!([String]::IsNullOrWhiteSpace($q_Time) -or $q_Time -eq 'null'))
                {
                if($q_Time -match '(?i)\d+[:.]*\d*\s*(?:AM|PM)?')
                {
                    if($q_Time -match'\d+.\d*\s*[ap]*[m]*'){$q_Time=$q_Time -replace '(?<=\d)\.(?=\d)',':' -replace '\.',''}
                    if($q_Time -match'(?i)\d+\s?:\s?\d+\s*(?:AM|PM)?'){$q_Time=$q_Time -replace '(?<=\d)\s+(?=:)','' -replace '(?<=:)\s+(?=\d)','' -replace '(?<=\d)\s*(?=A|P)',' '}
                    $q_Time=($q_Time | Select-String '(?i)\d+:*\d*\s*(?:AM|PM)?' | select -ExpandProperty matches).value
                    try{$q_time=[datetime]$q_time
                    $q_Time=$q_Time.ToString("hh:mm tt")}catch{$incorrect_time[$marsha]+=" Q2 : "+$original_time;$q_Time="11:00 PM"} 
                }else{$incorrect_time[$marsha]+=" Q2 : "+$q_Time;$q_Time="11:00 PM"}
            }else{$q_Time="11:00 PM"}

        $Flag=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)? (Required)'
        if([string]::IsNullOrWhiteSpace($flag)){$flag=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)?'}
        #querying meta data
                $meta_object=$meta_data_d[$Marsha]
      
         $Table+=[pscustomobject]@{'Marsha'=$Marsha;'Property Name'=$meta_object.'Property Name';'Q1 Schedule'=$null;'Q2 Schedule'=$q_Date+' '+$q_Time;'Q3 Schedule'=$null;'Q4 Schedule'=$null;
                'Q1 Upgrade Date'=$null;'Q1 Upgrade Time'=$null;'Q1 OPERA PMS Fiscal or Legal Requirements'=$null;
                'Q2 Upgrade Date'=$q_Date;'Q2 Upgrade Time'=$q_Time;'Q2 OPERA PMS fiscal or legal requirements'=$Flag;
                'Q3 Upgrade Date'=$null;'Q3 Upgrade Time'=$null;'Q3 OPERA PMS fiscal or legal requirements'=$null;
                'Q4 Upgrade Date'=$null;'Q4 Upgrade Time'=$null;'Q4 OPERA PMS fiscal or legal requirements'=$null;
       
                'Ownership'=$meta_object.'Ownership';'Property Type'=$null;'Multi-Property Marsha'=$add_marsha;
                'Continent'=$meta_object.'continent';'Country'=$meta_object.'country';'Time Zone'=$meta_object.'Time Zone';
                'GM'=$meta_object.'GM';'GM Email'=$meta_object.'GM Email';'POC Name'=$meta_object.'POC Name';'POC Email'=$meta_object.'POC Email';'POC Phone'=$meta_object.'POC Phone';'IT Email'=$meta_object.'IT Email';'IT Phone'=$meta_object.'IT Phone';
                'POC Name (Booking)'=$header_poc;'Email ID (Booking)'=$header_email;'Ph No (Booking)'=$header_phone;
                'POC Name (Details)'=$Body_Poc;'Email ID (Details)'=$Body_email;'Ph No (Details)'=$Body_Phone;
                'On MI Network'=$meta_object.'On MI Network';'Hosting'=$meta_object.'Hosting'}
        
        }
        if(!($meta_data.'marsha' -match $marsha)){$No_metaData+=$Marsha}
    }

}
#end of q2

#Q3
foreach($element in $csv_data)
{
    $loopcounter++
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow
    $skip_duplicate_check=$false
    if($element.'Service' -match 'Q3')
    {
        $temp=$element.'Custom Fields' | ConvertFrom-Json
        $Marsha=$temp.'MARSHA Code (Required)'
        if([String]::IsNullOrWhiteSpace($marsha)){$marsha=$temp.'MARSHA Code (Required)'}
        if($Marsha -match "^[A-Z]{5}"){$Marsha=($Marsha | Select-String "^[A-Z]{5}" | select -ExpandProperty matches).value}
        $marsha=$Marsha.ToUpper()
        $marsha=$marsha.Trim()
        $date=$element.'Date Time'
        $date=[datetime]$date
        $q_date=$date.ToString("MM\/dd\/yyyy")

        $find_booking_date=$booked_date | Where-Object {$_.Marsha -eq $marsha -and ![string]::IsNullOrEmpty($_.'Q3 Booking date')}
        #setting up booked_on date with schedule date in booked_date.csv

        if(!($find_booking_date.'Q3 Schedule' -match $q_date))
        {
            
            $booked_date+=[pscustomobject]@{'Marsha'=$Marsha;'Q3 Schedule'=$q_date;'Q3 Booking date'="$((get-date).ToString("MM\/dd\/yyyy"))";}
            $skip_duplicate_check=$true
        }
        #check if marsha is present in table
        if($table.'marsha' -match $Marsha)
        {
            $table_object=$table | Where-Object {$_.'MARSHA' -match $marsha}
            
            $add_marsha=$temp.'Additional MARSHA Codes, if multi-property configuration'
            if([string]::IsNullOrWhiteSpace($add_marsha)){$add_marsha=$temp."Additional MARSHA Codes"}
            if($add_marsha -eq $Marsha){$add_marsha=$null}
            
            #handle duplicate entry
        if(!([String]::IsNullOrWhiteSpace($table_object.'Q3 Schedule')))
            {
                 
        #to avoid duplicate entry
        if($table.'marsha' -match $Marsha -and $skip_duplicate_check -eq $false)
            {
                #1.check if current element is on latest booked_on date
                #2.if it is remove previous entry of same marsha
                #3.else skip this element
                $is_latest_flag=$true
                $current_element_booked_on=$find_booking_date | Where-Object {$_.Marsha -eq $marsha -and $_.'Q3 Schedule' -eq $q_date}
                $current_element_booked_on=[datetime]$current_element_booked_on.'Q3 Booking date'
                foreach($row in $find_booking_date){
                    if($current_element_booked_on -lt [datetime]$row.'Q3 Booking date'){$is_latest_flag=$false}
                }

                 if($is_latest_flag -eq $true){
                 $table_object=$table | Where-Object {$_.'MARSHA' -match $marsha}
                 $table_object.'Q3 Schedule'=$null
                 $table_object.'Q3 Upgrade Date'=$null
                 $table_object.'Q3 Upgrade Time'=$null
                 $table_object.'Q3 OPERA PMS fiscal or legal requirements'=$null 
                 }else{continue}
        
            }       
    }
            $table_object.'Q3 Upgrade Date'=$q_date
            
            $q_Time=$temp.'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
            $original_time=$q_time
            

            if(!([String]::IsNullOrWhiteSpace($q_Time) -or $q_Time -eq 'null'))
                {
                if($q_Time -match '(?i)\d+[:.]*\d*\s*(?:AM|PM)?')
                {
                    if($q_Time -match'\d+.\d*\s*[ap]*[m]*'){$q_Time=$q_Time -replace '(?<=\d)\.(?=\d)',':' -replace '\.',''}
                    if($q_Time -match'(?i)\d+\s?:\s?\d+\s*(?:AM|PM)?'){$q_Time=$q_Time -replace '(?<=\d)\s+(?=:)','' -replace '(?<=:)\s+(?=\d)','' -replace '(?<=\d)\s*(?=A|P)',' '}
                    $q_Time=($q_Time | Select-String '(?i)\d+:*\d*\s*(?:AM|PM)?' | select -ExpandProperty matches).value
                    try{$q_time=[datetime]$q_time
                    $q_Time=$q_Time.ToString("hh:mm tt")}catch{$incorrect_time[$marsha]+=" Q3 : "+$original_time;$q_Time="11:00 PM"} 
                }else{$incorrect_time[$marsha]+=" Q3 : "+$q_Time;$q_Time="11:00 PM"}
            }else{$q_Time="11:00 PM"}

            $table_object.'Q3 Schedule'=$q_date+' '+$q_Time
            $table_object.'Q3 Upgrade Time'=$q_Time
            $table_object.'Q3 OPERA PMS Fiscal or Legal Requirements'=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)? (Required)'
            if([string]::IsNullOrWhiteSpace($table_object.'Q3 OPERA PMS Fiscal or Legal Requirements')){$table_object.'Q3 OPERA PMS Fiscal or Legal Requirements'=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)?'}
            
            if($table_object.'Property Type' -eq "Single-Property"){$table_object.'Property Type'=$null}
            if($table_object.'Multi-Property Marsha' -notmatch $add_marsha -and !([String]::IsNullOrWhiteSpace($add_marsha))){$table_object.'Multi-Property Marsha'=$table_object.'Multi-Property Marsha'+" "+$add_marsha}
            
            if([String]::IsNullOrWhiteSpace($table_object.'POC Name (Booking)')){$table_object.'POC Name (Booking)'=$element.'Customer Name'}
            if($table_object.'Email ID (Booking)' -notmatch $element.'Customer Email'){$table_object.'Email ID (Booking)'=$table_object.'Email ID (Booking)'+' ; '+$element.'Customer Email'}
            if($table_object.'Ph No (Booking)' -notmatch $element.'Customer Phone'){$table_object.'Ph No (Booking)'=$table_object.'Ph No (Booking)'+' ; '+$element.'Customer Phone'}
          

            if(($table_object.'POC Name (Details)' -notmatch $temp.'Full Name of Property Contact During Upgrade/Patching, if different than above') -and !([String]::IsNullOrWhiteSpace($temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'))){$table_object.'POC Name (Details)'=$table_object.'POC Name (Details)'+' ; '+$temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'}
            if(($table_object.'Email ID (Details)' -notmatch [Regex]::Escape($temp.'Property Contact Email')) -and !([String]::IsNullOrWhiteSpace($temp.'Property Contact Email'))){$table_object.'Email ID (Details)'=$table_object.'Email ID (Details)'+' ; '+$temp.'Property Contact Email'}
            if(($table_object.'Ph No (Details)' -notmatch [Regex]::Escape($temp.'Property Contact Phone Number ')) -and !([String]::IsNullOrWhiteSpace($temp.'Property Contact Phone Number '))){$table_object.'Ph No (Details)'=$table_object.'Ph No (Details)'+' ; '+$temp.'Property Contact Phone Number '}
        }

        #creating new entry for only Q3
        if(!($table.'marsha' -match $Marsha))
        {
        $date=$element.'Date Time'
        $date=[datetime]$date
        $q_date=$date.ToString("MM\/dd\/yyyy")
        $header_poc=$element.'Customer Name'
        $header_email=$element.'Customer Email'
        $header_phone=$element.'Customer Phone'

        $add_marsha=$temp.'Additional MARSHA Codes, if multi-property configuration'
        if([string]::IsNullOrWhiteSpace($add_marsha)){$add_marsha=$temp."Additional MARSHA Codes"}
        if($add_marsha -eq $Marsha){$add_marsha=$null}
            

        $Body_Poc=$temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'
        
        
        
        $Body_Phone=$temp.'Property Contact Phone Number '
        $Body_email=$temp.'Property Contact Email'
        $q_Time=$temp.'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
        $original_time=$q_time
        

        if(!([String]::IsNullOrWhiteSpace($q_Time) -or $q_Time -eq 'null'))
                {
                if($q_Time -match '(?i)\d+[:.]*\d*\s*(?:AM|PM)?')
                {
                    if($q_Time -match'\d+.\d*\s*[ap]*[m]*'){$q_Time=$q_Time -replace '(?<=\d)\.(?=\d)',':' -replace '\.',''}
                    if($q_Time -match'(?i)\d+\s?:\s?\d+\s*(?:AM|PM)?'){$q_Time=$q_Time -replace '(?<=\d)\s+(?=:)','' -replace '(?<=:)\s+(?=\d)','' -replace '(?<=\d)\s*(?=A|P)',' '}
                    $q_Time=($q_Time | Select-String '(?i)\d+:?\d*\s*(?:AM|PM)?' | select -ExpandProperty matches).value
                    try{$q_time=[datetime]$q_time
                    $q_Time=$q_Time.ToString("hh:mm tt")}catch{$incorrect_time[$marsha]+=" Q3 : "+$original_time;$q_Time="11:00 PM"} 
                }else{$incorrect_time[$marsha]+=" Q3 : "+$q_Time;$q_Time="11:00 PM"}
            }else{$q_Time="11:00 PM"}

        $Flag=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)? (Required)'
        if([string]::IsNullOrWhiteSpace($flag)){$flag=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)?'}
        #querying meta data
                $meta_object=$meta_data_d[$Marsha]
      
         $Table+=[pscustomobject]@{'Marsha'=$Marsha;'Property Name'=$meta_object.'Property Name';'Q1 Schedule'=$null;'Q2 Schedule'=$null;'Q3 Schedule'=$q_Date+' '+$q_Time;'Q4 Schedule'=$null;
                'Q1 Upgrade Date'=$null;'Q1 Upgrade Time'=$null;'Q1 OPERA PMS Fiscal or Legal Requirements'=$null;
                'Q2 Upgrade Date'=$null;'Q2 Upgrade Time'=$null;'Q2 OPERA PMS fiscal or legal requirements'=$null;
                'Q3 Upgrade Date'=$q_Date;'Q3 Upgrade Time'=$q_Time;'Q3 OPERA PMS fiscal or legal requirements'=$Flag;
                'Q4 Upgrade Date'=$null;'Q4 Upgrade Time'=$null;'Q4 OPERA PMS fiscal or legal requirements'=$null;
       
                'Ownership'=$meta_object.'Ownership';'Property Type'=$null;'Multi-Property Marsha'=$add_marsha;
                'Continent'=$meta_object.'continent';'Country'=$meta_object.'country';'Time Zone'=$meta_object.'Time Zone';
                'GM'=$meta_object.'GM';'GM Email'=$meta_object.'GM Email';'POC Name'=$meta_object.'POC Name';'POC Email'=$meta_object.'POC Email';'POC Phone'=$meta_object.'POC Phone';'IT Email'=$meta_object.'IT Email';'IT Phone'=$meta_object.'IT Phone';
                'POC Name (Booking)'=$header_poc;'Email ID (Booking)'=$header_email;'Ph No (Booking)'=$header_phone;
                'POC Name (Details)'=$Body_Poc;'Email ID (Details)'=$Body_email;'Ph No (Details)'=$Body_Phone;
                'On MI Network'=$meta_object.'On MI Network';'Hosting'=$meta_object.'Hosting'}
        
        }
        if(!($meta_data.'marsha' -match $marsha)){$No_metaData+=$Marsha}
    }

}
#end of q3
#Q4
foreach($element in $csv_data)
{
    $loopcounter++
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow
    $skip_duplicate_check=$false
    if($element.'Service' -match 'Q4')
    {
        $temp=$element.'Custom Fields' | ConvertFrom-Json
        $Marsha=$temp.'MARSHA Code (Required)'
        if([String]::IsNullOrWhiteSpace($marsha)){$marsha=$temp.'MARSHA Code (Required)'}
        if($Marsha -match "^[A-Z]{5}"){$Marsha=($Marsha | Select-String "^[A-Z]{5}" | select -ExpandProperty matches).value}
        $marsha=$Marsha.ToUpper()
        $marsha=$marsha.Trim()
        $date=$element.'Date Time'
        $date=[datetime]$date
        $q_date=$date.ToString("MM\/dd\/yyyy")

        $find_booking_date=$booked_date | Where-Object {$_.Marsha -eq $marsha -and ![string]::IsNullOrEmpty($_.'Q4 Booking date')}
        #setting up booked_on date with schedule date in booked_date.csv

        if(!($find_booking_date.'Q4 Schedule' -match $q_date))
        {
            
            $booked_date+=[pscustomobject]@{'Marsha'=$Marsha;'Q4 Schedule'=$q_date;'Q4 Booking date'="$((get-date).ToString("MM\/dd\/yyyy"))";}
            $skip_duplicate_check=$true
        }
        #check if marsha is present in table
        if($table.'marsha' -match $Marsha)
        {
            $table_object=$table | Where-Object {$_.'MARSHA' -match $marsha}
            
            $add_marsha=$temp.'Additional MARSHA Codes, if multi-property configuration'
            if([string]::IsNullOrWhiteSpace($add_marsha)){$add_marsha=$temp."Additional MARSHA Codes"}
            if($add_marsha -eq $Marsha){$add_marsha=$null}
            
            #handle duplicate entry
        if(!([String]::IsNullOrWhiteSpace($table_object.'Q4 Schedule')))
            {
                 
        #to avoid duplicate entry
        if($table.'marsha' -match $Marsha -and $skip_duplicate_check -eq $false)
            {
                #1.check if current element is on latest booked_on date
                #2.if it is remove previous entry of same marsha
                #3.else skip this element
                $is_latest_flag=$true
                $current_element_booked_on=$find_booking_date | Where-Object {$_.Marsha -eq $marsha -and $_.'Q4 Schedule' -eq $q_date}
                $current_element_booked_on=[datetime]$current_element_booked_on.'Q4 Booking date'
                foreach($row in $find_booking_date){
                    if($current_element_booked_on -lt [datetime]$row.'Q4 Booking date'){$is_latest_flag=$false}
                }

                 if($is_latest_flag -eq $true){
                 $table_object=$table | Where-Object {$_.'MARSHA' -match $marsha}
                 $table_object.'Q4 Schedule'=$null
                 $table_object.'Q4 Upgrade Date'=$null
                 $table_object.'Q4 Upgrade Time'=$null
                 $table_object.'Q4 OPERA PMS fiscal or legal requirements'=$null 
                 }else{continue}
        
            }       
    }
            $table_object.'Q4 Upgrade Date'=$q_date
            
            $q_Time=$temp.'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
            $original_time=$q_time
            

            if(!([String]::IsNullOrWhiteSpace($q_Time) -or $q_Time -eq 'null'))
                {
                if($q_Time -match '(?i)\d+[:.]*\d*\s*(?:AM|PM)?')
                {
                    if($q_Time -match'\d+.\d*\s*[ap]*[m]*'){$q_Time=$q_Time -replace '(?<=\d)\.(?=\d)',':' -replace '\.',''}
                    if($q_Time -match'(?i)\d+\s?:\s?\d+\s*(?:AM|PM)?'){$q_Time=$q_Time -replace '(?<=\d)\s+(?=:)','' -replace '(?<=:)\s+(?=\d)','' -replace '(?<=\d)\s*(?=A|P)',' '}
                    $q_Time=($q_Time | Select-String '(?i)\d+:*\d*\s*(?:AM|PM)?' | select -ExpandProperty matches).value
                    try{$q_time=[datetime]$q_time
                    $q_Time=$q_Time.ToString("hh:mm tt")}catch{$incorrect_time[$marsha]+=" Q4 : "+$original_time;$q_Time="11:00 PM"} 
                }else{$incorrect_time[$marsha]+=" Q4 : "+$q_Time;$q_Time="11:00 PM"}
            }else{$q_Time="11:00 PM"}

            $table_object.'Q4 Schedule'=$q_date+' '+$q_Time
            $table_object.'Q4 Upgrade Time'=$q_Time
            $table_object.'Q4 OPERA PMS Fiscal or Legal Requirements'=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)? (Required)'
            if([string]::IsNullOrWhiteSpace($table_object.'Q4 OPERA PMS Fiscal or Legal Requirements')){$table_object.'Q4 OPERA PMS Fiscal or Legal Requirements'=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)?'}
            if($table_object.'Property Type' -eq "Single-Property"){$table_object.'Property Type'=$null}
            if($table_object.'Multi-Property Marsha' -notmatch $add_marsha -and !([String]::IsNullOrWhiteSpace($add_marsha))){$table_object.'Multi-Property Marsha'=$table_object.'Multi-Property Marsha'+" "+$add_marsha}
            
            if([String]::IsNullOrWhiteSpace($table_object.'POC Name (Booking)')){$table_object.'POC Name (Booking)'=$element.'Customer Name'}
            if($table_object.'Email ID (Booking)' -notmatch $element.'Customer Email'){$table_object.'Email ID (Booking)'=$table_object.'Email ID (Booking)'+' ; '+$element.'Customer Email'}
            if($table_object.'Ph No (Booking)' -notmatch $element.'Customer Phone'){$table_object.'Ph No (Booking)'=$table_object.'Ph No (Booking)'+' ; '+$element.'Customer Phone'}
          

            if(($table_object.'POC Name (Details)' -notmatch $temp.'Full Name of Property Contact During Upgrade/Patching, if different than above') -and !([String]::IsNullOrWhiteSpace($temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'))){$table_object.'POC Name (Details)'=$table_object.'POC Name (Details)'+' ; '+$temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'}
            if(($table_object.'Email ID (Details)' -notmatch [Regex]::Escape($temp.'Property Contact Email')) -and !([String]::IsNullOrWhiteSpace($temp.'Property Contact Email'))){$table_object.'Email ID (Details)'=$table_object.'Email ID (Details)'+' ; '+$temp.'Property Contact Email'}
            if(($table_object.'Ph No (Details)' -notmatch [Regex]::Escape($temp.'Property Contact Phone Number ')) -and !([String]::IsNullOrWhiteSpace($temp.'Property Contact Phone Number '))){$table_object.'Ph No (Details)'=$table_object.'Ph No (Details)'+' ; '+$temp.'Property Contact Phone Number '}
        }

        #creating new entry for only Q4
        if(!($table.'marsha' -match $Marsha))
        {
        $date=$element.'Date Time'
        $date=[datetime]$date
        $q_date=$date.ToString("MM\/dd\/yyyy")
        $header_poc=$element.'Customer Name'
        $header_email=$element.'Customer Email'
        $header_phone=$element.'Customer Phone'

        $add_marsha=$temp.'Additional MARSHA Codes, if multi-property configuration'
        if([string]::IsNullOrWhiteSpace($add_marsha)){$add_marsha=$temp."Additional MARSHA Codes"}
        if($add_marsha -eq $Marsha){$add_marsha=$null}


        $Body_Poc=$temp.'Full Name of Property Contact During Upgrade/Patching, if different than above'
        
        
        
        $Body_Phone=$temp.'Property Contact Phone Number '
        $Body_email=$temp.'Property Contact Email'
        $q_Time=$temp.'Preferred Start Time for Activity (If don''t want to start at 11:00 PM Property Local Time)'
        $original_time=$q_time
        

        if(!([String]::IsNullOrWhiteSpace($q_Time) -or $q_Time -eq 'null'))
                {
                if($q_Time -match '(?i)\d+[:.]*\d*\s*(?:AM|PM)?')
                {
                    if($q_Time -match'\d+.\d*\s*[ap]*[m]*'){$q_Time=$q_Time -replace '(?<=\d)\.(?=\d)',':' -replace '\.',''}
                    if($q_Time -match'(?i)\d+\s?:\s?\d+\s*(?:AM|PM)?'){$q_Time=$q_Time -replace '(?<=\d)\s+(?=:)','' -replace '(?<=:)\s+(?=\d)','' -replace '(?<=\d)\s*(?=A|P)',' '}
                    $q_Time=($q_Time | Select-String '(?i)\d+:*\d*\s*(?:AM|PM)?' | select -ExpandProperty matches).value
                    try{$q_time=[datetime]$q_time
                    $q_Time=$q_Time.ToString("hh:mm tt")}catch{$incorrect_time[$marsha]+=" Q4 : "+$original_time;$q_Time="11:00 PM"} 
                }else{$incorrect_time[$marsha]+=" Q4 : "+$q_Time;$q_Time="11:00 PM"}
            }else{$q_Time="11:00 PM"}

        $Flag=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)? (Required)'
        if([string]::IsNullOrWhiteSpace($flag)){$flag=$temp.'Does your OPERA PMS have fiscal or legal requirements that may interface or export to an external source (e.g., a police interface, government export)?'}
        #querying meta data
                $meta_object=$meta_data_d[$Marsha]
      
         $Table+=[pscustomobject]@{'Marsha'=$Marsha;'Property Name'=$meta_object.'Property Name';'Q1 Schedule'=$null;'Q2 Schedule'=$null;'Q3 Schedule'=$null;'Q4 Schedule'=$q_Date+' '+$q_Time;
                'Q1 Upgrade Date'=$null;'Q1 Upgrade Time'=$null;'Q1 OPERA PMS Fiscal or Legal Requirements'=$null;
                'Q2 Upgrade Date'=$null;'Q2 Upgrade Time'=$null;'Q2 OPERA PMS fiscal or legal requirements'=$null;
                'Q3 Upgrade Date'=$null;'Q3 Upgrade Time'=$null;'Q3 OPERA PMS fiscal or legal requirements'=$null;
                'Q4 Upgrade Date'=$q_Date;'Q4 Upgrade Time'=$q_Time;'Q4 OPERA PMS fiscal or legal requirements'=$Flag;
       
                'Ownership'=$meta_object.'Ownership';'Property Type'=$null;'Multi-Property Marsha'=$add_marsha;
                'Continent'=$meta_object.'continent';'Country'=$meta_object.'country';'Time Zone'=$meta_object.'Time Zone';
                'GM'=$meta_object.'GM';'GM Email'=$meta_object.'GM Email';'POC Name'=$meta_object.'POC Name';'POC Email'=$meta_object.'POC Email';'POC Phone'=$meta_object.'POC Phone';'IT Email'=$meta_object.'IT Email';'IT Phone'=$meta_object.'IT Phone';
                'POC Name (Booking)'=$header_poc;'Email ID (Booking)'=$header_email;'Ph No (Booking)'=$header_phone;
                'POC Name (Details)'=$Body_Poc;'Email ID (Details)'=$Body_email;'Ph No (Details)'=$Body_Phone;
                'On MI Network'=$meta_object.'On MI Network';'Hosting'=$meta_object.'Hosting'}
        }
        if(!($meta_data.'marsha' -match $marsha)){$No_metaData+=$Marsha}
    }

}
#end of Q4

#Optimizing code using dictonaries instead of where object
$table_d=@{}
$smartsheet_data_d=@{}

foreach ($row in $Table) {
if($row.'Marsha' -ne $null){
    $table_d[$row.'Marsha'] = $row
    }
}
foreach ($row in $smartsheet_data) {
    $smartsheet_data_d[$row."Marsha"] = $row
}


#add in master sheet
Echo "Adding data to meata-file"
$master_sheet=$meta_data
$master_sheet_print=@()
foreach($element in $master_sheet)
{
    $loopcounter++
    
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow
    $marsha=$element.'MARSHA'

    if(!($marsha -match "^[A-Z]{5}$") -and $marsha -ne "OSA41")
    {
        $element.psobject.properties.name | foreach{$element.psobject.properties.remove("$_")}        
    }


    if($table.'marsha' -match $Marsha)
    {
        $table_object=$table_d[$marsha]
        
        if(!([string]::IsNullOrEmpty($table_object.'Q1 Schedule'))){$element.'Q1 Schedule'=$table_object.'Q1 Schedule'}
        if(!([string]::IsNullOrEmpty($table_object.'Q2 Schedule'))){$element.'Q2 Schedule'=$table_object.'Q2 Schedule'}
        if(!([string]::IsNullOrEmpty($table_object.'Q3 Schedule'))){$element.'Q3 Schedule'=$table_object.'Q3 Schedule'}
        if(!([string]::IsNullOrEmpty($table_object.'Q4 Schedule'))){$element.'Q4 Schedule'=$table_object.'Q4 Schedule'}

        if(!([string]::IsNullOrEmpty($table_object.'Q1 OPERA PMS Fiscal or Legal Requirements'))){$element.'Q1 OPERA PMS Fiscal or Legal Requirements'=$table_object.'Q1 OPERA PMS Fiscal or Legal Requirements'}
        if(!([string]::IsNullOrEmpty($table_object.'Q2 OPERA PMS Fiscal or Legal Requirements'))){$element.'Q2 OPERA PMS Fiscal or Legal Requirements'=$table_object.'Q2 OPERA PMS Fiscal or Legal Requirements'}
        if(!([string]::IsNullOrEmpty($table_object.'Q3 OPERA PMS fiscal or legal requirements'))){$element.'Q3 OPERA PMS fiscal or legal requirements'=$table_object.'Q3 OPERA PMS fiscal or legal requirements'}
        if(!([string]::IsNullOrEmpty($table_object.'Q4 OPERA PMS Fiscal or Legal Requirements'))){$element.'Q4 OPERA PMS Fiscal or Legal Requirements'=$table_object.'Q4 OPERA PMS Fiscal or Legal Requirements'}

        $element.'POC Name'=$table_object.'POC Name (Booking)'

        $element.'POC Email'=$table_object.'Email ID (Booking)'+" ; "+$table_object.'Email ID (Details)'+" ; "+$table_object.'POC Email'
        $element.'POC Email'=$element.'POC Email' -replace ";\s+;" ,';' -replace "; $" , ''
        $element.'POC Email'=Get-ValidEmails -InputString $element.'POC Email'

        $element.'POC Phone'=$table_object.'Ph No (Booking)'+' ;'+$table_object.'Ph No (Details)'
        $element.'POC Phone'=$element.'POC Phone'  -replace ";\s*;" ,';' -replace ";$" , ''
    }

    #adding city
        $IATA=($marsha | Select-String "^[\w]{3}" | select -ExpandProperty matches).value
        $city_object=$IATA_city | Where-Object {$_.'IATA code' -match $IATA}
        $element.'City'=$city_object.'City'

    if($element.Continent -eq 'US Canada'){$element.Continent='US/CAN'}
    if([string]::IsNullOrWhiteSpace($element.Continent)){$element.Continent="NULL"}
        if($element.Continent -eq 'Asia Pacific')
            {  
                        $element.Continent="APAC"     
            }
    
    if($smartsheet_data.'marsha' -match $Marsha)
    {
        $table_object=$smartsheet_data_d[$Marsha]
        $element.Ownership=$table_object.OT
        $element.'Property Name'=$table_object.'Property Name'
        $element.'IT Email'=$table_object.'IT Email'
        $element.'IT Phone'=$table_object.'IT Phone'
        $element.'On MI Network'=$table_object.'On MI Network'
        $element.GM=$table_object.GM
        $element.'hosting'=$table_object.hosting
        $element.'GM Email'=$table_object.'GM Email'
        $element.'InScope'=$table_object.'In-Scope'
        $element.'InScope Reason'=$table_object.'In-Scope Reason'
        
    }else{
    Write-Warning "Remove: $marsha"
    $element.psobject.properties.name | foreach{$element.psobject.properties.remove("$_")}
         }
    if([string]::IsNullOrWhiteSpace($element.'InScope') -and $marsha -eq 'DHADL'){$element.'InScope Reason'="opt in";$element.'InScope'='Yes'}
}

#IST conversion
foreach($element in $master_sheet)
{
    $loopcounter++
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow
    $marsha=$element.'MARSHA'
    if([string]::IsNullOrWhiteSpace($element.'Time Zone offset')){continue}
 
    $q1=$q2=$q3=$q4=$null
    try{$q1=[datetime]$element.'Q1 Schedule'}catch{write-warning "$marsha not valid IST conversion"}
    try{$q2=[datetime]$element.'Q2 Schedule'}catch{write-warning "$marsha not valid IST conversion"}
    try{$q3=[datetime]$element.'Q3 Schedule'}catch{write-warning "$marsha not valid IST conversion"}
    try{$q4=[datetime]$element.'Q4 Schedule'}catch{write-warning "$marsha not valid IST conversion"}
    
    $tz=$element.'Time Zone offset'
    $add_t=$tz | Select-String "GMT([\+-])(\d\d?):(\d\d)$" | Select-Object -ExpandProperty Matches |  Select-Object -ExpandProperty groups
    $e_time=New-TimeSpan -Hours $add_t[2].value  -Minutes $add_t[3].value
    $e_time=[decimal]"$($add_t[1].value)$($e_time.totalhours)"
    $e_time=-$e_time
    $e_time=$e_time+5.5

    try{$element.'Q1 Schedule IST'=($q1.AddHours($e_time)).ToString("MM\/dd\/yyyy hh:mm tt")}catch{}
    try{$element.'Q2 Schedule IST'=($q2.AddHours($e_time)).ToString("MM\/dd\/yyyy hh:mm tt")}catch{}
    try{$element.'Q3 Schedule IST'=($q3.AddHours($e_time)).ToString("MM\/dd\/yyyy hh:mm tt")}catch{}
    try{$element.'Q4 Schedule IST'=($q4.AddHours($e_time)).ToString("MM\/dd\/yyyy hh:mm tt")}catch{}

}


#multi-property check by hosting

echo "multi-property check refrencing hosting column"
foreach($element in $master_sheet)
{
    $loopcounter++
    
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow

    $marsha=$element.'MARSHA'
    $hosting=$element.'hosting'
    
   
    if($hosting -match "(HOST $marsha[: ]?)(.+)")
    {
        $element.'Property Type'='Multi-Property'
        $m=$matches.2
        $m=$m.trim()
        $element.'Multi-Property Marsha'=$m
        continue
    }elseif($hosting -match "(HOST:\s*$marsha[: ]?)(.+)"){
        $element.'Property Type'='Multi-Property'
        $m=$matches.2
        $m=$m.trim()
        $element.'Multi-Property Marsha'=$m
        continue}
    if(!($hosting -match "HOST")){$element.'Property Type'='Single-Property'}
    $hosting -match 'HOST[ :](.*):(.*)'
    $host_marsha=$Matches.1
    $multi_properties=$Matches.2
    if($multi_properties -ne $null)
    {
        $multi_properties=$multi_properties.trim()
        if($multi_properties -match $marsha)
        {
            $element.'Property Type'="Part of multi Property"
            $host_marsha=$host_marsha.Trim()
            $element.'Multi-Property Marsha'="$host_marsha"
            
        }
    }
    
}
#Applying from multi-property to Host
foreach($element in $master_sheet)
{
    $loopcounter++
    
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow

    $marsha=$element.'MARSHA'

    
    if($element.'Property Type' -eq "Part of multi Property")
    {
        $marsha=$element.'MARSHA'    
        $element.'Multi-Property Marsha' -match '(.*)'
        $host_marsha=$Matches.1
        $Main_marsha=$master_sheet | Where-Object {$_.'MARSHA' -match $host_marsha}
        if([string]::IsNullOrWhiteSpace($Main_marsha.'Q1 Schedule') -and !([string]::IsNullOrWhiteSpace($element.'Q1 Schedule')))
        {
            $Main_marsha.'Q1 Schedule'=$element.'Q1 Schedule'
            $Main_marsha.'Q1 Schedule IST'=$element.'Q1 Schedule IST'
            $Main_marsha.'Q1 OPERA PMS Fiscal or Legal Requirements'=$element.'Q1 OPERA PMS Fiscal or Legal Requirements'
        }
        if([string]::IsNullOrWhiteSpace($Main_marsha.'Q2 Schedule') -and !([string]::IsNullOrWhiteSpace($element.'Q2 Schedule'))){
            $Main_marsha.'Q2 Schedule'=$element.'Q2 Schedule'
            $Main_marsha.'Q2 Schedule IST'=$element.'Q2 Schedule IST'
            $Main_marsha.'Q2 OPERA PMS Fiscal or Legal Requirements'=$element.'Q2 OPERA PMS Fiscal or Legal Requirements'
        }
        if([string]::IsNullOrWhiteSpace($Main_marsha.'Q3 Schedule') -and !([string]::IsNullOrWhiteSpace($element.'Q3 Schedule'))){
            $Main_marsha.'Q3 Schedule'=$element.'Q3 Schedule'
            $Main_marsha.'Q3 Schedule IST'=$element.'Q3 Schedule IST'
            $Main_marsha.'Q3 OPERA PMS Fiscal or Legal Requirements'=$element.'Q3 OPERA PMS Fiscal or Legal Requirements'
        }
        if([string]::IsNullOrWhiteSpace($Main_marsha.'Q4 Schedule') -and !([string]::IsNullOrWhiteSpace($element.'Q4 Schedule'))){
            $Main_marsha.'Q4 Schedule'=$element.'Q4 Schedule'
            $Main_marsha.'Q4 Schedule IST'=$element.'Q4 Schedule IST'
            $Main_marsha.'Q4 OPERA PMS Fiscal or Legal Requirements'=$element.'Q4 OPERA PMS Fiscal or Legal Requirements'
        }
    }
}

#applying marsha from host to multi property
foreach($element in $master_sheet)
{
    $loopcounter++
    $marsha=$element.'MARSHA'
    $percentage=($loopcounter/$total_loop)*100
    write-host "PMS Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor Yellow

    if($element.'Property Type' -eq 'Multi-Property')
    {
     
        $multi=$element.'Multi-Property Marsha'
        $multi=$multi.split(',')
        $multi=$multi.trim()
        $multi=$multi.split('&')
        $multi=$multi.split('-')
        $multi=$multi.trim()
        $multi=$multi.split(' ')
        $multi=$multi.trim()
        foreach($a in $multi)
        {
            $table_object=$master_sheet | Where-Object {$_.'MARSHA' -match $a}

            <#if(![string]::IsNullOrWhiteSpace($table_object.'Q1 Schedule') -or 
                 ![string]::IsNullOrWhiteSpace($table_object.'Q2 Schedule') -or
                 ![string]::IsNullOrWhiteSpace($table_object.'Q3 Schedule') -or
                 ![string]::IsNullOrWhiteSpace($table_object.'Q4 Schedule'))
               {
                continue
               }#>
            if($table_object.'Property Type' -eq "Part of multi Property")
            {
                if(!([string]::IsNullOrWhiteSpace($element.'Q1 Schedule'))){
                $table_object.'Q1 Schedule'=$element.'Q1 Schedule'
                $table_object.'Q1 Schedule IST'=$element.'Q1 Schedule IST'
                $table_object.'Q1 OPERA PMS Fiscal or Legal Requirements'=$element.'Q1 OPERA PMS Fiscal or Legal Requirements'
                }


                if(!([string]::IsNullOrWhiteSpace($element.'Q2 Schedule'))){
                $table_object.'Q2 Schedule'=$element.'Q2 Schedule'
                $table_object.'Q2 Schedule IST'=$element.'Q2 Schedule IST'
                $table_object.'Q2 OPERA PMS Fiscal or Legal Requirements'=$element.'Q2 OPERA PMS Fiscal or Legal Requirements'
                }


                if(!([string]::IsNullOrWhiteSpace($element.'Q3 Schedule'))){
                $table_object.'Q3 Schedule'=$element.'Q3 Schedule'
                $table_object.'Q3 Schedule IST'=$element.'Q3 Schedule IST'
                $table_object.'Q3 OPERA PMS fiscal or legal requirements'=$element.'Q3 OPERA PMS fiscal or legal requirements'
                }


                if(!([string]::IsNullOrWhiteSpace($element.'Q4 Schedule'))){
                $table_object.'Q4 Schedule'=$element.'Q4 Schedule'
                $table_object.'Q4 Schedule IST'=$element.'Q4 Schedule IST'
                $table_object.'Q4 OPERA PMS Fiscal or Legal Requirements'=$element.'Q4 OPERA PMS Fiscal or Legal Requirements'
                }

                $table_object.'Multi-Property Marsha'="$Marsha"
                
            }
            

        }
    }
} 


foreach($element in $master_sheet)
{
    $marsha=$element.'MARSHA'
    if(!([String]::IsNullOrWhiteSpace($marsha))){$master_sheet_print+=$element}

    if($element.'Property Type' -eq "Multi-Property")
    {
    $multi_marshas=$element.'Multi-Property Marsha' | Select-String -Pattern "\b\w{5}\b" -AllMatches | ForEach-Object { $_.Matches.Value }
    $element.'Multi-Property Marsha'=$multi_marshas -join ","
    }
    #if($marsha -eq "OSA41"){$element.'Q1 Schedule'=$element.'Q2 Schedule'=$element.'Q3 Schedule'=$element.'Q4 Schedule'=$element.'Q1 Schedule IST'=$element.'Q2 Schedule IST'=$element.'Q3 Schedule IST'=$element.'Q4 Schedule IST'=$null}
    #if("AHNRZ","ATLRZ","ASUAL","ASUSI","AMDCY","ASUTX","JSAMC" -contains $marsha){$element.'InScope'='No';$element.'InScope Reason'="Out of scope only for Q2"}
    
}
$master_sheet_print.count



if(!(Get-Module -ListAvailable -Name importexcel))
{
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet
Install-Module ImportExcel -AllowClobber -Force
Get-Module ImportExcel -ListAvailable | Import-Module -Force -Verbose
}


if(Test-Path "$loc\PMS_Bookings_master_sheet_2025.xlsx"){Remove-Item -Path "$loc\PMS_Bookings_master_sheet_2025.xlsx"}

if(Test-Path "$loc\PMS_Faulty_booking.xlsx"){Remove-Item -Path "$loc\PMS_Faulty_booking.xlsx"}

$master_sheet_print | Export-Excel -Path "$loc\PMS_Bookings_master_sheet_2025.xlsx"  -WorksheetName "Bookingsdata" -TableStyle Medium17 -AutoSize 



$faulty_booking=@()
foreach($element in $No_metaData)
            {
                $marsha=$element
                if($faulty_booking.marsha -eq $marsha){continue}
                $table_object=$table_d[$marsha]
                $faulty_booking+=[pscustomobject]@{'Marsha'=$Marsha;'Fault Reason'='Meta data not available';'Q1 Schedule'=$table_object.'Q1 Schedule';'Q2 Schedule'=$table_object.'Q2 Schedule';'Q3 Schedule'=$table_object.'Q3 Schedule';'Q4 Schedule'=$table_object.'Q4 Schedule';
                                                   'POC Name'=$table_object.'POC Name (Booking)';'Email ID'=$table_object.'Email ID (Booking)';'Phone No'=$table_object.'Ph No (Booking)';}
            }

foreach($element in $incorrect_time.keys)
            {
                $marsha=$element
                
                $table_object=$table_d[$marsha]
                $faulty_booking+=[pscustomobject]@{'Marsha'=$Marsha;'Fault Reason'="Incorrect Time Format $($incorrect_time[$marsha])";'Q1 Schedule'=$table_object.'Q1 Schedule';'Q2 Schedule'=$table_object.'Q2 Schedule';'Q3 Schedule'=$table_object.'Q3 Schedule';'Q4 Schedule'=$table_object.'Q4 Schedule';
                                                   'POC Name'=$table_object.'POC Name (Booking)';'Email ID'=$table_object.'Email ID (Booking)';'Phone No'=$table_object.'Ph No (Booking)';}
            }


$faulty_booking | Export-Excel -Path "$loc\PMS_Faulty_booking.xlsx"  -WorksheetName "faulty" -TableStyle Medium7 -AutoSize  

$booked_date | ConvertTo-Csv -Delimiter "," -NoTypeInformation | Out-File "$loc\Meta Data\Booked_date.csv"
stop-Transcript
Start-Sleep -Seconds 5