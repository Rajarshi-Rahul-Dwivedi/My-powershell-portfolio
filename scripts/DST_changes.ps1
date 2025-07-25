$loc="$PSScriptRoot"

$meta_input=Get-Content ""
$meta_data=ConvertFrom-Csv -Delimiter "," -InputObject $meta_input

$DST_input=Get-Content ""
$DST_data=ConvertFrom-Csv -Delimiter "," -InputObject $DST_input
$changed_countries=@()

$current_date=(Get-Date).AddDays(5)


foreach($element in $meta_data)
{
    $marsha=$element.'MARSHA'
    if($element.'In DST' -eq "NA"){continue}
    $Marsha_IATA=($marsha | Select-String "^[A-Z]{3}" | select -ExpandProperty matches).value
    #exceptions conditions
    if($element.Country -eq 'Oman'){$element.'In DST'="NA";$element.'Time Zone Offset'="GMT+4:00";continue}
    if("HNL","HNM","PHX","TUS","PER","OOL","BNE" -match $Marsha_IATA){$element.'In DST'="NA";continue}
    if($element.Country -eq 'USA' -and $element.'Time Zone Offset' -eq 'GMT-10:00'){$element.'In DST'="NA";continue}
    if($element.Country -eq "Australia" -and $element.'Time Zone Offset' -eq "GMT+8:00"){$element.'In DST'="NA";continue}
    #$exceptions="Australia",'Chile','Morocco','New Zealand','Norfolk Island','Paraguay','Western Sahara','Morocco','Haiti'
    #if($exceptions -match $element.Country){$element.'In DST'='Yes';continue}
    if($dst_data.country -match $element.Country)
    { 
        $dst_object=$dst_data | Where-Object {$_.'country' -match $element.Country}
        $dst_start=[datetime]$dst_object.'DST Start Date'
        $dst_end=[datetime]$dst_object.'DST End Date'
        
        if($current_date -ge $dst_start -and $current_date -le $dst_end)
        {
          if($element.'In DST' -ne 'Yes')
           {
            $element.'In DST'='Yes'
            #for GMT offset
            $time_zone=$element.'Time Zone offset'
            $time_zone -match 'GMT([\+\-]\d\d?)\:\d\d' | Out-Null
            $hour=$Matches.1
            $new_hour=[int]$hour+1
            $time_zone=$time_zone -replace "\d\d?\:","$new_hour`:"
            $time_zone=$time_zone -replace "--","-"
            Write-host "`n`n $marsha $($element.Country) $($element.'Time Zone offset') changed to $time_zone`n`n" -ForegroundColor Magenta
            "$marsha $($element.Country) $($element.'Time Zone offset') changed to $time_zone`n" | out-file "$loc\DST changes dashboard.txt" -append
            

            #for elit timezone name
            $time_zone_elit=$meta_data | Where-Object {$_.'Time Zone offset' -eq $time_zone}
            $time_zone_elit=[string]($time_zone_elit[0].'Time Zone')
            #setting timezones
            $element.'Time Zone offset'=$time_zone
            $element.'Time Zone'=$time_zone_elit
            if ($changed_countries -notcontains $element.Country) { $changed_countries += $element.Country }
            continue
            }
        }
        #reverse DST
        if($element.'In DST' -eq "Yes"){
            if($current_date -ge $dst_end -and $current_date -le $dst_start){
                
                $element.'In DST'='No'
                #for GMT offset
                $time_zone=$element.'Time Zone offset'
                $time_zone -match 'GMT([\+\-]\d\d?)\:\d\d' | Out-Null
                $hour=$Matches.1
                $new_hour=[int]$hour-1
                $time_zone=$time_zone -replace "\d\d?\:","$new_hour`:"
                $time_zone=$time_zone -replace "--","-"
                Write-host "`n`n $marsha $($element.Country) $($element.'Time Zone offset') changed to $time_zone`n`n" -ForegroundColor Yellow
                "$marsha $($element.Country) $($element.'Time Zone offset') changed to $time_zone`n" | out-file "$loc\DST changes dashboard.txt" -append
            

                #for elit timezone name
                $time_zone_elit=$meta_data | Where-Object {$_.'Time Zone offset' -eq $time_zone}
                $time_zone_elit=[string]($time_zone_elit[0].'Time Zone')
                #setting timezones
                $element.'Time Zone offset'=$time_zone
                $element.'Time Zone'=$time_zone_elitne
                $element.'Time Zone'=$time_zone_elit
                if ($changed_countries -notcontains $element.Country) { $changed_countries += $element.Country }
                continue
            }
        }
        
        if([string]::IsNullOrWhiteSpace($element.'In DST') -and $element.'In DST' -ne 'Yes'){$element.'In DST'="No"}

    }else{$element.'In DST'="NA"}
}

$output_csv=$meta_data | ConvertTo-Csv -Delimiter ',' -NoTypeInformation
write-host "Please note changes are made to following countries`n"

#Rename-Item -Path "$loc\Meta data\PMS_meta_data.csv" -NewName "PMS_meta_data_$dat.csv"
$output_csv | Out-File "C:\Automated booking\2025\PMS\Meta Data\PMS_meta_data.csv"
return $changed_countries

