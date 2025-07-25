#Hello, This Script will not function in it's current form please read the discription to update the all API parameters and credentials or feel free to contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281
#This is testing version with host output

$username=
$Password= 
#$cred=Get-Credential

$tokenURL = 
$headers = @{
    "Authorization" = "V0hCT0MwMDFfQ2xpZW50OlEyX0czUnRXM05qSlRSNVdhMXZRTy1Ybg=="
    "x-app-key" = "REDACTED_APP_KEY"
    "Accept" = "application/json"
    #"Content-Type" = "application/x-www-form-urlencoded"
}
$response = Invoke-RestMethod -Uri $tokenURL -Method Post -Headers $headers -ContentType "application/x-www-form-urlencoded"

$response.access_token
$external_system='AILPRODBEQ'

$tokenURL2 = "https://api.example.com/sample-endpoint"

$headers = @{
'Authorization' ="Bearer $($response.access_token)";
"x-app-key" = "REDACTED_APP_KEY";
"Accept" = "application/json";
'Accept-Language' = "application/json";
'x-hotelId' = "$hotelID";
}
$body = @{
'enquiryStartDate'="2024-05-07";

}

$response2 = Invoke-RestMethod -Uri $tokenURL2 -Method GET -Headers $headers


foreach($hotelID in $hotelIDs)
{
Write-Host "Time Elasped $([System.Math]::Round($stopwatch.Elapsed.TotalMinutes,2)) Minutes" -ForegroundColor Yellow
if($stopwatch.Elapsed.TotalMinutes -ge 59){write-warning "Access Token 1 hour expired able to fetch only $lopp_count Property Details";break}
Write-Host "$lopp_count`.Fetching for Property - $hotelID" -ForegroundColor Green

if([string]::IsNullOrEmpty($hotelID)){continue}
$lopp_count+=1
$headers = @{
'Authorization' ="Bearer $($response.access_token)"
"x-app-key" = "REDACTED_APP_KEY"
"Accept" = "application/json"
'Accept-Language' = "application/json"
'x-hotelId' = "$hotelID"
}

$tokenURL2 = "https://api.example.com/sample-endpoint"

$table=@()
$housekeepingOverview = Invoke-RestMethod -Uri $tokenURL2 -Method GET -Headers $headers
Start-Sleep -Milliseconds 50
    foreach($room in $housekeepingOverview.housekeepingRoomInfo.housekeepingRooms.room)
    {
        $reservationIds=$room.resvInfo.reservationId.id
        foreach($reservationId in $reservationIds)
        {
            if([string]::IsNullOrEmpty($reservationId)){continue}
            #calling reservations API
            $token_append2="/rsv/v1/hotels/$hotelID/reservations/$($reservationId)?fetchInstructions=Comments"
            $tokenURL3 = "https://api.example.com/sample-endpoint"
            
            $guest_name=$null
            $final_comment=$null

            $reservations_details =Invoke-RestMethod -Uri $tokenURL3 -Method GET -Headers $headers
            Start-Sleep -Milliseconds 5
            $guest_name=$room.resvInfo.guestname -join ' / '
            
            $comments=$reservations_details.reservations.reservation.comments
            foreach($comment in $comments.comment){
            $final_comment+=$comment.commentTitle+" - "+$comment.text.value+"  "
            }

            $Table+=[pscustomobject]@{
            'Property'=$hotelID
            'Room Number'= $room.roomId;
            'let Type' =$room.roomType.roomType;;
            'Status'=$room.housekeeping.housekeepingRoomStatus.housekeepingRoomStatus;
            'Departure date'=$reservations_details.reservations.reservation.roomstay.departureDate;
            'Special Request and Memo'=$final_comment
            'Arriving Guest Name'=$guest_name
            'Reservation ID'=$reservationId;
                }
        }
         $Table[-1] | Format-Table
    }
$Table |  Export-Excel -Path "D:\Job related\SampleClient\Property_outputs\$($hotelID)_$dat.xlsx"  -WorksheetName "Housekeeping" -TableStyle Medium15 -AutoSize 
$master_table+=$Table
}

$master_table | ConvertTo-Csv -NoTypeInformation -Delimiter "," | out-file "D:\Job related\SampleClient\Housekeeping_Info.csv"

$master_table | ConvertTo-json  | out-file "D:\Job related\SampleClient\Property_outputs\Housekeeping_Info.json"

$stopwatch.stop()
Write-Host "Execution Completed in $($stopwatch.Elapsed.TotalMinutes) Minutes" -ForegroundColor Red
$stopwatch.Elapsed


$dat=(Get-Date).ToString("dd_MMM_yy hh_mm tt")
Write-Host "Genrating access token" -BackgroundColor Green -ForegroundColor Black
$response = Invoke-RestMethod -Uri $tokenURL -Method Post -Headers $headers -ContentType "application/x-www-form-urlencoded"
$hotelID="REGCIT"
$interface_types="interfaceTypes=Bms&interfaceTypes=Cas&interfaceTypes=Ccw&interfaceTypes=Dls&interfaceTypes=Eft&interfaceTypes=Exp&interfaceTypes=Mak&interfaceTypes=Mbs&interfaceTypes=Msc&interfaceTypes=Pbx&interfaceTypes=Pos&interfaceTypes=Svs&interfaceTypes=Tik&interfaceTypes=Vid&interfaceTypes=Vms&interfaceTypes=Www&interfaceTypes=Xml"
$interface_types_limt="interfaceTypes=Bms&interfaceTypes=eft"
$headers = @{
'Authorization' ="Bearer $($response.access_token)"
"x-app-key" = "REDACTED_APP_KEY"
"Accept" = "application/json"
'Accept-Language' = "application/json"
'x-hotelId' = "$hotelID"
}

$tokenURL2 = "https://api.example.com/sample-endpoint"
$Interfacetypes = Invoke-RestMethod -Uri $tokenURL2 -Method GET -Headers $headers
($Interfacetypes.hotelInterfaces).count



$tokenURL_resortChains = "https://api.example.com/sample-endpoint"
$list_of_hotels=Invoke-RestMethod -Uri $tokenURL_resortChains -Method GET -Headers $headers
$hotelIDs=$list_of_hotels.listOfValues.items.code



$Opera_cloud_hotels=@()
$loopcounter=0
[decimal]$total_loop=$hotelIDs.count
foreach($hotelID in $hotelIDs)
{
write-host "$hotelID" -BackgroundColor Red
$loopcounter++
$percentage=($loopcounter/$total_loop)*100
write-host "Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor green
Start-Sleep -Milliseconds 5
$tokenURL_hotel_details = "https://api.example.com/sample-endpoint"
$tokenURL_hotel_isOpera=Invoke-RestMethod -Uri $tokenURL_hotel_details -Method GET -Headers $headers
$tokenURL_hotel_isOpera.hotelDetails.code
if($tokenURL_hotel_isOpera.hotelDetails.code -eq "OPERA" -and $tokenURL_hotel_isOpera.hotelDetails.category -eq 'PMS'){$Opera_cloud_hotels+=$tokenURL_hotel_isOpera.hotelDetails.hotelID}
}
$Opera_cloud_hotels


$loopcounter=0
[decimal]$total_loop=$Opera_cloud_hotels.count
foreach($hotelID in $Opera_cloud_hotels)
{
write-host "$hotelID" -BackgroundColor Red
$loopcounter++
$percentage=($loopcounter/$total_loop)*100
write-host "Percentage Completed :$([math]::Round($percentage,2)) %" -ForegroundColor green
$tokenURL_roomNumbers="https://api.example.com/sample-endpoint"
$roomNumbers=Invoke-RestMethod -Uri $tokenURL_roomNumbers -Method GET -Headers $headers
$total_room=($roomNumbers.listOfValues.items).count
$double_room=$roomNumbers.listOfValues.items | Where-Object { $_.name -eq "DOUBLE" -and $_.active -eq "True"}
if($double_room -is [array]){$double_room_count=$double_room.count}elseif($double_room -eq $null){$double_room_count=0}else{$double_room_count=1}
Start-Sleep -Milliseconds 5

    $master_table+=[pscustomobject]@{
    'Hotel Code'=$hotelID
    'Room Type'="Double";
    'Physical Rooms'=$double_room_count
    }
}

$stopwatch.stop()
$stopwatch.Elapsed


$response = [PSCustomObject]@{
    Status     = "Success"
    Timestamp  = (Get-Date).ToString("s")
    Operation  = $SelectedOperation   
    PropertyID = $HotelPropertyID      
    Message    = "success"
}

$response | ConvertTo-Json -Depth 3