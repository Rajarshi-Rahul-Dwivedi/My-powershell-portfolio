#Hello, This Script will not function in it's current form please read the discription to update the script acordingly or contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281

$loc=$PSScriptRoot
mkdir "$loc\output" -ErrorAction SilentlyContinue
Remove-Item "$loc\outputs\*.txt" -ErrorAction SilentlyContinue

function execute-sql {
    [CmdletBinding()]
    param (
        [string]$sys_password
    )
    "conn SID_NAME/$sys_password@SID_NAME
    set newpage 0; 
    set echo off;
    SET LINESIZE 32767;
    SET TRIMSPOOL ON; 
    set feedback off;
    SET WRAP OFF; 
    set pagesize 0 ;
    @""$loc\scripts\scripts\Country.sql""
    @""$loc\scripts\ACCOUNT_TYPES.sql""
    @""$loc\scripts\ADDRESS_TYPES.sql""
    @""$loc\scripts\ADJUSTMENT_CODES.sql""
    @""$loc\scripts\AIRPORTS.sql""
    @""$loc\scripts\ALERT_MESSAGES.sql""
    @""$loc\scripts\AMENITIES.sql""
    @""$loc\scripts\ARTICLES.sql""
    @""$loc\scripts\AR_ACCOUNT_TYPES.sql""
    @""$loc\scripts\ATTRACTION_CATEGORIES.sql""
    @""$loc\scripts\ATTRACTION_CODES.sql""
    @""$loc\scripts\AUTO_SETTLEMENT_FOLIO_TYPES.sql""
    @""$loc\scripts\BANK_ACCOUNTS.sql""
    @""$loc\scripts\BED_TYPES.sql""
    @""$loc\scripts\BIRTH_COUNTRY.sql""
    @""$loc\scripts\BUILDINGS.sql""
    @""$loc\scripts\BUSINESS_BLOCK_TYPES.sql""
    @""$loc\scripts\BUSINESS_SEGMENTS.sql""
    @""$loc\scripts\CANCELLATION_REASONS.sql""
    @""$loc\scripts\CANCEL_PENALTY.sql""
    @""$loc\scripts\CANCEL_PENALTY_SCHEDULES.sql""
    @""$loc\scripts\CASHIERS.sql""
    @""$loc\scripts\COMMISSION_CODES_RESERVATION.sql""
    @""$loc\scripts\COMMISSION_CODES_REVENUE.sql""
    @""$loc\scripts\COMMUNICATION_TYPES.sql""
    @""$loc\scripts\COMPANY_TYPES.sql""
    @""$loc\scripts\COUNTRY_ENTRY_POINT.sql""
    @""$loc\scripts\CURRENCY_EXCHANGE.sql""
    @""$loc\scripts\DEPARTMENT.sql""
    @""$loc\scripts\DEPOSIT_RULES.sql""
    @""$loc\scripts\DEPOSIT_RULE_SCHEDULE.sql""
    @""$loc\scripts\DESTINATION_CODES.sql""
    @""$loc\scripts\DISCOUNT_REASONS.sql""
    @""$loc\scripts\DISPLAY_SETS.sql""
    @""$loc\scripts\EVENT_CODES.sql""
    @""$loc\scripts\FLOOR.sql""
    @""$loc\scripts\FOLIO_ARRAGEMENT_CODES.sql""
    @""$loc\scripts\FOREIGN_CURRENCY_CODES.sql""
    @""$loc\scripts\GLOBAL_ALERT_DEFINITIONS.sql""
    @""$loc\scripts\GROUP_ARRAGEMENT_CODES.sql""
    @""$loc\scripts\GUEST_MESSAGES.sql""
    @""$loc\scripts\GUEST_STATUS.sql""
    @""$loc\scripts\GUEST_TYPES.sql""
    @""$loc\scripts\HK_ATTENDANT.sql""
    @""$loc\scripts\HK_SECTION_CODE.sql""
    @""$loc\scripts\IDENTIFICATION_COUNTRY.sql""
    @""$loc\scripts\IDENTIFICATION_TYPES.sql""
    @""$loc\scripts\IMMIGRATION_STATUS.sql""
    @""$loc\scripts\INACTIVE_REASONS.sql""
    @""$loc\scripts\ITEM_CLASS.sql""
    @""$loc\scripts\ITEM_INVENTORY.sql""
    @""$loc\scripts\LANGUAGE.sql""
    @""$loc\scripts\LOCATORS.sql""
    @""$loc\scripts\LOST_REASONS.sql""
    @""$loc\scripts\MALING_ACTION_CODES.sql""
    @""$loc\scripts\MARKETING_CITIES.sql""
    @""$loc\scripts\MARKETING_REGIONS.sql""
    @""$loc\scripts\MARKET_CODES.sql""
    @""$loc\scripts\MARKET_GROUPS.sql""
    @""$loc\scripts\NATIONALITIES.sql""
    @""$loc\scripts\NOTE_TYPES.sql""
    @""$loc\scripts\NO_SHOW_POSTING_RULES.sql""
    @""$loc\scripts\ORIGIN_CODES.sql""
    @""$loc\scripts\OUT_OF_ORDER_AND_OUT_OF_SERVICE.sql""
    @""$loc\scripts\PACKAGES_FORECAST_GROUP.sql""
    @""$loc\scripts\PACKAGE_CODES.sql""
    @""$loc\scripts\PAYMENT_TYPES.sql""
    @""$loc\scripts\PREFERENCES.sql""
    @""$loc\scripts\PREFERENCE_GROUPS.sql""
    @""$loc\scripts\PRICING_SCHEDULES_DAILY_RATES.sql""
    @""$loc\scripts\PRICING_SCHEDULES_STANDARD_RATES.sql""
    @""$loc\scripts\PROMOTION_CODES.sql""
    @""$loc\scripts\PROMOTION_GROUP.sql""
    @""$loc\scripts\PROPERTIES.sql""
    @""$loc\scripts\PROPERTY_DETAILS.sql""
    @""$loc\scripts\PURPOSE_OF_STAY.sql""
    @""$loc\scripts\RATE_CATEGORY.sql""
    @""$loc\scripts\RATE_CLASS.sql""
    @""$loc\scripts\RATE_CODES.sql""
    @""$loc\scripts\RATE_OVERRIDE_REASONS.sql""
    @""$loc\scripts\REFUSED_REASONS.sql""
    @""$loc\scripts\REGION.sql""
    @""$loc\scripts\RESERVATION_METHOD.sql""
    @""$loc\scripts\RESERVATION_TYPES.sql""
    @""$loc\scripts\RESERVATION_TYPES_SCHEDULES.sql""
    @""$loc\scripts\RESTRICTED_REASONS.sql""
    @""$loc\scripts\ROOMS.sql""
    @""$loc\scripts\ROOM_CLASS.sql""
    @""$loc\scripts\ROOM_CONDITIONS.sql""
    @""$loc\scripts\ROOM_FEATURES.sql""
    @""$loc\scripts\ROOM_HEIRARCHY_CLASS.sql""
    @""$loc\scripts\ROOM_HEIRARCHY_TYPES.sql""
    @""$loc\scripts\ROOM_HIERARCHY_CLASSES.sql""
    @""$loc\scripts\ROOM_HIERARCHY_TYPES.sql""
    @""$loc\scripts\ROOM_MAINTAINANCE.sql""
    @""$loc\scripts\ROOM_MOVE_REASONS.sql""
    @""$loc\scripts\ROOM_MOVE_REASONS.txt""
    @""$loc\scripts\ROOM_TYPES.sql""
    @""$loc\scripts\ROUTING_CODES.sql""
    @""$loc\scripts\SCHEDULES.sql""
    @""$loc\scripts\SOURCE_CODES.sql""
    @""$loc\scripts\SOURCE_GROUPS.sql""
    @""$loc\scripts\SPECIAL_REQUESTS.sql""
    @""$loc\scripts\STATUS_CODES.sql""
    @""$loc\scripts\STATUS_CODES_FLOW.sql""
    @""$loc\scripts\STOP_PROCESS_REASONS.sql""
    @""$loc\scripts\TASKS.sql""
    @""$loc\scripts\TAX_TYPES.sql""
    @""$loc\scripts\TITLE.sql""
    @""$loc\scripts\TRACE_TEXTS.sql""
    @""$loc\scripts\TRACKIT_ACTION.sql""
    @""$loc\scripts\TRACKIT_LOCATION.sql""
    @""$loc\scripts\TRACKIT_TYPES.sql""
    @""$loc\scripts\TRANSACTION_CODES.sql""
    @""$loc\scripts\TRANSACTION_GENERATES.sql""
    @""$loc\scripts\TRANSACTION_GROUPS.sql""
    @""$loc\scripts\TRANSACTION_SUBGROUPS.sql""
    @""$loc\scripts\TRANSPORTATION.sql""
    @""$loc\scripts\TURNAWAY_CODE.sql""
    @""$loc\scripts\VIP.sql""
    @""$loc\scripts\WAITLIST_CODES.sql""
    @""$loc\scripts\WAITLIST_PRIORITY.sql""
	exit" | sqlplus  /nolog

}


#Get-ChildItem "$loc" | foreach {"spool #location\$($_.Name) `n"+(gc $_.FullName)+"`nSpool off" | Set-Content $_.fullname}
#Get-ChildItem "$loc" | foreach {if($_.Extension -eq ".txt"){Rename-Item $_.FullName -NewName "$($_.BaseName).sql"}}

$oracle_password=Read-host "Please enter Opera password`n"
Get-ChildItem "$loc\Scripts" | foreach {(get-content $_.fullname) -replace "#location", "$loc\Output" | Set-Content $_.fullname }
execute-sql -sys_password $oracle_password

Write-Host "Creating one output file at $loc" -ForegroundColor Cyan
$master_output=@()
foreach($file in Get-ChildItem "$loc\scripts"){
    
    $sql_query=gc "$loc\scripts\$($file.name)" -raw;
    $sql_query=$sql_query -replace "spool.*(txt|off)" , ""

    $master_output+=[pscustomobject]@{
    'Script name'=$file.name;
    'Sql Query'=$sql_query;
    'Output'=gc "$loc\Output\$($file.name)" -raw;
    }

}
$output=$master_output | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
$output | Out-File "$loc\Combined output.csv"  