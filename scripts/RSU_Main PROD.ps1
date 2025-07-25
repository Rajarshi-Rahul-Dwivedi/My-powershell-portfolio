# Check if 'python' is available in the PATH
$pythonCheck = Get-Command python -ErrorAction SilentlyContinue
$currentDate=(Get-Date).ToString("dd-MM-yyyy")

$loc="$PSScriptRoot\setup"
$retry_counter=0
cd $loc

function Setup-Environment {
    param(
        [int]$RetryCount = 0
    )

    Write-Host "`n--- Running Environment Setup (Attempt $($RetryCount + 1)) ---`n" -ForegroundColor Cyan

    $continue = $true

    try {
        # Check if Python is available
        $pythonCheck = Get-Command python -ErrorAction SilentlyContinue
        if (-not $pythonCheck) {
            Write-Warning "Python is not found in the system PATH. Please install Python and add it to your PATH."
            $continue = $false
        }

        # Check and install ImportExcel module
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "Installing 'ImportExcel' module..." -ForegroundColor Yellow
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Install-PackageProvider -Name NuGet -Scope CurrentUser -Force
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
            Import-Module ImportExcel -Force
        } else {
            Write-Host "'ImportExcel' module is already installed." -ForegroundColor Green
        }

        # Ensure Python Scripts path is in PATH
        $pythonScriptsPath = "$env:APPDATA\Python\Python312\Scripts"
        $oldPath = [Environment]::GetEnvironmentVariable("PATH", "User")

        if ($oldPath -notlike "*$pythonScriptsPath*") {
            $newPath = "$oldPath;$pythonScriptsPath"
            [Environment]::SetEnvironmentVariable("PATH", $newPath, "User")
            Write-Host "Added Python Scripts path to PATH: $pythonScriptsPath" -ForegroundColor Green
        } else {
            Write-Host "Python Scripts path already in PATH." -ForegroundColor Green
        }

        # Check for required Python packages
        $requiredPackages = @(
            "azure-keyvault-secrets",
            "azure-identity",
            "openpyxl",
            "pandas"
        )

        foreach ($package in $requiredPackages) {
            $check = python -m pip show $package 2>$null
            if (-not $check) {
                Write-Host "Installing missing Python package: $package" -ForegroundColor Yellow
                python -m pip install --user $package
            } else {
                Write-Host "Python package '$package' is already installed." -ForegroundColor Green
            }
        }

    } catch {
        Write-Error "An error occurred during environment setup: $_"
        $continue = $false
    }

    if (-not $continue) {
        $choice = Read-Host "Setup failed. Do you want to retry? (Y/N)"
        if ($choice -match '^(y|Y)') {
            Setup-Environment -RetryCount ($RetryCount + 1)
        } else {
            Write-Error "Exiting setup."
            exit 1
        }
    } else {
        Write-Host "`nEnvironment setup completed successfully!" -ForegroundColor Green
    }
}

#clearing up the input files if exist
Remove-Item "$loc\Restaurants Closure Process\input\*"
if(Test-Path "$loc\Restaurants Closure Process\output\$($currentDate)"){
Rename-item "$loc\Restaurants Closure Process\output\$($currentDate)" -NewName "$((Get-Date).ToString("dd-MM-yyyy-ss-ff"))"
}

function Validate-MealPackage{
[CmdletBinding()]
    param (
        $str1
    )
    
    $Validate_mealPackage= Import-Excel "$loc\Restaurants Closure Process\output\$currentDate\outputFile_removePackage.xlsx"

    $invalidStatus = $Validate_mealPackage | Where-Object { $_.'Status Code' -ne 200 } | Select-Object -ExpandProperty 'Confirmation Number'
    if ($invalidStatus.Count -eq 0) {
        Write-Output "All status codes are 200"
        Return 200
    } else {
        Write-Warning "Failed to Remove packages for following confirmation numbers :`n$($invalidStatus -join ', ')`n`nPress any key to Retry";pause;$retry_counter++
        if($retry_counter -gt 5){Write-Warning "Unable to complete for mentioned Confirmation Number moving to next step";pause;return 100}
            $retrysheet = @()
            foreach($conf in  $invalidStatus){      
                    $retrysheet += [PSCustomObject]@{
                    'Hotel ID'            = $hotel
                    'Confirmation Number' = $conf 
                    }
            }
            if(Test-Path $remove_meal_package_input){Remove-Item -Path $remove_meal_package_input}
            $retrysheet | Export-Excel -Path $remove_meal_package_input -WorksheetName "property_code" -TableStyle Medium7 -AutoSize
            $sheet2 | Export-Excel -Path $remove_meal_package_input -WorksheetName "package_code" -TableStyle Medium7 -AutoSize
            cd "$loc\Restaurants Closure Process"
            #python .\removeMealPackage.py PROD
            Start-Sleep -Seconds 20
            Validate-MealPackage
            
    }


}

function Validate-status{
[CmdletBinding()]
    param (
        $file,
        $validCode
    )
    
    $Validate_status= Import-Excel $file

    $invalidStatus = $Validate_status | Where-Object { $_.'Status Code' -ne $validCode } | Select-Object -ExpandProperty 'Confirmation Number'
    if ($invalidStatus.Count -eq 0) {
        Write-Output "All status codes are $validCode"
        Return "VALID"
    } else {
        Write-Warning "Failed for following confirmation numbers :`n$($invalidStatus -join ', ')`n"
        pause
        }


}
Setup-Environment
# Load JSON file
$config = Get-Content -Path "$loc\setup.json" | ConvertFrom-Json

# Assign values to variables
$hotel     = $config.hotel
$start_date = Get-Date $config.start_date
$end_date   = Get-Date $config.end_date
$username  = $config.username
$password  = $config.password
$cashierID = $config.cashierID

$package_list=gc "$loc\Package_discription.txt"
$intial_setup=@()
foreach($Confno in (GC "$loc\Confirmation Numbers.txt"))
{
        $intial_setup += [PSCustomObject]@{
                    'Hotel ID'            = $hotel;
                    'Confirmation Number' = $Confno;
                    }
}
Write-Host "Gathering Confirmation Numbers details for $hotel" -ForegroundColor Yellow
$intial_setup | Export-Excel -Path "$($loc)\Restaurants Closure Process\input\Cardtoken Details.xlsx" -WorksheetName "HotelCode" -TableStyle Medium7 -AutoSize
cd "$loc\Restaurants Closure Process"
python .\get_product.py PROD
    Rename-Item "$loc\Restaurants Closure Process\output\$currentDate\ReservationStatusAndPackages.xlsx" -NewName "ReservationStatusAndPackages_initial.xlsx"
    $product_details=Import-Excel -Path "$loc\Restaurants Closure Process\output\$currentDate\ReservationStatusAndPackages_initial.xlsx"
    Rename-Item "$($loc)\Restaurants Closure Process\input\Cardtoken Details.xlsx" -NewName "All Confirmations.xlsx"

#$intial_details=Import-Excel -Path  "$loc\EXPORT.xlsx"
$intial_details=$product_details

# Processed results
$sheet1 = @()
$stayoverguest=@()
$productSet = New-Object System.Collections.Generic.HashSet[string]

# Remove duplicates
$seenConfirmations = @{}

foreach ($res in $intial_details) {
    $truncBegin = Get-Date $res.'Arrival Date'
    $truncEnd = Get-Date $res.'Departure Date'
    $products = $res.PRODUCTS -split ','
    $products = $products | ForEach-Object { $_.Trim() }

    if (($truncBegin -ge $start_date -and $truncBegin -le $end_date) -and ($truncEnd -ge $start_date -and $truncEnd -le $end_date)) {
        if (-not $seenConfirmations.ContainsKey($res.'Confirmation No.')) {
            # Check if at least one product is in allowed package list
            if ($products | Where-Object { $_ -in $package_list }) {
                # Add confirmation number to Sheet1
                $sheet1 += [PSCustomObject]@{
                    'Hotel ID'            = $hotel
                    'Confirmation Number' = $res.'Confirmation No.' 
                }
                # Add to seen
                $seenConfirmations[$res.'Confirmation No.'] = $true
            }
        }
        # Add all valid product codes to the product set
        foreach ($p in $products) {
            if ($p -in $package_list) {
                $productSet.Add($p) | Out-Null
            }
        }
    }
    elseif($products | Where-Object { $_ -in $package_list }){
        $stayoverguest += [PSCustomObject]@{
                    'Hotel ID'            = $hotel;
                    'Confirmation Number' = $res.'Confirmation No.';
                    'Arival Date'=$truncBegin;
                    'Departure Date'=$truncEnd;
                    'Products' =$res.'Products'                               
                }
            }
}

# Sheet2 - Distinct product codes
$sheet2 = $productSet | ForEach-Object {
    [PSCustomObject]@{ 'Package Code' = $_ }
}

# Export to Excel
$Rate_dettails_input = "$($loc)\Restaurants Closure Process\input\Reservation Details.xlsx"
$remove_meal_package_input="$($loc)\Restaurants Closure Process\input\Remove meal package.xlsx"

$stayoverguest | Export-Excel -Path "$($loc)\Stayover Guests $($hotel).xlsx" -WorksheetName "sheet1" -TableStyle Medium7 -AutoSize
$sheet1 | Export-Excel -Path $Rate_dettails_input -WorksheetName "sheet1" -TableStyle Medium7 -AutoSize
$sheet1 | Export-Excel -Path $remove_meal_package_input -WorksheetName "property_code" -TableStyle Medium7 -AutoSize
$sheet2 | Export-Excel -Path $remove_meal_package_input -WorksheetName "package_code" -TableStyle Medium7 -AutoSize
write-host "Stayover Reservations: $($stayoverguest.count)"




cd "$loc\Restaurants Closure Process"
python .\getRateDetails.py PROD
    #python .\getCommunication.py PROD
    Rename-Item "$loc\Restaurants Closure Process\output\$currentDate\outputFilegetRateDetails.xlsx" -NewName "outputFilegetRateDetails_before_removingPackage.xlsx"
    $Rate_details_BeforeMeal=Import-Excel -Path "$loc\Restaurants Closure Process\output\$currentDate\outputFilegetRateDetails_before_removingPackage.xlsx"
    
Write-Host "`nRemoving Meal Packages now for $($sheet1.count) Reservations" -ForegroundColor Yellow
if($Rate_details_BeforeMeal.Count -lt 1){Write-Warning "No eligible refund on the property";pause;exit}
pause
python .\removeMealPackage.py PROD
    Validate-MealPackage

python .\getRateDetails.py PROD
    Rename-Item "$loc\Restaurants Closure Process\output\$currentDate\outputFilegetRateDetails.xlsx" -NewName "outputFilegetRateDetails_after_removingPackage.xlsx"
    $Rate_details_After_meal=Import-Excel -Path "$loc\Restaurants Closure Process\output\$currentDate\outputFilegetRateDetails_after_removingPackage.xlsx"



$Remove_Meal_packages=@()
$Negative_Validations = @()
$get_cardTokens = @()


foreach ($before in $Rate_details_BeforeMeal) {
    $after = $Rate_details_After_meal | Where-Object { $_.'Confirmation No.' -eq $before.'Confirmation No.' }

    if (-not $after) {
        Write-Warning "No matching 'After meal package' entry found for Confirmation No. $($before.'Confirmation No.')"
        continue
    }

    $packageGross       = [math]::Round($before.'Package Cost' * 1.2, 2)
    $packageAfterGross  = [math]::Round($after.'Package Cost' * 1.2, 2)
    $packageDifference  = [math]::Round($packageGross - $packageAfterGross, 2)
    $packageDifference  = [math]::Round(($packageGross - $packageAfterGross), 2)
    $refund_amount      = $after.'Guest Pay'
    if(($before.'Guest Pay' - $packageDifference) -lt 0){
        if($after.'Guest Pay' -ge (-$packageDifference)){
            $refund_amount      = $after.'Guest Pay'    
            }else{$refund_amount=-$packageDifference}
        if(($refund_amount- $after.'Guest Pay') -le 0.09 -and ($refund_amount- $after.'Guest Pay') -ge -0.09){$refund_amount = $after.'Guest Pay'}        
        $validation = [math]::Round(($before.'Guest Pay'  - $packageDifference) - $after.'Guest Pay', 2)
       
    }
    if($after.'Guest Pay' -ge 0){$refund_amount=0}
    if($packageDifference -le 0){continue}

    # Check deposit
    if ($after.Deposit -ne 0) {
        if (($validation -le 0.09 -and $validation -ge -0.09) -and  $after.'Guest Pay' -lt 0){ 
            $Remove_Meal_packages += [PSCustomObject]@{
                'Hotel ID'             = $after.'Hotel ID';
                'Confirmation Number'  = $after.'Confirmation No.';
                'Reservation ID'       = $after.'Reservation ID';
         'Total Cost of Stay before'   = $before.'Total Cost Of Stay';
          'Total Cost of Stay After'   = $after.'Total Cost Of Stay' ;
                'Deposit Paid before'  = $before.Deposit
                'Deposit Paid after'   = $after.Deposit;
                'Package Cost Before'  = $packageGross;
                'Package Cost After'   = $packageAfterGross;
                'Guest Pay Before'     = $before.'Guest Pay';
                'Guest Pay After'      = $after.'Guest Pay';
                'Package Difference'   = $packageDifference;
                'Validation'           = $validation
                'Refund Amount'        = $refund_amount
            }
            $get_cardTokens += [PSCustomObject]@{
                'Hotel ID'            = $after.'Hotel ID';
                'Confirmation Number' = $after.'Confirmation No.';
                'Reservation ID'      = $after.'Reservation ID';
                'Deposit Paid'        = $refund_amount
                }
        }else{
            $Negative_Validations += [PSCustomObject]@{
                'Hotel ID'             = $after.'Hotel ID';
                'Confirmation Number'  = $after.'Confirmation No.';
                'Reservation ID'       = $after.'Reservation ID';
         'Total Cost of Stay before'   = $before.'Total Cost Of Stay';
          'Total Cost of Stay After'   = $after.'Total Cost Of Stay' ;
                'Deposit Paid before'  = $before.Deposit
                'Deposit Paid after'   = $after.Deposit;
                'Package Cost Before'  = $packageGross;
                'Package Cost After'   = $packageAfterGross;
                'Guest Pay Before'     = $before.'Guest Pay';
                'Guest Pay After'      = $after.'Guest Pay';
                'Package Difference'   = $packageDifference;
                'Validation'           = $validation
                'Refund Amount'        = $refund_amount
            }
        }
    }
}
$Remove_Meal_packages | Export-Excel -Path "$($loc)\RemoveMealPackage_ValidationReport.xlsx" -WorksheetName "sheet2" -TableStyle Medium7 -AutoSize -Show
$Negative_Validations | Export-Excel -Path "$($loc)\RemoveMealPackage_ValidationReport_negative.xlsx" -WorksheetName "sheet2" -TableStyle Medium7 -AutoSize 
#if(Test-Path "$($loc)\Restaurants Closure Process\input\Reservation Details.xlsx"){Rename-Item "$($loc)\Restaurants Closure Process\input\Reservation Details.xlsx" -NewName  "Reservation Details_initial.xlsx"}

$get_cardTokens | Export-Excel -Path "$($loc)\Restaurants Closure Process\input\Cardtoken Details.xlsx" -WorksheetName "HotelCode" -TableStyle Medium7 -AutoSize
cd "$loc\Restaurants Closure Process"
python .\getCardTokenDetails.py PROD
    $output_cardDetails=Import-Excel -Path "$($loc)\Restaurants Closure Process\Output\$($currentDate)\outputFile_getCardTokenDetails.xlsx"




$Post_payment_input = @()

foreach ($refund in $output_cardDetails) {
    # Find the matching package entry with near-equal Guest Pay and Refund Amount
    $match = $Remove_Meal_packages | Where-Object {
        $_.'Hotel ID' -eq $refund.'Property Code' -and
        $_.'Reservation ID' -eq $refund.'Reservation Id' -and
        [math]::Abs($_.'Refund Amount' - $refund.'Refund Amount') -lt 0.05
    } | Select-Object -First 1

    if ($match) {
        
            $Post_payment_input += [PSCustomObject]@{
                'Hotel ID'            = $match.'Hotel ID'
                'Reservation Number'  = $refund.'Reservation Id'
                'Amount'              = $refund.'Refund Amount'
                'Currency Code'       = $refund.'Currency Code'
                'Reservation Type'    = $refund.'Reservation Type'
                'Payment Method'      = $refund.'Payment Method'
                'Token'               = $refund.'Token'
                'Expiry Date'         = $refund.'Expiry Date'
                'Card Holder Name'    = $refund.'Card Holder Name'
                'Reference'           = 'Temporary Restaurant Closure'
                'Suppliment'          = 'Temporary Restaurant Closure'
                'Cashier ID'          = $cashierID
            }
        
    }
    else {
        Write-Warning "No matching refund found for Reservation ID: $($refund.'Reservation Id')"
    }
}


#add status check based on reservation number
 
python .\get_product.py PROD
    $product_details=Import-Excel -Path "$loc\Restaurants Closure Process\output\$currentDate\ReservationStatusAndPackages.xlsx"
    $nonReservedrevid = $product_details | Where-Object { $_.'Check-in Status' -ne "Reserved" } | Select-Object -ExpandProperty 'Reservation ID'
    $Post_payment_input = $Post_payment_input | Where-Object { $nonReservedrevid -notcontains $_.'Reservation Number' } 
    $get_cardTokens=$get_cardTokens | Where-Object { $nonReservedrevid -notcontains $_.'Reservation ID' }
    if($nonReservedrevid -ne $null){Write-Warning "Reservation status is Not-reserved for following Reservation ID : `n$nonReservedrevid `nAbove IDs will be removed from the payment process";pause}
    $product_details | format-table

$Post_payment_input | Export-Excel -Path "$($loc)\Restaurants Closure Process\input\Post Payment.xlsx" -WorksheetName "Reservation" -TableStyle Medium7 -AutoSize
write-host "`nProceeding to Payments Now for $($Post_payment_input.count) Reservations" -ForegroundColor Yellow


cd "$loc\Restaurants Closure Process"
pause
python postPayment.py PROD
    Validate-status "$($loc)\Restaurants Closure Process\Output\$($currentDate)\outputFile_postPayment.xlsx" 201

$output_Postpayment=Import-Excel "$($loc)\Restaurants Closure Process\Output\$($currentDate)\outputFile_postPayment.xlsx"
$unalloacted_deposit_input=@()
foreach ($deposit in $output_Postpayment){
    
    if($deposit.'Status Code' -ne 201){write-warning "Failed Status for $($deposit.'Status Code')";pause;Continue}
        $unalloacted_deposit_input += [PSCustomObject]@{
                'Proprty Code'            = $deposit.'Hotel ID'
                'Confirmation number'     = $deposit.'Confirmation Number'
                'Transaction Number'      = $deposit.'Transacion No'
            }
}

$unalloacted_deposit_input  | Export-Excel -Path "$($loc)\Restaurants Closure Process\input\Unallocated Deposit Mapping.xlsx" -WorksheetName "Sheet1" -TableStyle Medium7 -AutoSize

python unallocatedDepositMapping.py PROD
    $unalloacted_deposit_output="$($loc)\Restaurants Closure Process\Output\$($currentDate)\outputFile_postUnallocatedDeposit.xlsx"
    Validate-status $unalloacted_deposit_output 200

if(test-path $Rate_dettails_input){Rename-Item $Rate_dettails_input -NewName "Reservation Details_beforeRefund.xlsx"}
$get_cardTokens | Export-Excel -Path $Rate_dettails_input -WorksheetName "sheet1" -TableStyle Medium7 -AutoSize

python .\getRateDetails.py PROD
    Rename-Item "$loc\Restaurants Closure Process\output\$currentDate\outputFilegetRateDetails.xlsx" -NewName "outputFilegetRateDetails_after_Refund.xlsx"
    $Rate_details_AfterRefund=Import-Excel -Path "$loc\Restaurants Closure Process\output\$currentDate\outputFilegetRateDetails_after_Refund.xlsx"

$valid_refunds_counter=0
foreach($item in $Rate_details_AfterRefund)
{
    $before_refund = $Rate_details_BeforeMeal | Where-Object { $_.'Confirmation No.' -eq $item.'Confirmation No.' }
    if($item.'Guest Pay' -eq 0){Write-Host "$($item.'Confirmation No.') Guest Pay is 0 : VALID" -ForegroundColor Green;$valid_refunds_counter++}
    elseif([math]::Round($item.'Guest Pay',0) -eq [math]::Round($before_refund.'Guest Pay',0)){Write-Host "$($item.'Confirmation No.') Guest Pay is same as Guest pay amount before removing package  $($item.'Guest Pay') : VALID" -ForegroundColor Green;$valid_refunds_counter++}
    else{Write-Host "$($item.'Confirmation No.') Guest Pay amount is $($item.'Guest Pay') please check : INVALID" -ForegroundColor Red;pause}
}
write-host "`nProcess Complete`n" -ForegroundColor Green

Write-host "
Property: $($hotel)
Total Reservation: $($intial_setup.count)
Reservation with breakfast package: $($sheet1.count)
Eligible Refunds: $($valid_refunds_counter)
" -ForegroundColor Cyan
pause
$documentaion=
"
1. Get Reservation Details
Source: OHIP API via Python script (get_product.py)

Fields Extracted: Resort, DATE_BEGIN, DATE_END, CONFIRMATION_NO, PRODUCTS

Output: ReservationStatusAndPackages.xlsx

2. Filter Response
Remove duplicates by CONFIRMATION_NO

Keep only products listed in Package_discription.txt (master list)

Create Excel Files:

sheet1: Hotel ID + Confirmation Number

sheet2: Distinct valid products

Output: Remove meal package.xlsx & Reservation Details.xlsx

3. Get Rate Details (Before Removal)
Run: getRateDetails.py

Validate status codes (expecting 200)

Output: outputFilegetRateDetails_before_removingPackage.xlsx

4. Get Communication Details
Run: getCommunication.py

Validate response

5. Remove Meal Packages
Run: removeMealPackage.py

Validate via function Validate-MealPackage

Retry up to 5 times if status codes are not 200

6. Get Rate Details (After Removal)
Run: getRateDetails.py

Output: outputFilegetRateDetails_after_removingPackage.xlsx

7. Compare Before vs After
Match on Confirmation No.

Calculate:

Package Gross = Before * 1.19

Package Cost After (gross)

Difference + Guest Pay After = Validation

Output:

RemoveMealPackage_ValidationReport.xlsx

RemoveMealPackage_ValidationReport_negative.xlsx

8. Create Input for Card Details
Fields:

Hotel ID

Confirmation Number

Reservation ID

Refund Amount (from validation logic)

Output: Cardtoken Details.xlsx

9. Get Card Details
Run: GetCardTokenDetails.py

Validate status

Output: outputFile_getCardTokenDetails.xlsx

10. Create Input for Payment Posting
Combine info from step 7 & 9

Fields:

Hotel ID, Reservation Number, Amount, Currency Code

Payment Method, Token, Expiry Date, Card Holder Name

Reference = “"Temporary Restaurant Closure”"

Cashier ID = from JSON config

Output: Post Payment.xlsx

11. Post Payment
Run: postPayment.py

Validate using Validate-status

Output: outputFile_postPayment.xlsx

12. Create Input for Deposit Mapping
Extract:

Property Code

Confirmation Number

Transaction Number

Output: Unallocated Deposit Mapping.xlsx

13. Unallocated Deposit Mapping
Run: unallocatedDepositMapping.py

Validate status

14. Validate Payments
Check for any payments with non-201 status

Manual pause if any issues

15. Get Rate Details (Post-Payment)
Run: getRateDetails.py

Output: outputFilegetRateDetails_after_Refund.xlsx

16. Validate Final Guest Pay
Ensure Guest Pay = 0

Output: List of Confirmation Numbers where validation failed
"