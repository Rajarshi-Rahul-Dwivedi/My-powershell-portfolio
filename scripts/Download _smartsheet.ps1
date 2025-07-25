$workingPath = ""
$download_location="$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\downloads"
if (($env:Path -split ';') -notcontains $workingPath) {$env:Path += "$workingPath"}

Add-Type -Path "$($workingPath)\scripts\WebDriver.dll"
#Add-Type -Path "$($workingPath)\WebDriver.Support.dll"
#Import-Module "$($workingPath)\WebDriver.dll"


$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$ChromeOptions.AddArgument('start-maximized')
$ChromeOptions.AcceptInsecureCertificates = $True
#$ChromeOptions.AddArgument("C:\Games\selenium\profile")
$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeOptions)
$ChromeDriver.Navigate().GoToURL("https://app.smartsheet.com/")

$ChromeDriver.FindElementByXPath('//*[@id="loginEmail"]').SendKeys("")
$ChromeDriver.FindElementByXPath('//*[@id="formControl"]').click()
$ChromeDriver.FindElementByXPath('//div//button[text()=" Sign in with email and password"]').click()
$ChromeDriver.FindElementByXPath('//*[@id="loginPassword"]').SendKeys("")
$ChromeDriver.FindElementByXPath('//*[@id="formControl"]').click()
Start-Sleep -Seconds 3

#removing any old smartsheet
try{
remove-item "$download_location\Property Contact List.xlsx"
remove-item "$download_location\Property Contact List.xlsx"
#remove-item "$download_location\19c Database Upgrade - Property Contact List.xlsx"
}catch{echo "No old smartsheet to remove"}

#download POS
do{
$ChromeDriver.Navigate().GoToURL("")
Start-Sleep -Seconds 10
$ChromeDriver.FindElementByXPath('//span[text()="File"]').Click()
$ChromeDriver.FindElementByXPath('//table//td[text()="Export"]').Click()
$ChromeDriver.FindElementByXPath('//td[@class="clsStandardMenuText " and text()="Export to Microsoft Excel"]').Click()
Start-Sleep -Seconds 15
}while(!(test-path "$download_location\Property Contact List.xlsx"))

#Download PMS
do{
$ChromeDriver.Navigate().GoToURL("")
Start-Sleep -Seconds 10
$ChromeDriver.FindElementByXPath('//span[text()="File"]').Click()
$ChromeDriver.FindElementByXPath('//table//td[text()="Export"]').Click()
$ChromeDriver.FindElementByXPath('//td[@class="clsStandardMenuText " and text()="Export to Microsoft Excel"]').Click()
Start-Sleep -Seconds 15
}while(!(test-path "$download_location\Property Contact List.xlsx"))

$ChromeDriver.Quit()


try{
Import-Excel -Path "$download_location\Property Contact List.xlsx" | Export-Csv -path "$workingPath\PMS\Meta Data\PMS_smartsheet.csv" -NoTypeInformation
Import-Excel -Path "$download_location\Property Contact List.xlsx" | Export-Csv -path "$workingPath\POS\For Dashboard\Meta data\POS_smartsheet.csv" -NoTypeInformation
}catch{write-warning "unable to save CSV file";return 1}
return 0
