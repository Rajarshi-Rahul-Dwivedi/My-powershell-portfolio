#Hello, This Script will not function in it's current form please read the discription to update the script acordingly or contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281
Get-Process -Name chrome -ErrorAction SilentlyContinue | Stop-Process -Force
$loc="$PSScriptRoot"
#$mineral_data=ConvertFrom-Csv -Delimiter "," -InputObject (gc "$loc\Mineral_info.csv")
$mineral_list=Get-Content "$loc\Mineral_list.txt"
$date_today=get-date -Format 'dd-MMM'
$download_location="$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\downloads"
$mineral_data=@()

#Setting Up Selenium Web Driver
$workingPath = 'C:\Automated booking'
if (($env:Path -split ';') -notcontains $workingPath) {$env:Path += "$workingPath"}

Add-Type -Path "$($workingPath)\scripts\WebDriver.dll"
#Add-Type -Path "$($workingPath)\WebDriver.Support.dll"
#Import-Module "$($workingPath)\WebDriver.dll"


$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
#$ChromeOptions.AddArgument('start-maximized')
#$ChromeOptions.AcceptInsecureCertificates = $True

#loading profile
$ChromeOptions.AddArgument("--user-data-dir=$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\AppData\Local\Google\Chrome\User Data")
$ChromeOptions.AddArgument("--profile-directory=Default")
$ChromeOptions.addargument('profile-directory=Default')
#$chromeOptions.AddArgument("--disable-extensions")
#$chromeOptions.AddArgument("--disable-automation")
$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeOptions)
$actions = [OpenQA.Selenium.Interactions.Actions]::new($ChromeDriver)


#$name="Orthoclase"
foreach($name in $mineral_list){

$ChromeDriver.Navigate().GoToURL("https://www.mindat.org")
start-sleep -Seconds 1
$search_bar=$ChromeDriver.FindElementByXPath('//input[@name="search"]')
$search_bar.SendKeys("$name")
$search_bar.SendKeys([OpenQA.Selenium.Keys]::Return)
start-sleep -Seconds 1
$mindat_formula=$ChromeDriver.FindElementByXPath('//div[text()="Mindat Formula:"]/..//following-sibling::div//span')
$lustre=$ChromeDriver.FindElementByXPath('//div//a[text()="Lustre:"]/..//following-sibling::div')
$Hardness=$ChromeDriver.FindElementByXPath('//div[text()="Hardness:"]/..//following-sibling::div')
$Cleavage=$ChromeDriver.FindElementByXPath('//div[text()="Cleavage:"]/..//following-sibling::div')
$streak=$ChromeDriver.FindElementByXPath('//div[text()="Streak:"]/..//following-sibling::div')
$Density=$ChromeDriver.FindElementByXPath('//div[text()="Density:"]/..//following-sibling::div')
$Type=$ChromeDriver.FindElementByXPath('//div[text()="Type:"]/..//following-sibling::div')
$RI_values=$ChromeDriver.FindElementByXPath('//div[text()="RI values:"]/..//following-sibling::div')
$Birefringence=$ChromeDriver.FindElementByXPath('//div[text()="Max Birefringence:"]/..//following-sibling::div')
$Dispersion=$ChromeDriver.FindElementByXPath('//div[text()="Dispersion:"]/..//following-sibling::div')
$Surface_Relief=$ChromeDriver.FindElementByXPath('//div[text()="Surface Relief:"]/..//following-sibling::div')


$mineral_data+=[pscustomobject]@{
'Mineral Name'=$name;
'Mindat Formula'=$mindat_formula.Text;
'Lustre'=$lustre.Text;
'Hardness'=$Hardness.Text;
'Cleavage'=(($Cleavage.Text) -replace "`n","," -replace "`r","");
'streak'=$streak.Text;
'Density'=$Density.Text;
'Type'=$Type.Text;
'RI values'=$RI_values.Text
'Max Birefringence'=(($Birefringence.Text) -split '\r?\n' | Select-Object -First 1);
'Surface Relief'=$Surface_Relief.Text;
'Dispersion'=$Dispersion.Text;
}
$mineral_data | Select-Object -Last 2
Start-Sleep -Seconds 1
}
#Write-Host "$name :`nFormula: $($mindat_formula.text) `nlustre: $($lustre.Text)`nTransparency: $($Transparency.Text)" -ForegroundColor Cyan


$output_csv=$mineral_data | ConvertTo-Csv -Delimiter ',' -NoTypeInformation
$output_csv | Out-File "$loc\Mineral_info_2.csv"
pause
$ChromeDriver.Quit()


<#
Mindat Formula:
Lustre
Hardness
Cleavage
streak
Density
optical data - Type,RI values: ,Max Birefringence:, Surface Relief: ,Dispersion:
#>