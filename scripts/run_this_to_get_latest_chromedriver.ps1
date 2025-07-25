#Script Created by Rajarshi fully functional
$date_today=get-date -Format 'dd-MMM'
function Get-ChromeVersion {
    $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    
    if (Test-Path $chromePath) {
        $chromeVersion = (Get-Item $chromePath).VersionInfo.FileVersion
        return $chromeVersion.Split('.')[0],$chromeVersion  # Return only the major version
    } else {
        Write-Host "Chrome is not installed at the expected path."
        return $null
    }
}


# Get the Chrome version
$chromeVersions = Get-ChromeVersion
$chromeMajorVersion=$chromeVersions[0]
$chromeFullVersion=$chromeVersions[1]
$loc="$PSScriptRoot"

$chromedriver_location="$loc\chromedriver.exe"
$Current_cd_version= &$chromedriver_location --version
$Current_cd_version=$Current_cd_version | Select-String -Pattern "ChromeDriver\s([\d\.]+)" | ForEach-Object { $_.Matches.Groups[1].Value }


if ($chromeMajorVersion -ne $null) {
    # Chromedriver base URL
    $chromeDriverBaseURL = "https://storage.googleapis.com/chrome-for-testing-public"

    # Get the latest chromedriver version that matches the installed Chrome version
    try {

        #$chromeDriverVersion = Invoke-RestMethod "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_$chromeMajorVersion"
        $chromeDriverVersion=$chromeFullVersion
        if($Current_cd_version -eq $chromeDriverVersion){Write-Warning "The latest chromedriver version is already installed $Current_cd_version";start-sleep -Seconds 5;exit}
        
        # Define the download URL for the matching chromedriver
        $chromeDriverDownloadURL = "$chromeDriverBaseURL/$chromeDriverVersion/win64/chromedriver-win64.zip"
        $downloadPath = "$loc\chromedriver.zip"
        $targetPath = $loc 
        

        #backup old chromedriver
        Rename-Item "$loc\chromedriver.exe" -NewName "OLD_$($date_today)_chromedriver.exe" -Force -ErrorAction SilentlyContinue
        
        # Download the chromedriver zip file
        try{Invoke-WebRequest -Uri $chromeDriverDownloadURL -OutFile $downloadPath}catch{write-warning "Unable to connect Google APIs";start-sleep -Seconds 5;exit}

        
        
        

        # Extract the chromedriver.exe from the zip
        New-Item -ItemType Directory "$loc\temp"
        Expand-Archive $downloadPath -DestinationPath "$loc\temp"
        $chromedriver_path=Get-ChildItem "$loc\*chromedriver.exe" -Recurse
        Move-Item $chromedriver_path.FullName -Destination $loc -Force

        # Clean up the zip file
        Remove-Item $downloadPath -Force
        Remove-Item "$loc\temp" -Recurse -Force

        # Verify that chromedriver is updated
        $chromedriverPath = Join-Path $targetPath "chromedriver.exe"
        if (Test-Path $chromedriverPath) {
            Write-Host "Chromedriver updated to version $chromeDriverVersion"
        } else {
            Write-Host "Failed to update Chromedriver."
        }
    } catch {
        Write-Host "Failed to get the Chromedriver version for Chrome version $chromeMajorVersion. Error: $_"
    }
}
start-sleep 10
