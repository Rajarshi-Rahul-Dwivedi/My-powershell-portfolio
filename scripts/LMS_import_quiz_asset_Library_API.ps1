$dat=get-date -UFormat "%M-%d-%b"
#Created by Rajarshi
Add-Type -AssemblyName System.Net.Http
$loc=$PSScriptRoot
$table=@()
#URL=https://lmswebap.netlify.app/courseStore
function Clean-jason{
param (
        [string]$json_file
    )
$jsonContent = Get-Content $json_file -Raw
$assessment = ConvertFrom-Json -InputObject $jsonContent
# Loop through questionIds
$questionIds = $assessment.'spayee:resource'.'spayee:questionIds'

# Loop properly through keys of PSCustomObject
foreach ($key in $questionIds.PSObject.Properties.Name) {
    $qBlock = $questionIds.$key
    $qData = $qBlock.data

    # Clean the question text
    <#
    if ($qData.text) {
        if(!($qData.text -match '<img\s+src')){
        $qData.text = $qData.text -replace '<.*?>', ''
        }
    }
    #>

    # Clean each option content
    if ($qData.options -and $qData.options.option) {
        foreach ($option in $qData.options.option) {
            if ($option.content) {
                $option.content = $option.content -replace '<.*?>', '' -replace '\\n', '' -replace '\\t', ''
            }
        }
    }

    # Leave solution.text untouched
}

# Convert to clean JSON (escaped, but tags are *removed*, not preserved)
$cleanedJson = $assessment | ConvertTo-Json -Depth 10

# (Optional) save to file
# Set-Content -Path "cleaned.json" -Value $cleanedJson -Encoding utf8
$cleanedJson | Set-Content $json_file
}
function Create-Zip {
    param (
        [string]$zipPath
    )

    # Check file exists
    if (-not (Test-Path $zipPath)) {
        Write-Error "File not found: $zipPath"
        return
    }

    # Create temp folder
    $tempFolder = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tempFolder | Out-Null

    # Extract zip
    Expand-Archive -Path $zipPath -DestinationPath $tempFolder -Force

    # Locate assessmentData.json
    $jsonPath = Get-ChildItem -Path $tempFolder -Recurse -Filter 'assessmentData.json' | Select-Object -First 1

    if (-not $jsonPath) {
        Write-Error "assessmentData.json not found inside the zip."
        Remove-Item $tempFolder -Recurse -Force
        return
    }

    # Clean the JSON file
    Clean-jason $jsonPath.FullName

    # Re-zip the folder with the same name
    $zipfileName = [System.IO.Path]::GetFileName($zipPath)
    Compress-Archive -Path "$tempFolder\*" -DestinationPath "$loc\To_upload\$zipfileName"
    
    # Clean up
    Remove-Item $tempFolder -Recurse -Force
    return "$loc\To_upload\$zipfileName"
}

function call-uploadAPI{
    param (
        [string]$filePath
    )

# Step 2: Set API endpoint and file path
$uploadURL = ""

# Create HttpClient
$client = [System.Net.Http.HttpClient]::new()
$client.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $accessToken)

# Read the file
$fileBytes = [System.IO.File]::ReadAllBytes($filePath)
$fileContent = [System.Net.Http.ByteArrayContent]::new($fileBytes)
$fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/zip")

# Create multipart content
$multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
$multipartContent.Add($fileContent, "file", [System.IO.Path]::GetFileName($filePath))

# Send POST request
$response = $client.PostAsync($uploadUrl, $multipartContent).Result
$responseContent = $response.Content.ReadAsStringAsync().Result

# Output the result
#Write-host $response.StatusCode -BackgroundColor Red
Return $response.StatusCode,$responseContent
}


#Get-access Toke
$tokenURL = ""
$body = @{
     "email"="";
     "password_hash"="";
}
$response = Invoke-RestMethod -Uri $tokenURL -Method Post -Body $body 
$response.accessToken
$accessToken = $response.accessToken

"Free Aptitude Test
Free Geology Quiz
Free GS Quiz
Geology Test Series - UPSC Geo
GS-Test Series for UPSC-GSI
others
Question update
"
$completed_uploads=@()
$repeated_quiz=@()
$completed_uploads=@(gc "$($loc)\uploaded quizes.txt")
$course_name=(Get-ChildItem "D:\FIVER PROJECT\GC- extract quiz\Export_complete").name
#$course_name='GATE-NET-UPSC CGSE Live Online Classes 2025-26',"Free Aptitude Test","Free Geology Quiz","Free GS Quiz","Geology Test Series - UPSC Geo","GS-Test Series for UPSC-GSI","Question update","GATE-NET Test Series 2025-26","IIT-JAM 2025-26 Test Series","new"
#$course_name=@("Live Online Classes 2025-26")
pause
$count=1
foreach($course in $course_name){
$selected_folder= Get-ChildItem "D:\FIVER PROJECT\GC- extract quiz\Export_complete\$course"
Write-Host "Working on $course"
    foreach($filename in $selected_folder)
    {
        #$To_upload_file=create-zip $filename.fullname
        #$response=call-uploadAPI $To_upload_file

        $file_location = $filename.FullName
        $quiz_name=$filename.name
        if ($completed_uploads -match $filename.Name) {
            Write-Host "`nAlready uploaded duplicate Quiz skipping $($quiz_name)" -BackgroundColor Red
            $repeated_quiz += $filename.Name
            $Table+=[pscustomobject]@{'S.no'=0;'Course Name'=$course;'Quiz_Name'=$quiz_name;'Response Code'="SKIPPED";'Success'="FALSE";'Message'="Skipped for Course : $course";'Response String'=$null}
            continue     
        }
    # Proceed with upload
    $response = call-uploadAPI $file_location

    #Remove-item $To_upload_file
    $json_response=$response[1] | ConvertFrom-Json 
    $Table+=[pscustomobject]@{'S.no'=$count;'Course Name'=$course;'Quiz_Name'=$quiz_name;'Response Code'=$response[0];'Success'=$json_response.success;'Message'=$json_response.message;'Response String'=$response[1]
    }
    if($response[0] -eq 'OK'){ 
        Write-Host "$($count) Completed - $($filename.name)" -ForegroundColor Green
        $completed_uploads+=$filename.name      
        }
    $count++
    start-sleep -seconds 3
    write-host $response[0] -BackgroundColor Red
}
$completed_uploads |Set-Content "$($loc)\uploaded quizes.txt"
}
 $Table | Export-Excel -Path "$($loc)\Quiz_upload_result_$($dat).xlsx"  -WorksheetName "Upload_response" -TableStyle Medium17 -AutoSize 
 $repeated_quiz | Set-Content "$($loc)\Repeated_quiz.txt"

 $completed_uploads |Set-Content "$($loc)\uploaded quizes.txt"


 pause

 #$response = call-uploadAPI "D:\FIVER PROJECT\GC- extract quiz\Export\new\2025_Weekly_test__1__Strucutral_and_Mineralogy_.zip"