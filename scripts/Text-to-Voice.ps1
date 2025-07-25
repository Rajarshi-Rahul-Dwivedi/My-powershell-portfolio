#script created by Rajarshi
Add-Type -AssemblyName System.Speech
$loc=$PSScriptRoot
# Create a SpeechSynthesizer object
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

# Customize voice and speed
$speak.SelectVoice("Microsoft Zira Desktop") 
$speak.Rate = 0 # -10 to 10, where -10 is slowest
#$speak.SpeakAsync()
$text_content=Get-Content "$loc\text_input.txt"
write-host "`n`nReading Text : `n`n" -ForegroundColor Red
write-host "$text_content`n`n" -ForegroundColor Cyan
$speak.Speak($text_content) | out-null
pause
