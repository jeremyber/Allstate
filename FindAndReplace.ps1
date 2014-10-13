$postScriptFiles=get-childitem . *.ps -rec
foreach ($file in $postScriptFiles)
{
(Get-Content $file.PSPath) | 
Foreach-Object {$_ -replace "Duplex DuplexNoTumble", "Duplex None"} | 
Foreach-Object {$_ -replace "<</Duplex true /Tumble false>>", "<</Duplex false>>"} |
Set-Content $file.PSPath
}
