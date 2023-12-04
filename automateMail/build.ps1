$exclude = @("venv", "automateGetInfIfood.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "automateGetInfIfood.zip" -Force