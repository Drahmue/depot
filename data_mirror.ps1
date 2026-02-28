# Define source and destination paths
$sourceInput = "\\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Input"
$sourceOutput = "\\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Output"
$destInput = "C:\Users\ah\Dev\depot\Data Mirror\Finance_Input"
$destOutput = "C:\Users\ah\Dev\depot\Data Mirror\Finance_Output"

# Copy files and folders
Write-Host "Copying Finance_Input..."
Copy-Item -Path $sourceInput\* -Destination $destInput -Recurse -Force

Write-Host "Copying Finance_Output..."
Copy-Item -Path $sourceOutput\* -Destination $destOutput -Recurse -Force

Write-Host "Data mirror copy complete."

