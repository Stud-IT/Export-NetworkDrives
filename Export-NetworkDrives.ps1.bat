# & cls & powershell -Command "Invoke-Command -ScriptBlock ([ScriptBlock]::Create(((Get-Content """%0""") -join """`n""")))" & exit
# The above line makes the script executable when renamed .cmd or .bat

# Show and export all network drives
# Author: Anton Pusch (Stud-IT) 
# Last update: 2021-03-11

# All data is saved to a user folder in the root directory
$folder = $PWD.Drive.Root + $env:username

# Create folder if necessary and open it
if (Test-Path -Path $folder) {
    Write-Host "Using directory $folder`n"
} else {
    Write-Host "Create directory $folder`n"
    New-Item -ItemType Directory -Path $folder | Out-Null
}
Set-location -Path $folder

# Get list of all network drives
Write-Host "Scanning through network drives... " -NoNewLine
$net = New-Object -ComObject Shell.Application
$drives = @() # Formats drives as an array
$drives += Get-PSDrive `
   | Where-Object {$_.Provider.Name -eq "FileSystem"} `
   | Select-Object @{Label="Drive"; Expression={$_.Name}}, `
   @{Label="DisplayName"; Expression={$net.NameSpace($_.Root).Self.Name}}, `
   @{Label="Network Location"; Expression={$_.DisplayRoot}}, `
   @{Label="Used|GB"; Expression={if($_.Used -ne "") {[math]::Round($_.Used / 1GB, 2)}}}, `
   @{Label="Free|GB"; Expression={if($_.Free -ne "") {[math]::Round($_.Free / 1GB, 2)}}}, `
   @{Label="% Full"; Expression={if($_.Used -ne "" -and $_.Free -ne "") {[math]::Round(100 * $_.Free / ($_.Used + $_.Free))}}}, `
   Description, `
   Root

# Count output
Write-Host ("Found " + $drives.Count + " drives")

# Display all network drives
Write-Output ($drives | Format-Table)

# Export table to Excel
do {
   $error.clear()
   try {
      # Try to override file
      $drives | Export-Csv -Path "drives-$env:username.csv" -Delimiter ";" -NoTypeInformation
   } catch {
      Read-Host -Prompt "Couldn't override file $folder\drives-$env:username.csv`nTry again?`t^C to cancel`tEnter to repeat"
   }
} while ($error)
Write-Host "Exported to $folder\drives-$env:username.csv`n"

# Keep console window open
$reply = Read-Host -Prompt "Finished!`to to open spreadsheet file`tEnter to quit"
if ($reply.ToLower() -eq 'o') {
   # Open the spreadsheet
   Start-Process -FilePath ".\drives-$env:username.csv"
}
