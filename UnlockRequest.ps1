###########################
## Set request file location (actual driver letter, not map drive letter)
#########################
$UnlockRequestInput = 'F:\ShareLocation\S\UnlockRequest\UnlockRequest.txt'

# Read unlock request file
try {
  $FileName = Get-Content $UnlockRequestInput -ErrorAction stop
} catch {
  # Custom action to perform before terminating
  throw 'Unable to read input file'
}

###########################
## Set paths
#########################

# Change path from share path to local path
$FileName = $FileName -replace 'S:\\','F:\ShareLocation\S'

###########################
## Set path
#########################

# Verify change
if ($FileName -match '^F:\\ShareLocations\\S\\' -eq $false) {
  throw 'Invalid Path'
}


# Check if file exists
if(![System.IO.File]::Exists($FileName)){
  throw "File Doesn't Exist"
}

# Unlock file
.\psfile64.exe $FileName -c -nobanner

# If office temp file exists, unlock it also
$FileName  -Match '(.+\\)([^\\].+)'
$TmpFileName = $Matches[1]  + '~$' + $Matches[2] 

if([System.IO.File]::Exists($TmpFileName)){
  .\psfile64.exe $TmpFileName -c   -nobanner -accepteula
  # Unlock file again
  Sleep 5
  .\psfile64.exe $FileName -c -nobanner -accepteula
}

Set-Content -path $UnlockRequestInput -Value ''
