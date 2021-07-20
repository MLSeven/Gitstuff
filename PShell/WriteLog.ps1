function Write-Log {
  param (
    $path="$($env:SystemDrive)\psTranscript.log",
    $text="",
    [object]$object,
    [switch]$newlog
  )
  
  if ($newlog) {
    set-content -path $path -value "$([datetime]::now.ToUniversalTime().ToString("s")) : Opening new Log File" -Encoding UTF8
  }
  # for each line of text pre-pend the UTC Timesamp
  $text | foreach-object {"$([datetime]::now.ToUniversalTime().ToString("s")) : $($_)"} | add-content -path $path -Encoding UTF8
  $object | out-string -width 132 | add-content -path $path -Encoding UTF8 
}
