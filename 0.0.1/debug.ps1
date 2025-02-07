Remove-Module SPS-WinTools -ErrorAction Ignore
Import-Module "$($PSScriptRoot)\SPS-WinTools.psd1" -Force
# Set the most constrained mode
Set-StrictMode -Version Latest
# Set the error preference
$ErrorActionPreference = 'Stop'
# Set the verbose preference in order to get some insights
$VerbosePreference = 'Continue'

# change the verbose color so it's not the same color than the warnings
if (Get-Variable -Name PSStyle -ErrorAction SilentlyContinue) {
    $PSStyle.Formatting.Verbose = $PSStyle.Foreground.Cyan
}

# Test Get-Enum
# Create an enum
Enum MyEnum {
    Value1 = 1
    Value2 = 2
    Value3 = 3
}
# Write-Host "Enum MyEnum created" -ForegroundColor Magenta
[MyEnum] | Get-Enum
# Write-Host "Enum MyEnum created using full" -ForegroundColor Magenta
[MyEnum] | Get-Enum -full



BREAK
$StopwatchNative = [System.Diagnostics.Stopwatch]::new()
$StopwatchNative.Start()
$Results = Get-ProcessRelatives -id $PID
$Results | Format-Table -Autosize
$StopwatchNative.Stop()
Write-Host "The function Get-ProcessRelatives took: $($StopwatchNative.Elapsed.TotalMilliseconds)ms"

$StopwatchNative.Restart()
Get-Caller
$StopwatchNative.Stop()

Write-Host "The function Get-Caller took: $($StopwatchNative.Elapsed.TotalMilliseconds)ms"