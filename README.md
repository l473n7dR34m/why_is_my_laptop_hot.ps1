why_is_my_laptop_hot.ps1

A self-contained PowerShell diagnostic script that monitors CPU clock speed, utilisation, memory usage, and throttling behaviour over time.
It automatically produces a clear summary at the end of each run to help identify thermal throttling, power limits, or background load.

Overview

The script logs CPU and memory statistics at fixed intervals and uses heuristic analysis to determine if performance issues are caused by heat, power policies, or background software.

It runs entirely in PowerShell, requires no installation or administrative rights, and leaves no traces on the system.

Usage

Run from PowerShell:

Set-ExecutionPolicy -Scope Process Bypass
.\why_is_my_laptop_hot.ps1


Optional parameters:

.\why_is_my_laptop_hot.ps1 -IntervalSeconds 5 -DurationMinutes 15


IntervalSeconds: how often samples are taken (default 5 seconds)

DurationMinutes: total runtime (default 10 minutes, 0 = continuous)

Press Ctrl+C to stop early.

Output

A CSV log is saved to your Desktop:

ThermalLog_YYYY-MM-DD_HH-MM-SS.csv


Each entry contains:

Column	Description
Timestamp	Time of sample (ISO 8601)
CpuLoadPct	Total CPU utilisation
ClockMHz / MaxMHz	Current and maximum CPU speed
RamUsedMB / RamAvailMB	Memory in use and available
TopCpuProc / TopCpuPct	Process using the most CPU
TopMemProc / TopMemMB	Process using the most RAM
ThermalEventsSinceStart	Number of Windows thermal events
Alert	Indicates sustained low clocks below 80 percent of max
LowClockStreak	Consecutive low-clock samples
Smart Diagnosis

When logging completes, the script analyses the run and prints a diagnosis with suggested next steps.

Possible conditions
Condition	Explanation
Healthy	No throttling or abnormal behaviour
Boost-limited	CPU locked at base frequency under load
Thermal Throttling	CPU slows down under load due to heat
Power Policy Suspect	Power plan restricting CPU boost
Firmware Suspect	BIOS or vendor utility disabling boost
Event Log Heat	Windows recorded thermal events
EDR Hint	Security or endpoint agent consuming CPU
Example output
--- Session summary ---
Samples: 240
Clock MHz avg/min/max: 3450/2200/3900
CPU load % avg/max: 62/100

--- Smart diagnosis ---
Conclusion: Low clocks while loaded. Likely thermal or power throttling.
Recommended next steps:
1. Clear vents and rerun. If unchanged, check power plan and BIOS settings.
Full log saved to C:\Users\<user>\Desktop\ThermalLog_2025-10-16_18-30-22.csv

Interpreting Results

Stable clocks near or above base = healthy.

Flat clocks at base under high load = boost disabled or limited by power plan.

Frequent dips below 80 percent of max = thermal or power throttling.

Thermal events in logs = BIOS or firmware triggered throttling.

EDR hint = background or security process consuming sustained CPU.

Optional: Built-in Load Generator

To simulate load without external tools, you can add this helper:

function Start-CpuLoad {
  param([int]$Minutes=5,[int]$TargetPercent=95)
  $cores=[Environment]::ProcessorCount
  $until=(Get-Date).AddMinutes($Minutes)
  1..$cores | ForEach-Object {
    Start-Job -ScriptBlock {
      param($u)
      while((Get-Date) -lt $u){1..20000|%{[void][math]::Sqrt($_)}}} -ArgumentList $until
  } | Out-Null
  Write-Host "CPU load running for $Minutes minutes on $cores cores."
}


Example:

Start-CpuLoad -Minutes 5 -TargetPercent 90


Run this in a separate PowerShell window while logging to simulate a heavy workload.

Typical Use Cases

Test airflow differences (stand vs desk)

Validate cooling pad or fan performance

Detect boost limits caused by power policy

Identify background CPU-heavy software

Provide evidence for IT or warranty claims

Notes

No system changes are made.

Logs are plain text and safe to delete.

Works on any modern Windows system with WMI/CIM enabled.

Best results on CPUs that report accurate clock metrics.

Author

Created for system administrators and technical users who want fast, transparent insight into CPU throttling and thermal behaviour using only PowerShell.
