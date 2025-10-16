param(
    [int]$IntervalSeconds = 5,
    [int]$DurationMinutes = 10    # 0 = run forever
)

# Output file
$desktop = [Environment]::GetFolderPath('Desktop')
$stamp   = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$csv     = Join-Path $desktop "ThermalLog_$stamp.csv"

# CSV header
"Timestamp,CpuLoadPct,ClockMHz,MaxMHz,RamUsedMB,RamAvailMB,TopCpuProc,TopCpuPct,TopMemProc,TopMemMB,ThermalEventsSinceStart,Alert,LowClockStreak" |
  Out-File $csv -Encoding utf8

Write-Host "Logging to $csv every $IntervalSeconds s. Duration: $DurationMinutes min (0 = infinite). Ctrl+C to stop early.`n"

# State
$startTime = Get-Date
$endTime   = if ($DurationMinutes -gt 0) { $startTime.AddMinutes($DurationMinutes) } else { [datetime]::MaxValue }
$logicalCores = [Environment]::ProcessorCount
[int]$thermalCount = 0
$stats = @()
$lastProcTimes = @{}  # pid -> { cpu cumulative seconds, name }
$procCpuHits = @{}    # name -> hit count as top CPU
$procMemHits = @{}    # name -> hit count as top memory

# Alert thresholds
$LowClockPctThreshold = 0.8
$LowClockMinStreak    = 3
[int]$lowClockStreak  = 0

# Diagnosis accumulators
[int]$samples = 0
[int]$lowClockSamples = 0                   # instantaneous low clock
[int]$highLoadSamples = 0                   # CpuLoadPct >= 70
[int]$highLoadLowClockSamples = 0          # high load while low clock
[int]$eventsTotal = 0

function Update-ProcTimesInit {
    foreach ($p in (Get-Process)) {
        try { $lastProcTimes[$p.Id] = @{ cpu = $p.CPU; name = $p.ProcessName } } catch { }
    }
}

function Get-TopCpuDelta {
    $exclude = @('Idle','System','WmiPrvSE','Memory Compression','svchost','MsMpEng')
    $procs = Get-Process | Where-Object { $_.Id -ne $PID -and $_.ProcessName -notin $exclude }
    $deltas = @()
    foreach ($p in $procs) {
        try {
            $prev = $lastProcTimes[$p.Id]
            $currSec = $p.CPU
            if ($prev) {
                $deltaSec = [math]::Max(0, $currSec - $prev.cpu)
                $cpuPct = [math]::Round(($deltaSec / $IntervalSeconds) * (100 / $logicalCores), 1)
                if ($cpuPct -gt 0) {
                    $deltas += [pscustomobject]@{
                        Id=$p.Id; Name=$p.ProcessName; CpuPct=$cpuPct; MemMB=[math]::Round($p.WorkingSet64/1MB,0)
                    }
                }
                $lastProcTimes[$p.Id] = @{ cpu = $currSec; name = $p.ProcessName }
            } else {
                $lastProcTimes[$p.Id] = @{ cpu = $currSec; name = $p.ProcessName }
            }
        } catch { }
    }
    if ($deltas.Count -gt 0) { $deltas | Sort-Object CpuPct -Descending | Select-Object -First 1 } else { $null }
}

function Get-ThermalEventsSince($since) {
    try {
        (Get-WinEvent -ProviderName "Microsoft-Windows-Kernel-Power" -ErrorAction Stop |
            Where-Object { $_.Id -in 86,87 -and $_.TimeCreated -ge $since }).Count
    } catch { 0 }
}

Update-ProcTimesInit

try {
    while ($true) {
        if ((Get-Date) -ge $endTime) { break }
        $nowIso = Get-Date -Format o

        # CPU via WMI
        $wmiCpu = Get-CimInstance Win32_Processor | Select-Object -First 1 LoadPercentage, CurrentClockSpeed, MaxClockSpeed
        $cpuLoad = [int]$wmiCpu.LoadPercentage
        $clk = [int]$wmiCpu.CurrentClockSpeed
        $max = [int]$wmiCpu.MaxClockSpeed

        # Low-clock detection
        $alert = ''
        $isLowClock = $false
        if ($max -gt 0 -and $clk -lt [int]([double]$max * $LowClockPctThreshold)) {
            $lowClockStreak++
            $isLowClock = $true
        } else {
            $lowClockStreak = 0
        }
        if ($lowClockStreak -ge $LowClockMinStreak) {
            $alert = 'LOW_CLOCK'
            Write-Warning "Clock below $([int]($LowClockPctThreshold*100)) percent of max for $lowClockStreak samples"
            [console]::Beep(2000,200) | Out-Null
        }

        # Memory
        $os = Get-CimInstance Win32_OperatingSystem
        $usedMB = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory)/1024,0)
        $freeMB = [math]::Round($os.FreePhysicalMemory/1024,0)

        # Top CPU delta and top memory proc
        $topCpu = Get-TopCpuDelta
        $topMem = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 1

        # Thermal events since start
        $thermalCount = Get-ThermalEventsSince $startTime
        $eventsTotal = $thermalCount

        # Record hits for diagnosis
        $samples++
        if ($isLowClock) { $lowClockSamples++ }
        if ($cpuLoad -ge 70) { $highLoadSamples++ }
        if ($isLowClock -and $cpuLoad -ge 70) { $highLoadLowClockSamples++ }
        if ($topCpu) {
            $name = $topCpu.Name
            if ($procCpuHits.ContainsKey($name)) { $procCpuHits[$name]++ } else { $procCpuHits[$name] = 1 }
        }
        if ($topMem) {
            $mname = $topMem.ProcessName
            if ($procMemHits.ContainsKey($mname)) { $procMemHits[$mname]++ } else { $procMemHits[$mname] = 1 }
        }

        $line = '{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}' -f `
            $nowIso,$cpuLoad,$clk,$max,$usedMB,$freeMB,
            ($topCpu.Name   -as [string]),
            ($topCpu.CpuPct -as [string]),
            $topMem.ProcessName,
            ([math]::Round($topMem.WorkingSet64/1MB,0)),
            $thermalCount,
            $alert,
            $lowClockStreak

        Add-Content -Path $csv -Value $line
        Write-Host $line

        # stats for numeric summary
        $stats += [pscustomobject]@{
            Timestamp = $nowIso
            CpuLoadPct= $cpuLoad
            ClockMHz  = $clk
            MaxMHz    = $max
        }

        Start-Sleep -Seconds $IntervalSeconds
    }
}
finally {
    if ($stats.Count -gt 0) {
        Write-Host "`n--- Session summary ---"
        $avgClock = [int](($stats.ClockMHz | Measure-Object -Average).Average)
        $minClock = ($stats.ClockMHz | Measure-Object -Minimum).Minimum
        $maxClock = ($stats.ClockMHz | Measure-Object -Maximum).Maximum
        $avgLoad  = [int](($stats.CpuLoadPct | Measure-Object -Average).Average)
        $maxLoad  = ($stats.CpuLoadPct | Measure-Object -Maximum).Maximum

        "Samples: $($stats.Count)" | Write-Host
        "Clock MHz avg/min/max: $avgClock/$minClock/$maxClock" | Write-Host
        "CPU load % avg/max: $avgLoad/$maxLoad" | Write-Host
        "Lowest clocks (first 5):" | Write-Host
        $stats | Sort-Object ClockMHz | Select-Object -First 5 Timestamp,ClockMHz | Format-Table -AutoSize

        Write-Host "`n--- Smart diagnosis ---"
        $lowClockPct = if ($samples -gt 0) { [math]::Round(100 * $lowClockSamples / $samples, 1) } else { 0 }
        $highLoadLowClockPct = if ($highLoadSamples -gt 0) { [math]::Round(100 * $highLoadLowClockSamples / $highLoadSamples, 1) } else { 0 }

        $phrases = @{
            Healthy              = "No signs of throttling. Clocks stable at or above base under load."
            BoostDisabled        = "CPU appears boost-limited. Clocks stayed at base even under high load."
            ThermalThrottling    = "Low clocks while loaded. Likely thermal or power throttling."
            PowerPolicySuspect   = "Power plan or policy likely capping frequency."
            FirmwareSuspect      = "Firmware or OEM setting may be disabling turbo."
            EventLogHeat         = "Windows logged thermal events during the run."
            WorkloadHeavy        = "Sustained high CPU load observed."
            WorkloadLight        = "Workload mostly light. No evidence of heat-induced slowdowns."
            EDRHint              = "Security or agent process frequently topped CPU. Review exclusions or schedules."
        }

        $conclusions = New-Object System.Collections.Generic.List[string]
        $actions     = New-Object System.Collections.Generic.List[string]

        $baseIsFlat = ($minClock -eq $maxClock -and $avgClock -eq $maxClock)

        if ($eventsTotal -gt 0) {
            $conclusions.Add($phrases.EventLogHeat)
            $actions.Add("Inspect Kernel-Power thermal events in Event Viewer and check airflow.")
        }

        if ($highLoadSamples -ge [math]::Ceiling($samples * 0.2) -and $highLoadLowClockPct -ge 20) {
            $conclusions.Add($phrases.ThermalThrottling)
            $actions.Add("Clear vents and rerun. If unchanged, review power plan and BIOS boost settings.")
        } elseif ($baseIsFlat -and $maxLoad -ge 90) {
            $conclusions.Add($phrases.BoostDisabled)
            $conclusions.Add($phrases.PowerPolicySuspect)
            $actions.Add("Temporarily switch to High performance, set PERFBOOSTMODE=2, PERFEPP=0, PROCTHROTTLEMAX=100, then rerun.")
            $actions.Add("If still flat, enable Turbo or Performance mode in BIOS or vendor utility.")
        } elseif ($avgLoad -lt 40 -and $lowClockPct -le 5) {
            $conclusions.Add($phrases.Healthy)
            $conclusions.Add($phrases.WorkloadLight)
        } else {
            if ($lowClockPct -gt 10) {
                $conclusions.Add($phrases.ThermalThrottling)
                $actions.Add("Check airflow, fan curve and paste. Confirm power plan is not clamping frequency.")
            } else {
                $conclusions.Add($phrases.Healthy)
            }
        }

        if ($procCpuHits.Keys.Count -gt 0) {
            $topCpuOffender = ($procCpuHits.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
            if ($topCpuOffender -match 'defender|msmpeng|crowdstrike|sentinel|tanium|mcafee|sophos|carbon|agent|edr|security') {
                $conclusions.Add($phrases.EDRHint)
                $actions.Add("If allowed, schedule scans outside work hours or add exclusions for heavy tools.")
            }
        }

        "Conclusion: " + ($conclusions -join " ") | Write-Host

        if ($actions.Count -gt 0) {
            Write-Host "Recommended next steps:"
            $i = 1
            foreach ($a in $actions | Select-Object -Unique) {
                Write-Host ("{0}. {1}" -f $i, $a)
                $i++
            }
        }

        Write-Host "`nFull log saved to $csv`n"
    }
}
