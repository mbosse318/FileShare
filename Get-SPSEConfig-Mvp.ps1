function EnsureFolder {
    param(
        [parameter(Mandatory = $true)] [string]$Path
    )

    if (!(Test-Path $Path)) {
        New-Item $Path -ItemType Directory | Out-Null
    }
}

function EnsureTrainingBackslash($string) {
    if ($string -notmatch '\\$')
    {
        $string += '\'
    }

    return $string
}

function Get-OutputPath {
    $outputPath = "C:\users\l-2hieanfnhta3m\Output"
#   $outputPath = "\\SP-S-V-300\SPOTenant$"
    return $outputPath
}

function Get-LogPath {
    $currentDateTime = "{0:yyyy-MM-dd_HH-mm-ss}" -f (Get-Date)
    $basePath = (EnsureTrainingBackslash (Get-OutputPath))
    return "$($basePath)SPSEConfig_$($currentDateTime).log"
}

function Get-DateTimeOutputPath {
    $currentDateTime = "{0:yyyy-MM-dd_HH-mm-ss}" -f (Get-Date)
    $basePath = (EnsureTrainingBackslash (Get-OutputPath))
    return "$($basePath)SPSEConfigFiles_$($currentDateTime)\"
}

function Write-LogEntry {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)] [string] $logEntry
    )

    $logLine = "$("{0:yyyy-MM-dd_HH-mm-ss}`t" -f (Get-Date))$($logEntry)"
    Write-Host $logLine
}

function Run-Cmdlet {
    param(
        [parameter(Mandatory = $true)] [string] $CmdletName,
        [parameter(Mandatory = $true)] [string] $OutputPath,
        [parameter(Mandatory = $true)] [string] $LogPath,
        [parameter(Mandatory = $false)] [string] $Depth="2"
    )

    Write-LogEntry -logEntry "Calling $($CmdletName)"
    try {
        (Invoke-Expression -Command $cmdletName) | Export-Clixml "$($outputPath)$($cmdletName).xml" -Depth $depth
    }
    catch {
        Write-LogEntry $_.ToString()
    }
}

$logPath = Get-LogPath
Start-Transcript -Path $logPath

$outputPath = Get-DateTimeOutputPath
EnsureFolder -Path $outputPath

# start generating the config files and saving to the output folder for this run
Run-Cmdlet -CmdletName Get-SPWebApplication -OutputPath $outputPath -LogPath $logPath -Depth 1

# Retrieve farm information
Run-Cmdlet -CmdletName Get-SPFarm -OutputPath $outputPath -LogPath $logPath -Depth 2
Run-Cmdlet -CmdletName Get-SPFarmConfig -OutputPath $outputPath -LogPath $logPath -Depth 2

# Retrieve information about the servers in the farm
Run-Cmdlet -CmdletName Get-SPServer -OutputPath $outputPath -LogPath $logPath -Depth 3

# Retrieve information about the logging levels that have been set
Run-Cmdlet -CmdletName Get-SPLogLevel -OutputPath $outputPath -LogPath $logPath -Depth 2

Stop-Transcript