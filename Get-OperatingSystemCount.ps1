<#
    Source: https://evotec.xyz/getting-windows-10-build-version-from-active-directory


    Date:    02.11.2020
    Name:    Geir-Hugo Hanssen

    Comment: Script to get the count of all Windows klient build versions, and save to file. Then manually update excel document to see trends.


    Run:
        .\Get-OperatingSystemCount.ps1            Will sett Week to current Week 
        .\Get-OperatingSystemCount.ps1 -Week 10   Override Week number 



    Screen Output: 
        -----------------------------------------------------------------------------------
        Year: 2020, Week: 24
        2020.06.07 18:12

        Computers: 11

        Name                   Count
        ----                   -----
        Windows 10 1809            8
        Windows 10 1909            2
        Windows 7 Professional     1


        Saved to: I:\Powershell\AD\Get-OperatingSystem Count\OperatingSystemCount_2020_24.txt
        -----------------------------------------------------------------------------------
    


    File Output: 
        -----------------------------------------------------------------------------------
        Year: 2020, Week: 24
        2020.06.07 18:12

        Computers: 11

        Count - OperatingSystem
        8 - Windows 10 1809
        2 - Windows 10 1909
        1 - Windows 7 Professional
        -----------------------------------------------------------------------------------
#>

[CmdletBinding()]
    param (
        [Parameter(Mandatory = $false) ]
        [string]$Week
    )



Import-Module ActiveDirectory


#---------------------------------- Functions Start -------------------------------------------
function ConvertTo-OperatingSystem {
    [CmdletBinding()]
    param(
        [string] $OperatingSystem,
        [string] $OperatingSystemVersion
    )

    if ($OperatingSystem -like 'Windows 10*') {
        $Systems = @{
			'10.0 (19042)' = "Windows 10 20H2"
            '10.0 (19041)' = "Windows 10 2004"
            '10.0 (18363)' = "Windows 10 1909"
            '10.0 (18362)' = "Windows 10 1903"
            '10.0 (17763)' = "Windows 10 1809"
            '10.0 (17134)' = "Windows 10 1803"
            '10.0 (16299)' = "Windows 10 1709"
            '10.0 (15063)' = "Windows 10 1703"
            '10.0 (14393)' = "Windows 10 1607"
            '10.0 (10586)' = "Windows 10 1511"
            '10.0 (10240)' = "Windows 10 1507"
        }
        $System = $Systems[$OperatingSystemVersion]
    } elseif ($OperatingSystem -notlike 'Windows 10*') {
        $System = $OperatingSystem
    }
    if ($System) {
        $System
    } else {
        'Unknown'
    }
}

Function GetWeekOfYear($date) {
    # Note: first day of week is Sunday
    $intDayOfWeek = (get-date -date $date).DayOfWeek.value__
    $daysToWednesday = (3 - $intDayOfWeek)
    $wednesdayCurrentWeek = ((get-date -date $date)).AddDays($daysToWednesday)

    # %V basically gets the amount of '7 days' that have passed this year (starting at 1)
    $weekNumber = get-date -date $wednesdayCurrentWeek -uFormat %V

    return $weekNumber
}
#---------------------------------- Functions End -------------------------------------------


$DateAndTime = (Get-date -Format "yyyy.MM.dd HH:mm")
$Year = (Get-Date -UFormat %Y)
if (-Not $Week) { $Week = GetWeekOfYear (get-date) }
$Date = (Get-Date -Format "yyyy.MM.dd")
$Time = (Get-Date -Format "HHmm")
$File = $PSScriptRoot + "\OperatingSystemCount_" + $Year + "_" + $Week + ".txt"

$Computers = Get-ADComputer -Filter { operatingsystem -notlike "*server*" -and operatingsystem -notlike "HDS NAS OS" }  -properties Name, OperatingSystem, OperatingSystemVersion
$ComputerList = foreach ($_ in $Computers) {
    [PSCustomObject] @{
        Name                   = $_.Name
        OperatingSystem        = $_.OperatingSystem
        OperatingSystemVersion = $_.OperatingSystemVersion
        System                 = ConvertTo-OperatingSystem -OperatingSystem $_.OperatingSystem -OperatingSystemVersion $_.OperatingSystemVersion
    }
}

$Count = ($Computers | Measure-Object).Count

Write-Host "-----------------------------------------------------------------------------------"
Write-Host "Year: $Year, Week: $Week" -ForegroundColor Green
Write-Host $DateAndTime
Write-Host ""
Write-Host "Computers: $Count"
"-----------------------------------------------------------------------------------" | Out-File -Append -FilePath $File
"Year: $Year, Week: $Week" | Out-File -Append -FilePath $File
$DateAndTime | Out-File -Append -FilePath $File
"" | Out-File -Append -FilePath $File
"Computers: $Count" | Out-File -Append -FilePath $File
"" | Out-File -Append -FilePath $File
"Count - OperatingSystem" | Out-File -Append -FilePath $File

$Export = $ComputerList | Group-Object -Property System | Sort-Object -Property Name
$Export | Format-Table -Property Name, Count

foreach ($_ in $Export) {
    $Name         = $_.Name
    $Count        = $_.Count
    "$Count - $Name" | Out-File -Append -FilePath $File
}

Write-Host "Saved to: $File" -ForegroundColor Green