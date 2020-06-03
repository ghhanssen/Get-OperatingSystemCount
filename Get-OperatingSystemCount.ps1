<#
    Source: https://evotec.xyz/getting-windows-10-build-version-from-active-directory

    Date: 03.06.2020
    Name: Geir-Hugo Hanssen

    Output: 
        -----------------------------------------------------------------------------------
        Year: 2020, Week: 23
        2020.06.03 08:24

        Computers: 11

        Name                   Count
        ----                   -----
        Windows 7 Professional     1
        Windows 10 1809            8
        Windows 10 1909            2
        -----------------------------------------------------------------------------------
#>

Import-Module ActiveDirectory

function ConvertTo-OperatingSystem {
    [CmdletBinding()]
    param(
        [string] $OperatingSystem,
        [string] $OperatingSystemVersion
    )

    if ($OperatingSystem -like 'Windows 10*') {
        $Systems = @{
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



$Computers = Get-ADComputer -Filter 'operatingsystem -notlike "*server*"' -properties Name, OperatingSystem, OperatingSystemVersion, LastLogonDate, whenCreated
$ComputerList = foreach ($_ in $Computers) {
    [PSCustomObject] @{
        Name                   = $_.Name
        OperatingSystem        = $_.OperatingSystem
        OperatingSystemVersion = $_.OperatingSystemVersion
        System                 = ConvertTo-OperatingSystem -OperatingSystem $_.OperatingSystem -OperatingSystemVersion $_.OperatingSystemVersion
        LastLogonDate          = $_.LastLogonDate
        WhenCreated            = $_.WhenCreated
    }
}


$Count = ($Computers | Measure-Object).Count
$Time = (Get-date -Format "yyyy.MM.dd HH:mm")
$Year = (Get-Date -UFormat %Y)
$Week = GetWeekOfYear (get-date)
#$Week = 22

Write-Host "-----------------------------------------------------------------------------------"
Write-Host "Year: $Year, Week: $Week" -ForegroundColor Green
Write-Host $Time
Write-Host ""
Write-Host "Computers: $Count"

#Update database
#<Removed>


#OperatingSystem overview
$ComputerList | Group-Object -Property System | Sort-Object -Property Name | Format-Table -Property Name, Count
Write-Host "-----------------------------------------------------------------------------------"

#All Computers
#$ComputerList | Format-Table -AutoSize
