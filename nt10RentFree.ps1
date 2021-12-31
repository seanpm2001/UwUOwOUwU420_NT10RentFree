<#
nt10RentFree. Tool for making Windows 10/11 live rent free in your PC.
Copyright (C) 2021 Gamers Against Weed

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
#>

#Requires -RunAsAdministrator

Param (
    [Alias('Key')]
    [Parameter(Position=0)]
    $ProvidedKey,

    [Parameter()]
    [Switch]$ForceKms38
)

$version = '0.1'
$kms38_forced = $ForceKms38.IsPresent
$slc_where = 'ApplicationId=''55c92734-d682-4d71-983e-d6ec3f16059f'' AND PartialProductKey != NULL'
$required_files = 'gatherosstate.exe', 'slc.dll'

function GetHwidProductKey($SkuId, $Build) {
    $SkuId = [String]$SkuId

    $keys = @{
        '4'   = 'XGVPP-NMH47-7TTHJ-W3FW7-8HV2C';
        '27'  = '3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT';
        '48'  = 'VK7JG-NPHTM-C97JM-9MPGT-3V66T';
        '49'  = '2B87N-8KFHP-DKV6R-Y2C8J-PKCKT';
        '98'  = '4CPRK-NM3K3-X6XXQ-RXX86-WXCHW';
        '99'  = 'N2434-X9D7W-8PF6X-8DV9T-8TYMD';
        '100' = 'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT';
        '101' = 'YTMG3-N6DKC-DKB77-7M9GH-8HVX7';
        '121' = 'YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY';
        '122' = '84NGF-MHBT6-FXBX8-QWJK7-DRR8H';
        '161' = 'DXG7C-N36C4-C4HTG-X4T3X-2YV77';
        '162' = 'WYPNQ-8C467-V2W6J-TX4WX-WT2RQ';
        '164' = '8PTT6-RNW4C-6V7J2-C2D3X-MHBPB';
        '165' = 'GJTYN-HDMQY-FRR76-HVGC7-QPF8P';
        '175' = 'NJCF7-PW8QT-3324D-688JX-2YV66';
        '188' = 'XQQYW-NFFMW-XJPBH-K8732-CKFFD';
        '191' = 'QPM6N-7J2WJ-P88HH-P3YRH-YY74H';
        '203' = 'KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W';
    }

    if($null -ne $keys[$SkuId]) {
        Return $keys[$SkuId]
    }

    switch ($Build) {
        {$_ -eq 10240} {
            $keys = @{
                '125' = 'FWN7H-PF93Q-4GGP8-M8RF3-MDWWW';
                '126' = '8V8WN-3GXBH-2TCMG-XHRX3-9766K';
            }
        }
        {$_ -eq 14393} {
            $keys = @{
                '125' = 'NK96Y-D9CD8-W44CQ-R8YTK-DYJWX';
                '126' = '2DBW3-N2PJG-MVHW3-G7TDK-9HKR4';
            }
        }
        {$_ -eq 17763} {
            $keys = @{
                '125' = '43TBQ-NH92J-XKTM7-KT3KK-P39PB';
                '126' = 'M33WV-NHY3C-R7FPM-BQGPT-239PG';
            }
        }
    }

    if($null -ne $keys[$SkuId]) {
        Return $keys[$SkuId]
    }

    Return $null
}

function GetKms38ProductKey($SkuId, $Build) {
    $SkuId = [String]$SkuId

    $keys = @{
        '4'   = 'NPPR9-FWDCX-D2C8J-H872K-2YT43';
        '27'  = 'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4';
        '48'  = 'W269N-WFGWX-YVC9B-4J6C9-T83GX';
        '49'  = 'MH37W-N47XK-V7XM9-C7227-GCQG9';
        '98'  = '3KHY7-WNT83-DGQKR-F7HPR-844BM';
        '99'  = 'PVMJN-6DFY6-9CCP6-7BKTT-D3WVR';
        '100' = '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH';
        '101' = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99';
        '121' = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2';
        '122' = '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ';
        '161' = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J';
        '162' = '9FNHH-K3HBT-3W4TD-6383H-6XYWF';
        '164' = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y';
        '165' = 'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC';
        '171' = 'YYVX9-NTFWV-6MDM3-9PT4T-4M68B';
        '172' = '44RPN-FTY23-9VTTB-MP9BX-T84FV';
        '175' = '7NBT4-WGBQX-MP4H7-QXFF8-YP3KX';
    }

    if($null -ne $keys[$SkuId]) {
        Return $keys[$SkuId]
    }

    switch ($Build) {
        {$_ -eq 10240} {
            $keys = @{
                '125' = 'WNMTR-4C88C-JK8YV-HQ7T2-76DF9';
                '126' = '2F77B-TNFGY-69QQF-B8YKP-D69TJ';
            }
        }
        {$_ -eq 14393} {
            $keys = @{
                '7'   = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY';
                '8'   = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG';
                '50'  = 'JCKRF-N37P4-C2D82-9YXRT-4M63B';
                '125' = 'DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ';
                '126' = 'QFFDN-GRT3P-VKWWX-X7T3R-8B639';
                '145' = '2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG';
                '146' = 'PTXN8-JFHJM-4WC78-MPCBR-9W4KR';
                '168' = 'VP34G-4NPPG-79JTQ-864T4-R3MQX';
            }
        }
        {$_ -eq 17763} {
            $keys = @{
                '7'   = 'N69G4-B89J2-4G8F4-WWYCC-J464C';
                '8'   = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG';
                '50'  = 'WVDHN-86M7X-466P6-VHXV7-YY726';
                '125' = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D';
                '126' = '92NFX-8DJQP-P6BBQ-THF9C-7CG2H';
                '145' = '6NMRW-2C8FM-D24W7-TQWMY-CWH2D';
                '146' = 'N2KJX-J94YW-TQVFB-DG9YT-724CC';
                '168' = 'FDNH6-VW9RW-BXPJ7-4XTYG-239TB';
            }
        }
        {$_ -gt 17763} {
            $keys = @{
                '125' = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D';
                '126' = '92NFX-8DJQP-P6BBQ-THF9C-7CG2H';
            }
        }
        {$_ -eq 20348} {
            $keys = @{
                '7'   = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H';
                '8'   = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33';
                '145' = 'QFND9-D3Y9C-J3KKY-6RPVP-2DPYV';
                '146' = '67KN8-4FYJW-2487Q-MQ2J7-4C4RG';
                '168' = '6N379-GGTMK-23C6M-XVVTC-CKFRQ';
            }
        }
    }

    if($null -ne $keys[$SkuId]) {
        Return $keys[$SkuId]
    }

    Return $null
}

function StartServices {
    $services = 'sppsvc', 'ClipSVC', 'wlidsvc'
    foreach($service in $services) {
        $svc = (Get-Service -Name $service -ErrorAction Stop)
        if($svc.Status -ne 'Running') {
            Write-Host "Starting required service ${service}..."
            $svc | Start-Service -ErrorAction Stop
        }
    }
}

function InstallProductKey($Key) {
    Return Invoke-CimMethod -Query 'SELECT Version FROM SoftwareLicensingService' `
        -MethodName 'InstallProductKey' -Arguments @{ProductKey=$Key} `
        -ErrorAction Stop
}

function GetProductKeyChannel {
    $query = "SELECT ProductKeyChannel FROM SoftwareLicensingProduct WHERE $slc_where"

    try {
        Return [String](Get-CimInstance -Query $query -ErrorAction Stop).ProductKeyChannel
    } catch {
        Return "Unknown"
    }
}

function SetKMS38Machine {
    $query = "SELECT ID FROM SoftwareLicensingProduct WHERE $slc_where"

    Return Invoke-CimMethod -Query $query `
        -MethodName 'SetKeyManagementServiceMachine' `
        -Arguments @{MachineName='127.0.0.1'} -ErrorAction Stop
}

function RefreshLicenseStatus {
    Return Invoke-CimMethod -Query 'SELECT Version FROM SoftwareLicensingService' `
        -MethodName 'RefreshLicenseStatus' -ErrorAction Stop
}

function ActivateProduct {
    $query = "SELECT ID FROM SoftwareLicensingProduct WHERE $slc_where"

    Return Invoke-CimMethod -Query $query `
        -MethodName 'Activate' -ErrorAction Stop
}

function GetLicenseStatus {
    $query = "SELECT LicenseStatus FROM SoftwareLicensingProduct WHERE $slc_where"

    try {
        Return [int](Get-CimInstance -Query $query -ErrorAction Stop).LicenseStatus
    } catch {
        Return 0
    }
}

function Cleanup {
    $clanup_files = 'gatherosstatemodified.exe', 'GenuineTicket.xml'

    foreach($file in $clanup_files) {
        if((Test-Path -Path $file) -eq $true) {
            Remove-Item -Path $file -Force
        }
    }
}

function CleanupAndExit {
    Cleanup
    Pop-Location
    Exit $args[0]
}

Push-Location -Path $PSScriptRoot

foreach($file in $required_files) {
    if((Test-Path -Path "$PSScriptRoot\$file") -eq $false) {
        Write-Error "Required file $file is missing"
        CleanupAndExit 0x80070002
    }
}

$kms38 = $false
$key = $null

$os = (Get-CimInstance -Class 'Win32_OperatingSystem')
$sku = $os.OperatingSystemSKU
$build = $os.BuildNumber

if($build -lt 10240) {
    Write-Error 'This script requires Windows build greater than or equal to 10240'
    CleanupAndExit 0x8007000a
}

Cleanup

Write-Host @"
============================================================
nt10RentFree $version
https://github.com/Gamers-Against-Weed/nt10RentFree
============================================================

"@ -ForegroundColor Yellow -BackgroundColor Black

if($null -ne $ProvidedKey) {
    $key = $ProvidedKey
}

if(($null -eq $key) -and ($kms38_forced -eq $false)) {
    $key = GetHwidProductKey -SkuId $sku -Build $build
}

if($null -eq $key) {
    $key = GetKms38ProductKey -SkuId $sku -Build $build
}

if($null -eq $key) {
    Write-Error 'Installed edition is not supported'
    CleanupAndExit 0x8007000a
}

Write-Host 'Checking required services...'

try {
    StartServices
} catch {
    Write-Error $_
    CleanupAndExit 0x8007042c
}

Write-Host 'Done.' -ForegroundColor Cyan -BackgroundColor Black
Write-Host

Write-Host "Installing product key ${key}..."

try {
    $null = InstallProductKey -Key $key
} catch {
    Write-Error $_
    CleanupAndExit 0x8007054f
}

if((GetProductKeyChannel) -eq "Volume:GVLK") {
    $kms38 = $true
}

Write-Host 'Done.' -ForegroundColor Cyan -BackgroundColor Black
Write-Host

if($kms38 -eq $true) {
    Write-Host 'Setting KMS machine to 127.0.0.1...'

    try {
        $null = SetKMS38Machine
    } catch {
        Write-Error $_
        CleanupAndExit 0x8007054f
    }

    Write-Host 'Done.' -ForegroundColor Cyan -BackgroundColor Black
    Write-Host
}

Write-Host 'Patching gatherosstate.exe...'

$patch = Start-Process -FilePath 'rundll32.exe' `
    -ArgumentList "$PSScriptRoot\slc.dll,PatchGatherosstate" -PassThru

Wait-Process -Id $patch.Id
if(($patch.ExitCode -ne 0) -or ((Test-Path "gatherosstatemodified.exe") -eq $false)) {
    Write-Error 'Failed to patch gatherosstate.exe'
    CleanupAndExit 0x8007054f
}

Write-Host 'Done.' -ForegroundColor Cyan -BackgroundColor Black
Write-Host

Write-Host 'Generating GenuineTicket.xml...'

$gos = Start-Process -FilePath "gatherosstatemodified.exe" -PassThru
Wait-Process -Id $gos.Id
if(($gos.ExitCode -ne 0) -or ((Test-Path "GenuineTicket.xml") -eq $false)) {
    Write-Error 'Failed to generate GenuineTicket.xml'

    if($gos.ExitCode -ne 0) {
        CleanupAndExit $gos.ExitCode
    } else {
        CleanupAndExit 0x8007054f
    }
}

Write-Host 'Done.' -ForegroundColor Cyan -BackgroundColor Black
Write-Host

try {
    Copy-Item -Path "GenuineTicket.xml" `
        -Destination "${Env:ALLUSERSPROFILE}\Microsoft\Windows\ClipSVC\GenuineTicket" `
        -Force -ErrorAction Stop
} catch {
    Write-Error $_
    CleanupAndExit 0x80070005
}

Write-Host 'Applying GenuineTicket.xml...'

ClipUp.exe -o
if($LASTEXITCODE -ne 0) {
    Write-Error 'Failed to apply GenuineTicket.xml'
    CleanupAndExit 0x8007054f
}

Write-Host 'Done.' -ForegroundColor Cyan -BackgroundColor Black
Write-Host

if($kms38 -eq $false) {
    Write-Host 'Activating Windows...'

    try {
        $null = RefreshLicenseStatus
        $null = ActivateProduct
    } catch {
        Write-Error $_
        CleanupAndExit 0x8007054f
    }
}

if((GetLicenseStatus) -ne 1) {
    Write-Error 'Failed to activate Windows'
    CleanupAndExit 0x8007054f
}

Write-Host 'Sucessfully activated Windows!' -ForegroundColor Green -BackgroundColor Black
CleanupAndExit 0
