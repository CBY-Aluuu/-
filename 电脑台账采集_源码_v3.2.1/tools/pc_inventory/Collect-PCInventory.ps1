#requires -version 5.1
<#
.SYNOPSIS
仅在本机读取电脑硬件资料，导出本地 JSON 与带 BOM 的 UTF-8 CSV。
#>
[CmdletBinding()]
param(
    [string]$OutputDirectory,
    [switch]$NonInteractive,
    [string]$Sequence,
    [string]$DeviceNumber,
    [string]$DeviceType,
    [string]$AssignedTo,
    [string]$DeviceStatus,
    [string]$AcquiredDate,
    [string]$Location,
    [string]$Remarks
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:InventoryVersion = '0.1.0'

function Get-OptionalProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Read-LocalInventoryCim {
    param([string]$ClassName, [System.Collections.Generic.List[string]]$Issues)
    try {
        # 不指定 ComputerName/CimSession：只读取本机，不查询远程设备。
        @(Get-CimInstance -ClassName $ClassName -ErrorAction Stop)
    }
    catch {
        $Issues.Add(('读取 {0} 失败，相关字段待核实。' -f $ClassName))
    }
}

function Get-TextOrPending {
    param($Value)
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '未读取到，待核实' }
    return $text.Trim()
}

function Get-SerialOrPending {
    param($Value)
    $serial = Get-TextOrPending $Value
    if ($serial -match '^(To Be Filled By O[.]?E[.]?M[.]?|Default string|System Serial Number|None|Unknown|Not Specified|Not Available|Not Applicable|N/?A)$') {
        return ('固件占位值，待核实（原值：{0}）' -f $serial)
    }
    if ($serial -match '^(0+|F+)$') { return ('疑似固件占位值，待核实（原值：{0}）' -f $serial) }
    return $serial
}

function Resolve-LocalOutputDirectory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ($Path -match '^[\\/]{2}' -or $Path -match '::') {
        throw '结果只能写入本地磁盘，不能使用 UNC、网络共享或提供程序路径。'
    }
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath -notmatch '^[A-Za-z]:[\\/]') {
        throw '结果路径必须位于本机盘符，例如 C:\台账采集结果。'
    }
    $root = [System.IO.Path]::GetPathRoot($fullPath)
    $driveName = $root.Substring(0, 1)
    $psDrive = Get-PSDrive -Name $driveName -PSProvider FileSystem -ErrorAction Stop
    $displayRoot = [string](Get-OptionalProperty $psDrive 'DisplayRoot')
    if ($displayRoot -match '^[\\/]{2}' -or $psDrive.Root -match '^[\\/]{2}') {
        throw '不能把结果写到映射网络驱动器。请改用本地文件夹。'
    }
    $drive = New-Object System.IO.DriveInfo($root)
    if ($drive.DriveType -eq [System.IO.DriveType]::Network) {
        throw '不能把结果写到映射网络驱动器。请改用本地文件夹。'
    }
    if (-not $drive.IsReady) { throw '目标磁盘不可用。' }

    # 拒绝经过目录联接/符号链接，避免路径表面在本地、实际指向其他位置。
    $cursor = $fullPath
    while (-not [string]::IsNullOrWhiteSpace($cursor)) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw '结果路径包含目录联接或符号链接，请选择普通本地目录。'
            }
            if (-not $item.PSIsContainer) { throw '结果路径或其上级不是文件夹。' }
        }
        $parent = [System.IO.Directory]::GetParent($cursor)
        if ($null -eq $parent) { break }
        $cursor = $parent.FullName
    }
    return $fullPath
}

function Get-PCInventoryData {
    param([System.Collections.IDictionary]$ManualFields)
    $issues = New-Object 'System.Collections.Generic.List[string]'
    $computers = @(Read-LocalInventoryCim -ClassName 'Win32_ComputerSystem' -Issues $issues)
    $processors = @(Read-LocalInventoryCim -ClassName 'Win32_Processor' -Issues $issues)
    $systems = @(Read-LocalInventoryCim -ClassName 'Win32_OperatingSystem' -Issues $issues)
    $biosItems = @(Read-LocalInventoryCim -ClassName 'Win32_BIOS' -Issues $issues)
    $diskItems = @(Read-LocalInventoryCim -ClassName 'Win32_DiskDrive' -Issues $issues)
    $adapterItems = @(Read-LocalInventoryCim -ClassName 'Win32_NetworkAdapter' -Issues $issues)
    $configurationItems = @(Read-LocalInventoryCim -ClassName 'Win32_NetworkAdapterConfiguration' -Issues $issues)

    $computer = $computers | Select-Object -First 1
    $system = $systems | Select-Object -First 1
    $bios = $biosItems | Select-Object -First 1
    $memoryBytes = Get-OptionalProperty $computer 'TotalPhysicalMemory'
    $memoryGiB = $null
    if ($null -ne $memoryBytes -and [double]$memoryBytes -gt 0) {
        $memoryGiB = [math]::Round(([double]$memoryBytes / 1GB), 2)
    }
    else { $issues.Add('内存容量未读取到，待核实。') }

    $disks = New-Object 'System.Collections.Generic.List[object]'
    $excludedDiskCount = 0
    foreach ($disk in $diskItems) {
        $interface = [string](Get-OptionalProperty $disk 'InterfaceType')
        $pnpId = [string](Get-OptionalProperty $disk 'PNPDeviceID')
        $mediaType = [string](Get-OptionalProperty $disk 'MediaType')
        if ($interface -eq 'USB' -or $pnpId -match '^USB(STOR)?\\' -or $mediaType -match 'Removable') {
            $excludedDiskCount++
            continue
        }
        $size = Get-OptionalProperty $disk 'Size'
        $sizeGiB = $null
        if ($null -ne $size -and [double]$size -gt 0) { $sizeGiB = [math]::Round(([double]$size / 1GB), 2) }
        $model = Get-TextOrPending (Get-OptionalProperty $disk 'Model')
        $serial = Get-SerialOrPending (Get-OptionalProperty $disk 'SerialNumber')
        $diskCheck = '请现场确认是机内硬盘；非 USB 外置盘可能不能自动区分。'
        if ($model -match 'Virtual|VMware|VBOX|QEMU|Xen') { $diskCheck = '疑似虚拟磁盘，待现场核实。' }
        if ($serial -eq '未读取到，待核实') { $issues.Add(('硬盘 {0} 的序列号未读取到，待核实。' -f (Get-OptionalProperty $disk 'Index'))) }
        $disks.Add([pscustomobject][ordered]@{
            DiskIndex = Get-OptionalProperty $disk 'Index'
            Model = $model
            SerialNumber = $serial
            CapacityGiB = $sizeGiB
            Interface = Get-TextOrPending $interface
            Verification = $diskCheck
        })
    }
    if ($disks.Count -eq 0) { $issues.Add('未读取到可列入的机内硬盘资料，不代表没有硬盘，待核实。') }
    if ($excludedDiskCount -gt 0) { $issues.Add(('排除了 {0} 个明显 USB/可移动存储设备；请核对是否有应列入资产的外置设备。' -f $excludedDiskCount)) }

    $adapters = New-Object 'System.Collections.Generic.List[object]'
    $excludedAdapterCount = 0
    foreach ($adapter in $adapterItems) {
        $name = [string](Get-OptionalProperty $adapter 'Name')
        $description = [string](Get-OptionalProperty $adapter 'Description')
        $physical = Get-OptionalProperty $adapter 'PhysicalAdapter'
        if (($null -ne $physical -and -not [bool]$physical) -or ($name + ' ' + $description) -match 'Virtual|VMware|Hyper-V|VBOX|TAP-Windows|VPN|Tunnel|Loopback|Bluetooth|WAN Miniport') {
            $excludedAdapterCount++
            continue
        }
        $index = Get-OptionalProperty $adapter 'Index'
        $configuration = $configurationItems | Where-Object { (Get-OptionalProperty $_ 'Index') -eq $index } | Select-Object -First 1
        $mac = [string](Get-OptionalProperty $adapter 'MACAddress')
        if ([string]::IsNullOrWhiteSpace($mac)) { $mac = [string](Get-OptionalProperty $configuration 'MACAddress') }
        $addresses = @((Get-OptionalProperty $configuration 'IPAddress') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        $adapterCheck = '自动读取当时地址；请核对实际使用的网络，地址可能随网络变化。'
        if ($null -eq $physical) { $adapterCheck += ' 物理网卡标记未读取到，是否应纳入待核实。' }
        if ($addresses.Count -eq 0) { $adapterCheck += ' 未读到 IP，可能未连接或未配置，待核对。' }
        if ([string]::IsNullOrWhiteSpace($mac)) { $issues.Add(('网卡 {0} 的 MAC 未读取到，待核实。' -f $name)) }
        $adapters.Add([pscustomobject][ordered]@{
            AdapterIndex = $index
            AdapterName = Get-TextOrPending $name
            ConnectionName = [string](Get-OptionalProperty $adapter 'NetConnectionID')
            MACAddress = Get-TextOrPending $mac
            IPAddresses = @($addresses)
            ConnectionStatusCode = Get-OptionalProperty $adapter 'NetConnectionStatus'
            Verification = $adapterCheck
        })
    }
    if ($adapters.Count -eq 0) { $issues.Add('未读取到可确认的物理网卡资料，不代表没有网卡，待核实。') }
    if ($excludedAdapterCount -gt 0) { $issues.Add('已过滤明显虚拟、隧道、VPN、蓝牙或非物理适配器；特殊用途适配器请人工另行核对。') }

    [pscustomobject][ordered]@{
        SchemaVersion = $script:InventoryVersion
        CollectedAt = [DateTimeOffset]::Now.ToString('o')
        Scope = '本机只读采集；自动字段须现场核对，人工字段不能由系统信息代替。'
        Manual = [pscustomobject]$ManualFields
        Quantity = 1
        Hardware = [pscustomobject][ordered]@{
            Manufacturer = Get-TextOrPending (Get-OptionalProperty $computer 'Manufacturer')
            Model = Get-TextOrPending (Get-OptionalProperty $computer 'Model')
            BIOSDeviceSerial = Get-SerialOrPending (Get-OptionalProperty $bios 'SerialNumber')
            CPU = @($processors | ForEach-Object { Get-TextOrPending (Get-OptionalProperty $_ 'Name') })
            MemoryGiB = $memoryGiB
            OperatingSystem = Get-TextOrPending (Get-OptionalProperty $system 'Caption')
            OSVersion = Get-TextOrPending (Get-OptionalProperty $system 'Version')
            OSArchitecture = Get-TextOrPending (Get-OptionalProperty $system 'OSArchitecture')
        }
        Disks = @($disks.ToArray())
        NetworkAdapters = @($adapters.ToArray())
        VerificationNotes = @($issues.ToArray())
    }
}

function Protect-CsvCell {
    param($Value)
    $text = [string]$Value
    # Excel 可能将 =、+、-、@ 开头的值解释为公式；JSON 保留原始值。
    if ($text -match '^[\s\p{Cf}]*[=+\-@]' -or $text -match '^[\t\r\n]') { return "'" + $text }
    return $text
}

function ConvertTo-InventoryCsvRow {
    param([Parameter(Mandatory = $true)]$Inventory)
    $diskLines = @($Inventory.Disks | ForEach-Object {
        '物理硬盘{0}：{1}；序列号：{2}' -f $_.DiskIndex, $_.Model, $_.SerialNumber
    })
    $macLines = @($Inventory.NetworkAdapters | ForEach-Object {
        '网卡{0} [{1}]：{2}' -f $_.AdapterIndex, $_.AdapterName, $_.MACAddress
    })
    $ipLines = @($Inventory.NetworkAdapters | ForEach-Object {
        $ipText = '未读取到，待核实（可能未连接）'
        if ($_.IPAddresses.Count -gt 0) { $ipText = $_.IPAddresses -join ' / ' }
        '网卡{0} [{1}]：{2}' -f $_.AdapterIndex, $_.AdapterName, $ipText
    })
    $pairLines = @($Inventory.NetworkAdapters | ForEach-Object {
        $ipText = '未读取到，待核实（可能未连接）'
        if ($_.IPAddresses.Count -gt 0) { $ipText = $_.IPAddresses -join ' / ' }
        '网卡{0} [{1}] MAC={2}；IP={3}' -f $_.AdapterIndex, $_.AdapterName, $_.MACAddress, $ipText
    })
    $diskText = '未读取到，待核实'
    $macText = '未读取到，待核实'
    $ipText = '未读取到，待核实'
    $pairText = '未读取到，待核实'
    if ($diskLines.Count -gt 0) { $diskText = $diskLines -join ' | ' }
    if ($macLines.Count -gt 0) { $macText = $macLines -join ' | ' }
    if ($ipLines.Count -gt 0) { $ipText = $ipLines -join ' | ' }
    if ($pairLines.Count -gt 0) { $pairText = $pairLines -join ' | ' }
    $cpuText = '未读取到，待核实'
    if ($Inventory.Hardware.CPU.Count -gt 0) { $cpuText = $Inventory.Hardware.CPU -join ' / ' }
    $memoryText = '未读取到，待核实'
    if ($null -ne $Inventory.Hardware.MemoryGiB) { $memoryText = [string]$Inventory.Hardware.MemoryGiB }
    $verification = @('自动识别的品牌型号、硬盘、网卡及整机序列号须人工核对。') + @($Inventory.VerificationNotes)
    $verification += @($Inventory.Disks | ForEach-Object { '硬盘{0}：{1}' -f $_.DiskIndex, $_.Verification })
    $verification += @($Inventory.NetworkAdapters | ForEach-Object { '网卡{0}：{1}' -f $_.AdapterIndex, $_.Verification })
    $raw = [ordered]@{
        '序号' = $Inventory.Manual.Sequence
        '设备编号' = $Inventory.Manual.DeviceNumber
        '设备类型' = $Inventory.Manual.DeviceType
        '品牌及型号' = $Inventory.Hardware.Manufacturer + ' / ' + $Inventory.Hardware.Model
        '使用人员' = $Inventory.Manual.AssignedTo
        '设备状态' = $Inventory.Manual.DeviceStatus
        '硬盘序列号' = $diskText
        '网卡MAC' = $macText
        '取得日期' = $Inventory.Manual.AcquiredDate
        '备注' = $Inventory.Manual.Remarks
        'IP地址（如有）' = $ipText
        '数量' = $Inventory.Quantity
        '所在位置' = $Inventory.Manual.Location
        '网卡IP与MAC配对明细' = $pairText
        '整机序列号（BIOS）' = $Inventory.Hardware.BIOSDeviceSerial
        'CPU' = $cpuText
        '内存（GiB）' = $memoryText
        '操作系统' = $Inventory.Hardware.OperatingSystem + ' / ' + $Inventory.Hardware.OSVersion + ' / ' + $Inventory.Hardware.OSArchitecture
        '采集时间' = $Inventory.CollectedAt
        '待核对说明' = $verification -join ' | '
    }
    $safe = [ordered]@{}
    foreach ($key in $raw.Keys) { $safe[$key] = Protect-CsvCell $raw[$key] }
    return [pscustomobject]$safe
}

function Write-NewUtf8BomFile {
    param([string]$Path, [string]$Text)
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $encoding = New-Object System.Text.UTF8Encoding($true)
        $writer = New-Object System.IO.StreamWriter($stream, $encoding)
        try { $writer.Write($Text) }
        finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Export-PCInventory {
    param([Parameter(Mandatory = $true)]$Inventory, [Parameter(Mandatory = $true)][string]$Directory)
    $safeDirectory = Resolve-LocalOutputDirectory -Path $Directory
    [void][System.IO.Directory]::CreateDirectory($safeDirectory)
    $baseName = 'PC_{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), ([Guid]::NewGuid().ToString('N').Substring(0, 12))
    $jsonPath = Join-Path $safeDirectory ($baseName + '.json')
    $csvPath = Join-Path $safeDirectory ($baseName + '.csv')
    $json = $Inventory | ConvertTo-Json -Depth 10
    $csv = (ConvertTo-InventoryCsvRow -Inventory $Inventory | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
    Write-NewUtf8BomFile -Path $jsonPath -Text ($json + "`r`n")
    Write-NewUtf8BomFile -Path $csvPath -Text ($csv + "`r`n")
    [pscustomobject]@{ JsonPath = $jsonPath; CsvPath = $csvPath }
}

# 点引用只载入函数，以便使用模拟对象测试，不执行本机采集。
if ($MyInvocation.InvocationName -ne '.') {
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $PSScriptRoot '结果' }
    [void](Resolve-LocalOutputDirectory -Path $OutputDirectory)
    $manual = [ordered]@{
        Sequence = $Sequence; DeviceNumber = $DeviceNumber; DeviceType = $DeviceType
        AssignedTo = $AssignedTo; DeviceStatus = $DeviceStatus; AcquiredDate = $AcquiredDate
        Location = $Location; Remarks = $Remarks
    }
    if (-not $NonInteractive) {
        Write-Host '电脑资产离线采集初版：只读取本机，结果仅保存到指定本地文件夹。'
        Write-Host '下列内容可留空，之后人工补齐；不得把主机名、登录账户或系统安装日期当成资产资料。'
        $labels = [ordered]@{
            Sequence = '序号'; DeviceNumber = '设备编号'; DeviceType = '设备类型（台式机/笔记本等）'
            AssignedTo = '实际使用人员'; DeviceStatus = '设备状态（在用/闲置/待报废等）'
            AcquiredDate = '取得日期（查采购或资产记录）'; Location = '所在位置（站点、房间）'; Remarks = '备注'
        }
        foreach ($key in $labels.Keys) {
            if (-not $PSBoundParameters.ContainsKey($key)) { $manual[$key] = Read-Host $labels[$key] }
        }
    }
    $inventory = Get-PCInventoryData -ManualFields $manual
    $result = Export-PCInventory -Inventory $inventory -Directory $OutputDirectory
    Write-Host '已完成本机读取。请核对自动字段并补齐人工字段；未读取到不等于不存在。'
    Write-Host ('CSV：' + $result.CsvPath)
    Write-Host ('JSON：' + $result.JsonPath)
}
