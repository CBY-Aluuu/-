#requires -version 5.1
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Collect-PCInventory.ps1')

function Assert-Inventory {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('模拟测试失败：' + $Message) }
}

# 此函数替代真实 CIM 调用。任意意外的 CIM 类名均使测试立即失败。
function Get-CimInstance {
    [CmdletBinding()]
    param([string]$ClassName)
    if ($script:MockMissing) { throw 'TEST_ONLY: simulated unavailable CIM' }
    switch ($ClassName) {
        'Win32_ComputerSystem' {
            [pscustomobject]@{ Manufacturer = 'TEST_ONLY'; Model = 'Sample PC'; TotalPhysicalMemory = 16GB }
        }
        'Win32_Processor' { [pscustomobject]@{ Name = 'TEST_ONLY CPU' } }
        'Win32_OperatingSystem' { [pscustomobject]@{ Caption = 'TEST_ONLY OS'; Version = '1'; OSArchitecture = '64-bit' } }
        'Win32_BIOS' { [pscustomobject]@{ SerialNumber = 'TEST_ONLY-BIOS' } }
        'Win32_DiskDrive' {
            [pscustomobject]@{ Index = 0; Model = 'TEST_ONLY SSD'; SerialNumber = '000123'; Size = 512GB; InterfaceType = 'SCSI'; PNPDeviceID = 'SCSI\MOCK'; MediaType = 'Fixed hard disk media' }
            [pscustomobject]@{ Index = 1; Model = 'TEST_ONLY HDD'; SerialNumber = ''; Size = 1TB; InterfaceType = 'IDE'; PNPDeviceID = 'IDE\MOCK'; MediaType = 'Fixed hard disk media' }
            [pscustomobject]@{ Index = 2; Model = 'TEST_ONLY USB'; SerialNumber = 'DO-NOT-INCLUDE'; Size = 32GB; InterfaceType = 'USB'; PNPDeviceID = 'USBSTOR\MOCK'; MediaType = 'Removable Media' }
        }
        'Win32_NetworkAdapter' {
            [pscustomobject]@{ Index = 5; Name = 'TEST_ONLY Ethernet'; Description = 'Ethernet'; PhysicalAdapter = $true; MACAddress = '02:00:00:00:00:05'; NetConnectionID = 'Ethernet'; NetConnectionStatus = 2 }
            [pscustomobject]@{ Index = 3; Name = 'TEST_ONLY Wi-Fi'; Description = 'Wi-Fi'; PhysicalAdapter = $true; MACAddress = '02:00:00:00:00:03'; NetConnectionID = 'Wi-Fi'; NetConnectionStatus = 2 }
            [pscustomobject]@{ Index = 8; Name = 'TEST_ONLY Virtual'; Description = 'Virtual'; PhysicalAdapter = $false; MACAddress = '02:00:00:00:00:08'; NetConnectionID = 'Virtual'; NetConnectionStatus = 2 }
        }
        'Win32_NetworkAdapterConfiguration' {
            # 故意打乱顺序，验证以适配器 Index 配对，而非以结果顺序配对。
            [pscustomobject]@{ Index = 3; MACAddress = '02:00:00:00:00:03'; IPAddress = @('192.0.2.3', '2001:db8::3') }
            [pscustomobject]@{ Index = 8; MACAddress = '02:00:00:00:00:08'; IPAddress = @('192.0.2.8') }
            [pscustomobject]@{ Index = 5; MACAddress = '02:00:00:00:00:05'; IPAddress = @('192.0.2.5') }
        }
        default { throw ('TEST_ONLY: unexpected CIM class ' + $ClassName) }
    }
}

$script:MockMissing = $false
$manual = [ordered]@{
    Sequence = ''; DeviceNumber = '0007'; DeviceType = ''; AssignedTo = ''
    DeviceStatus = ''; AcquiredDate = ''; Location = 'TEST_ONLY'; Remarks = '=1+1'
}
$inventory = Get-PCInventoryData -ManualFields $manual
Assert-Inventory ($inventory.Disks.Count -eq 2) '必须保留两块机内盘并排除 USB 盘'
Assert-Inventory ($inventory.Disks[0].SerialNumber -eq '000123') '不得丢失序列号前导零'
Assert-Inventory ($inventory.Disks[1].SerialNumber -eq '未读取到，待核实') '缺失盘号必须标明待核实'
Assert-Inventory ($inventory.NetworkAdapters.Count -eq 2) '必须排除虚拟适配器'
$ethernet = $inventory.NetworkAdapters | Where-Object AdapterIndex -eq 5
$wifi = $inventory.NetworkAdapters | Where-Object AdapterIndex -eq 3
Assert-Inventory ($ethernet.IPAddresses[0] -eq '192.0.2.5' -and $ethernet.MACAddress -eq '02:00:00:00:00:05') '有线网卡必须正确配对'
Assert-Inventory ($wifi.IPAddresses.Count -eq 2 -and $wifi.IPAddresses[0] -eq '192.0.2.3') '无线网卡及多 IP 必须正确配对'
Assert-Inventory ($inventory.Manual.AssignedTo -eq '' -and $inventory.Manual.AcquiredDate -eq '') '不得自动猜测人工字段'
$row = ConvertTo-InventoryCsvRow $inventory
Assert-Inventory ($row.'备注' -eq "'=1+1") '公式输入必须转成文本'
Assert-Inventory ($row.'硬盘序列号' -match '物理硬盘0' -and $row.'硬盘序列号' -match '物理硬盘1') 'CSV 必须逐盘列出物理盘号'
foreach ($unsafe in @('=1+1', '+cmd', '-1', '@SUM(A1)', '  =1', "`t=1")) {
    Assert-Inventory ((Protect-CsvCell $unsafe).StartsWith("'")) '危险 CSV 单元格必须受保护'
}
Assert-Inventory ((Protect-CsvCell '000123') -eq '000123') '正常文本不得改写'
foreach ($placeholder in @('To Be Filled By O.E.M.', 'Default string', 'System Serial Number')) {
    Assert-Inventory ((Get-SerialOrPending $placeholder) -match '固件占位值，待核实') '固件占位值不能充当有效序列号'
}
Assert-Inventory ((Get-SerialOrPending '000123') -eq '000123') '真实序列号不应因前导零改写'

$testDirectory = Join-Path $PSScriptRoot ('test-output\TEST_ONLY_' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
$first = Export-PCInventory -Inventory $inventory -Directory $testDirectory
$second = Export-PCInventory -Inventory $inventory -Directory $testDirectory
Assert-Inventory ($first.CsvPath -ne $second.CsvPath) '多次导出不能覆盖'
$csvBytes = [System.IO.File]::ReadAllBytes($first.CsvPath)
Assert-Inventory ($csvBytes[0] -eq 239 -and $csvBytes[1] -eq 187 -and $csvBytes[2] -eq 191) 'CSV 必须有 UTF-8 BOM'
$json = Get-Content -LiteralPath $first.JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-Inventory ($json.Manual.Remarks -eq '=1+1') 'JSON 应保留原始人工值'
$imported = @(Import-Csv -LiteralPath $first.CsvPath -Encoding UTF8)
Assert-Inventory ($imported.Count -eq 1 -and $imported[0].'设备编号' -eq '0007') 'CSV 应能完整往返读取'
$overwriteRejected = $false
try { Write-NewUtf8BomFile -Path $first.CsvPath -Text 'should not overwrite' }
catch { $overwriteRejected = $true }
Assert-Inventory $overwriteRejected '同名文件必须拒绝覆盖'

$networkRejected = $false
try { [void](Resolve-LocalOutputDirectory '\\TEST_ONLY\share') }
catch { $networkRejected = $true }
Assert-Inventory $networkRejected 'UNC 路径必须在访问前拒绝'

$script:MockMissing = $true
$empty = Get-PCInventoryData -ManualFields $manual
Assert-Inventory ($empty.Disks.Count -eq 0 -and $empty.NetworkAdapters.Count -eq 0) '读取失败不得伪造硬件'
Assert-Inventory ($empty.VerificationNotes.Count -ge 7) '读取失败必须保留待核实说明'
$emptyRow = ConvertTo-InventoryCsvRow $empty
Assert-Inventory ($emptyRow.'网卡MAC' -eq '未读取到，待核实') '读取失败不得写成不存在'
Write-Host '全部模拟测试通过；未采集本机硬件或地址。'
Write-Host ('模拟结果仅供测试：' + $testDirectory)
