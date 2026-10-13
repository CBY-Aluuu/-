param([string]$OutputDirectory = '_analysis\ios_qa')
$ErrorActionPreference = 'Stop'
$compilerRoot = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319'
$sourceDirectory = $PSScriptRoot
$targetDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
[System.IO.Directory]::CreateDirectory($targetDirectory) | Out-Null
$compileArguments = @(
    '/nologo', '/langversion:5', '/target:winexe', '/platform:anycpu', '/optimize+',
    ('/out:' + (Join-Path $targetDirectory '电脑台账采集.exe')),
    '/main:PcInventory.Ios.IosProgram',
    ('/win32manifest:' + (Join-Path $sourceDirectory 'app.manifest')),
    '/r:System.dll', '/r:System.Core.dll', '/r:System.Windows.Forms.dll', '/r:System.Drawing.dll',
    '/r:System.Management.dll', '/r:System.Runtime.Serialization.dll', '/r:System.IO.Compression.dll',
    '/r:System.IO.Compression.FileSystem.dll', '/r:System.Xaml.dll',
    ('/r:' + (Join-Path $compilerRoot 'WPF\WindowsBase.dll')),
    ('/r:' + (Join-Path $compilerRoot 'WPF\PresentationCore.dll')),
    ('/r:' + (Join-Path $compilerRoot 'WPF\PresentationFramework.dll')),
    ('/resource:' + (Join-Path $sourceDirectory 'IosTheme.xaml') + ',IosTheme.xaml'),
    ('/resource:' + (Join-Path $sourceDirectory 'IosWindow.xaml') + ',IosWindow.xaml')
)
$appIcon = Join-Path $sourceDirectory 'app.ico'
if (Test-Path -LiteralPath $appIcon) { $compileArguments += ('/win32icon:' + $appIcon) }
foreach ($asset in @('developer-avatar.jpg', 'donate-alipay.jpg', 'donate-wechat.jpg')) {
    $compileArguments += ('/resource:' + (Join-Path $sourceDirectory ('Assets\' + $asset)) + ',Assets.' + $asset)
}
foreach ($file in @('IosProgram.cs', 'PortableInstanceGuard.cs', 'IosWindow.cs', 'IosControls.cs', 'InventoryStore.cs', 'InventoryWorkbookReader.cs', 'InventoryPresentation.cs', 'HardwareCollector.cs')) { $compileArguments += (Join-Path $sourceDirectory $file) }
& (Join-Path $compilerRoot 'csc.exe') @compileArguments
if ($LASTEXITCODE -ne 0) { throw '编译失败。' }
Write-Output ('已编译：' + (Join-Path $targetDirectory '电脑台账采集.exe'))
