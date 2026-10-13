电脑台账采集 · 源码包 v3.2.1
开发者：CBY · Alu
GitHub：https://github.com/CBY-Aluuu
联系邮箱：1140357618@qq.com

一、当前版本
本源码包对应 Windows 桌面程序 3.2.1.0：
现场信息合并为一个“位置”，兼容旧台账；侧栏包含“关于”，内置头像、
支付宝和微信打赏图片，GitHub 链接已更新为 CBY-Aluuu。

二、编译方法
1. 将压缩包完整解压到一个可写文件夹。
2. 在 64 位 Windows 中双击根目录 build.cmd。
3. 编译成功后，程序位于 build 文件夹中的“电脑台账采集.exe”。

所需环境：.NET Framework 4.5 或更新的 4.x 版本及 WPF 组件。
构建脚本使用 C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe，
不依赖 Visual Studio、NuGet 或在线下载。32 位 Windows 需自行调整编译器路径。
如果提示找不到编译器，请先确认上述路径和 .NET Framework 安装情况。

也可以在源码包根目录打开 PowerShell，执行：
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\pc_inventory_portable\build-ios.ps1 -OutputDirectory .\build

build.cmd 中的 ExecutionPolicy Bypass 仅对本次编译进程生效，不更改系统策略。
包内 PowerShell 脚本采用 UTF-8 BOM，以便 Windows PowerShell 5.1 正确读取中文。

三、代码目录
tools\pc_inventory_portable\
  build-ios.ps1                当前 WPF 程序的构建入口
  IosProgram.cs                程序入口、版本、预览模式
  IosWindow.xaml / .cs         界面布局与事件逻辑
  IosTheme.xaml                界面样式
  IosControls.cs               日期等自定义控件
  InventoryStore.cs            本地记录、总表导出、备份与恢复
  InventoryWorkbookReader.cs   Excel 台账读取与兼容处理
  InventoryPresentation.cs     位置和硬件等显示格式
  HardwareCollector.cs         本机硬件采集
  PortableInstanceGuard.cs     防止同目录重复启动
  Assets\                     头像、支付宝及微信打赏原图
  app.ico / app.manifest       程序图标和启动清单
  *Tests.cs                    已有测试源码

该目录也保留早期 WinForms 界面源码（Program.cs、MainWindow.cs、ModernControls.cs），
以及相应测试和说明。默认构建只使用 build-ios.ps1 中明确列出的当前版本源文件，
请勿将所有 .cs 一起编译，否则会混入多个程序和测试入口。

tools\pc_inventory\ 为早期 PowerShell 采集工具及其测试、说明，供参考。
根目录“使用方法.txt”为当前软件的使用说明；旧目录内说明保留作为历史资料。

四、预览与已有测试
编译后可运行以下命令生成示例界面图片，不采集本机硬件：
.\build\电脑台账采集.exe --preview .\preview

*Tests.cs 为独立测试程序，需要分别编译入口并引用相应实现或应用程序集。
测试输出应使用单独文件夹。默认 build.cmd 不执行测试。
UpgradeCheck.cs 会重建传入目录的总表，仅用于复制出来的台账验证样本。

五、打包内容
包含源码、构建脚本、图标、图片资源、测试源码和说明。
不包含实际台账、记录、备份、缓存、编译产物或调试输出。
文件清单及 SHA-256.txt 提供包内文件的相对路径与 SHA-256 校验值。
