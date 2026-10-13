using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

[assembly: System.Reflection.AssemblyTitle("电脑台账采集")]
[assembly: System.Reflection.AssemblyDescription("离线电脑台账 · 开发者：CBY")]
[assembly: System.Reflection.AssemblyCompany("CBY")]
[assembly: System.Reflection.AssemblyProduct("电脑台账采集 · CBY")]
[assembly: System.Reflection.AssemblyCopyright("CBY")]
[assembly: System.Reflection.AssemblyVersion("3.2.1.0")]
[assembly: System.Reflection.AssemblyFileVersion("3.2.1.0")]

namespace PcInventory.Ios
{
    internal static class IosProgram
    {
        [STAThread]
        private static int Main(string[] args)
        {
            try
            {
                InventoryWindow.EnsureTheme();
                if (args.Length == 2 && args[0] == "--preview")
                {
                    Directory.CreateDirectory(args[1]);
                    var window = new InventoryWindow(args[1], true) { ShowInTaskbar = false, ShowActivated = false, Left = -20000, Top = -20000, WindowStartupLocation = WindowStartupLocation.Manual, Width = 1220, Height = 900 };
                    window.Show(); window.SetPreview(); window.UpdateLayout();
                    window.Dispatcher.Invoke(DispatcherPriority.Render, new Action(delegate { }));
                    Capture(window, Path.Combine(args[1], "界面预览.png")); window.SwitchPage(2, false); window.UpdateLayout(); Capture(window, Path.Combine(args[1], "备份页面.png"));
                    window.SwitchPage(3, false); window.UpdateLayout(); Capture(window, Path.Combine(args[1], "关于页面.png"));
                    window.Close(); return 0;
                }
                if (args.Length == 2 && args[0] == "--check-local")
                {
                    var record = HardwareCollector.Collect(); record.Site = "本机功能验证"; record.Location = "测试记录"; record.AssignedTo = "待核实";
                    var result = new InventoryStore(args[1]).Save(record); File.WriteAllText(Path.Combine(args[1], "验证状态.txt"), "记录数：" + result.Count + "\r\n总表已更新：" + result.SummaryUpdated + "\r\n备份已完成：" + result.BackupUpdated); return result.SummaryUpdated && result.BackupUpdated ? 0 : 2;
                }
                string directory = AppDomain.CurrentDomain.BaseDirectory;
                using (var instance = PortableInstanceGuard.Acquire(directory))
                {
                    if (!instance.IsOwner)
                    {
                        PortableInstanceGuard.ActivateExistingInstance(directory);
                        return 0;
                    }
                    if (PortableInstanceGuard.ActivateLegacyInstance(directory)) return 0;
                    var form = new InventoryWindow(directory, false);
                    IntPtr windowHandle = IntPtr.Zero;
                    form.SourceInitialized += delegate { windowHandle = new System.Windows.Interop.WindowInteropHelper(form).Handle; PortableInstanceGuard.MarkWindow(windowHandle); };
                    form.Closed += delegate { PortableInstanceGuard.UnmarkWindow(windowHandle); };
                    Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose; Application.Current.Run(form); return 0;
                }
            }
            catch (Exception e)
            {
                if (args.Length > 1) { try { Directory.CreateDirectory(args[1]); File.WriteAllText(Path.Combine(args[1], "启动错误.txt"), e.ToString()); } catch { } return 1; }
                MessageBox.Show("程序暂时无法打开。请将压缩包全部解压，再双击程序。\n\n" + e.Message, "电脑台账采集"); return 1;
            }
        }
        public static void Capture(InventoryWindow window, string path)
        {
            var visual = (FrameworkElement)window.Content; visual.UpdateLayout();
            var image = new RenderTargetBitmap((int)Math.Ceiling(visual.ActualWidth), (int)Math.Ceiling(visual.ActualHeight), 96, 96, PixelFormats.Pbgra32); image.Render(visual);
            var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(image)); using (var output = File.Create(path)) encoder.Save(output);
        }
    }
}
