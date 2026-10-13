using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PcInventory
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                if (args.Length == 2 && args[0] == "--check-local")
                {
                    var sample = HardwareCollector.Collect();
                    sample.Site = "本机功能验证"; sample.Location = "测试记录"; sample.AssignedTo = "待核实";
                    var store = new InventoryStore(args[1]);
                    var result = store.Save(sample);
                    File.WriteAllText(Path.Combine(args[1], "验证状态.txt"),
                        "仅在本机保存验证记录。\r\n记录数：" + result.Count + "\r\n汇总已更新：" + result.SummaryUpdated +
                        "\r\n型号可读：" + (!String.IsNullOrWhiteSpace(sample.ManufacturerModel) && !sample.ManufacturerModel.Contains("待核实")) +
                        "\r\nCPU可读：" + (!String.IsNullOrWhiteSpace(sample.Cpu) && !sample.Cpu.Contains("待核实")), System.Text.Encoding.UTF8);
                    return result.SummaryUpdated ? 0 : 2;
                }
                if (args.Length == 2 && args[0] == "--preview")
                {
                    Directory.CreateDirectory(args[1]);
                    using (var form = new InventoryForm(args[1], true))
                    {
                        form.ShowInTaskbar = false; form.Opacity = 0; form.Show(); Application.DoEvents();
                        form.SetPreview(); Application.DoEvents();
                        using (var bitmap = new Bitmap(form.Width, form.Height))
                        {
                            form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                            bitmap.Save(Path.Combine(args[1], "界面预览.png"));
                        }
                    }
                    return 0;
                }
                Application.Run(new InventoryForm(AppDomain.CurrentDomain.BaseDirectory, false));
                return 0;
            }
            catch (Exception e)
            {
                if (args.Length > 0)
                {
                    try { File.WriteAllText(Path.Combine(args[1], "启动错误.txt"), e.ToString()); } catch { }
                    return 1;
                }
                MessageBox.Show("暂时无法打开工具。请先把压缩包全部解压到普通本地文件夹或单位允许的U盘，再双击程序。\r\n\r\n" + e.Message,
                    "电脑台账采集", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 1;
            }
        }

        private static void PrepareHandles(Control control)
        {
            IntPtr handle = control.Handle;
            foreach (Control child in control.Controls) PrepareHandles(child);
        }
    }

}
