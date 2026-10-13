using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using PcInventory;

// Real window, fabricated records only. Does not call HardwareCollector.
internal static class UiTests
{
    private static int checks;
    private static Exception failure;
    private const BindingFlags Flags = BindingFlags.Instance | BindingFlags.NonPublic;
    private static T Field<T>(InventoryForm f, string n) { return (T)typeof(InventoryForm).GetField(n, Flags).GetValue(f); }
    private static object Invoke(InventoryForm f, string n, params object[] a) { return typeof(InventoryForm).GetMethod(n, Flags).Invoke(f, a); }
    private static Task Save(InventoryForm f) { return (Task)Invoke(f, "SaveComputer"); }
    private static void Require(bool ok, string note) { if (!ok) throw new Exception("FAILED: " + note); checks++; Console.WriteLine("PASS: " + note); }
    private static void CopyFolder(string source, string target)
    {
        Directory.CreateDirectory(target);
        foreach (string file in Directory.GetFiles(source)) File.Copy(file, Path.Combine(target, Path.GetFileName(file)), false);
        foreach (string dir in Directory.GetDirectories(source)) CopyFolder(dir, Path.Combine(target, Path.GetFileName(dir)));
    }
    private static int WorkbookRows(string path)
    {
        using (var input = File.OpenRead(path))
        using (var zip = new ZipArchive(input, ZipArchiveMode.Read))
        using (var stream = zip.GetEntry("xl/worksheets/sheet1.xml").Open())
        {
            var xml = new XmlDocument(); xml.Load(stream);
            var ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            return xml.SelectNodes("//s:sheetData/s:row", ns).Count;
        }
    }
    private static bool Fits(Control child, Control ancestor)
    {
        if (!child.Visible || child.Width < 1 || child.Height < 1) return false;
        return ancestor.ClientRectangle.Contains(new Rectangle(ancestor.PointToClient(child.PointToScreen(Point.Empty)), child.ClientSize));
    }
    private static void Shot(InventoryForm form, string directory, string name)
    {
        form.PerformLayout(); Application.DoEvents();
        using (var bitmap = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.Combine(directory, name)); }
    }
    private static void SmallWindow(InventoryForm f, string root)
    {
        Invoke(f, "SwitchPage", 0, false); f.Size = new Size(1024, 700); f.PerformLayout(); Application.DoEvents();
        var save = Field<ActionButton>(f, "save"); var status = Field<Label>(f, "status"); var details = Field<LinkLabel>(f, "statusDetails");
        Require(Fits(save, f) && Fits(save, save.Parent), "1024x700 save action remains inside client and action row");
        Require(Fits(status, f) && Fits(details, f), "1024x700 status and visible details link remain inside client");
        var date = Field<CalendarField>(f, "acquired"); ScrollableControl viewport = null;
        for (Control p = date.Parent; p != null; p = p.Parent)
        { var sc = p as ScrollableControl; if (sc != null && sc.AutoScroll) { viewport = sc; break; } }
        Require(viewport != null, "input card has a scrollable viewport");
        Shot(f, root, "02_小窗口_1024x700.png");
        viewport.ScrollControlIntoView(date); Application.DoEvents();
        Require(Fits(date, viewport), "calendar scrolls fully into viewport");
        Require(Fits(save, f) && Fits(status, f), "scrolling inputs leaves save action and status on screen");
        Shot(f, root, "03_小窗口_日历可见.png");
    }
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length != 1) return 2;
        string root = Path.GetFullPath(args[0]);
        if (Directory.Exists(root) && Directory.GetFileSystemEntries(root).Length > 0) { Console.Error.WriteLine("QA directory must be empty."); return 2; }
        Directory.CreateDirectory(root);
        Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
        using (var form = new InventoryForm(root, true))
        {
            form.ShowInTaskbar = false; form.Opacity = 0;
            form.Shown += async delegate
            {
                try
                {
                    var date = Field<CalendarField>(form, "acquired");
                    Require(date.DateValue == "", "new calendar defaults to unknown, not today");
                    form.SetPreview();
                    Require(Convert.ToString(Field<ComboBox>(form, "type").SelectedItem) == "台式机", "fabricated device type is filled");
                    Require(Convert.ToString(Field<ComboBox>(form, "state").SelectedItem) == "在用", "fabricated device status is filled");
                    Require(date.DateValue == "2023-08-15", "full date round-trips through field fill");
                    var seed = Field<InventoryRecord>(form, "current"); seed.CollectedAt = "2026-10-12T08:00:00+08:00";
                    seed.AcquiredDate = "2023-08"; Invoke(form, "FillManual", seed);
                    Require(date.DateValue == "2023-08", "legacy month date is not padded with an invented day");
                    seed.AcquiredDate = "2023"; Invoke(form, "FillManual", seed);
                    Require(date.DateValue == "2023", "legacy year date is not padded with month or day");
                    date.DateValue = ""; Field<TextBox>(form, "site").Text = "";
                    await Save(form);
                    Require(!File.Exists(new InventoryStore(root).SummaryPath), "empty required site prevents persistence");
                    Require(Field<LinkLabel>(form, "statusDetails").Visible, "validation warning exposes a details link");
                    Field<TextBox>(form, "site").Text = "TEST_ONLY示例管理站"; Field<TextBox>(form, "number").Text = "0007"; Field<TextBox>(form, "person").Text = "TEST_ONLY第一次";
                    Task first = Save(form);
                    Require(Field<bool>(form, "busy") && !Field<ActionButton>(form, "save").Enabled && !date.Enabled, "saving disables action and editing");
                    Require(Field<StatusGlyph>(form, "statusGlyph").State == "working", "working feedback remains active");
                    await Save(form);
                    form.Close();
                    Require(!form.IsDisposed && form.Visible, "close is prevented while persistence is busy");
                    await first;
                    var saved = new InventoryStore(root).LoadRecords();
                    Require(saved.Count == 1 && saved[0].AcquiredDate == "" && saved[0].DeviceNumber == "0007", "first row keeps unknown date and leading zeros; reentry adds nothing");
                    Require(Field<Label>(form, "status").Text.Contains("已追加第 1 条"), "append count is visible");
                    Require(Field<ActionButton>(form, "save").Enabled && date.Enabled, "controls reenable after save");
                    Field<TextBox>(form, "person").Text = "TEST_ONLY第二次"; date.DateValue = "2023-08"; await Save(form);
                    Field<TextBox>(form, "person").Text = "TEST_ONLY第三次"; date.DateValue = "2023-08-15"; await Save(form);
                    saved = new InventoryStore(root).LoadRecords();
                    Require(saved.Count == 3 && saved[0].AssignedTo == "TEST_ONLY第一次" && saved[1].AssignedTo == "TEST_ONLY第二次" && saved[2].AssignedTo == "TEST_ONLY第三次", "same device appends three rows and earlier values remain unchanged");
                    Require(saved[0].AcquiredDate == "" && saved[1].AcquiredDate == "2023-08" && saved[2].AcquiredDate == "2023-08-15", "all three original date precisions persist");
                    Require(new HashSet<string>(new[] { saved[0].RecordId, saved[1].RecordId, saved[2].RecordId }).Count == 3, "each saved row has a distinct record ID");
                    Require(Directory.GetFiles(Path.Combine(root, "备份"), "*.zip").Length == 3, "three saves create three independent complete backups");
                    Require(WorkbookRows(new InventoryStore(root).SummaryPath) == 4, "workbook has header plus all three same-device rows");
                    form.ClientSize = new Size(1216, 820); Shot(form, root, "01_正常采集界面.png");
                    using (var locked = File.Open(new InventoryStore(root).SummaryPath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        Field<TextBox>(form, "person").Text = "TEST_ONLY第四次"; Field<TextBox>(form, "remarks").Text = "TEST_ONLY总表占用"; await Save(form);
                        string message = Field<Label>(form, "status").Text;
                        Require(message.Contains("已追加第 4 条") && message.Contains("总表待更新") && message.Contains("完整备份已保存"), "Excel lock separately reports append success, workbook pending and backup success");
                        Require(Field<LinkLabel>(form, "statusDetails").Visible && Field<string>(form, "fullStatus").Contains("不要为更新总表而重复保存"), "lock explanation is reachable through visible details");
                    }
                    Require(new InventoryStore(root).LoadRecords().Count == 4 && Directory.GetFiles(Path.Combine(root, "备份"), "*.zip").Length == 4, "Excel lock preserves fourth row and fourth complete backup");
                    new InventoryStore(root).RebuildSummary();
                    Require(WorkbookRows(new InventoryStore(root).SummaryPath) == 5, "rebuild after unlock includes all four records");
                    SmallWindow(form, root);
                    form.ClientSize = new Size(1216, 820); Invoke(form, "SwitchPage", 1, false);
                    Require(Field<DataGridView>(form, "history").Rows.Count == 4, "history displays all repeated-device records");
                    Field<TextBox>(form, "historyFilter").Text = "第三次";
                    Require(Field<DataGridView>(form, "history").Rows.Count == 1 && Convert.ToString(Field<DataGridView>(form, "history").Rows[0].Cells[0].Value) == "3", "history filtering preserves original sequence");
                    Field<TextBox>(form, "historyFilter").Text = ""; Shot(form, root, "04_历史记录界面.png");
                    Invoke(form, "SwitchPage", 2, false);
                    Require(Field<ActionButton>(form, "backup").Visible && Field<ActionButton>(form, "restore").Visible, "backup page exposes backup and restore actions");
                    Shot(form, root, "05_备份恢复界面.png");
                }
                catch (Exception ex) { failure = ex; }
                finally { typeof(InventoryForm).GetField("busy", Flags).SetValue(form, false); form.Close(); }
            };
            Application.Run(form);
        }
        try
        {
            if (failure != null) throw failure;
            string relocated = root + "_换电脑后的文件夹"; CopyFolder(root, relocated);
            var moved = new InventoryStore(relocated);
            Require(moved.LoadRecords().Count == 4, "moving complete folder retains four prior records");
            var fifth = moved.Save(new InventoryRecord { DeviceKey = "TEST_ONLY_B", ManufacturerModel = "TEST_ONLY另一台电脑", Site = "另一示例站", Location = "办公室", AssignedTo = "TEST_ONLY公用" });
            Require(fifth.Count == 5 && fifth.BackupUpdated && WorkbookRows(moved.SummaryPath) == 6, "next machine after relocation appends fifth row and backup");
            Require(new InventoryStore(root).LoadRecords().Count == 4, "relocated-folder work leaves original QA records unchanged");
            Console.WriteLine("ALL " + checks + " UI AND PORTABILITY CHECKS PASSED");
            Console.WriteLine("Fabricated QA files: " + root); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
