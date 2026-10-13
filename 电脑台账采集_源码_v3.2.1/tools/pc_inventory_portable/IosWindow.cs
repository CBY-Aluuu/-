using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shell;
using System.Windows.Threading;
using Microsoft.Win32;

namespace PcInventory.Ios
{
    public sealed class HardwareLine { public string Label { get; set; } public string Value { get; set; } }
    public sealed class HistoryLine
    {
        public int Index { get; set; } public string Number { get; set; } public string Serial { get; set; } public string Site { get; set; } public string Location { get; set; }
        public string Person { get; set; } public string Model { get; set; } public string Time { get; set; }
        public InventoryRecord Record { get; set; }
    }

    public sealed class InventoryWindow : Window
    {
        private readonly string directory;
        private InventoryStore store;
        private readonly Grid view;
        private readonly bool preview;
        private InventoryRecord current;
        private List<InventoryRecord> records = new List<InventoryRecord>();
        private bool busy, syncing, closed;
        private Task syncTask = Task.FromResult(0);
        private readonly DispatcherTimer historyTimer;
        private bool changingWorkbook;
        private string selectedWorkbook = "电脑台账汇总.xlsx";
        private string fullStatus = "";
        private int activePage;
        public const string DeveloperGitHubUrl = "https://github.com/CBY-Aluuu";
        public const string DeveloperEmailUri = "mailto:1140357618@qq.com";
        public bool IsBusy { get { return busy; } }

        public static void EnsureTheme()
        {
            if (Application.Current == null) new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
            if (Application.Current.Resources.Contains("PrimaryButton")) return;
            using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("IosTheme.xaml"))
                Application.Current.Resources.MergedDictionaries.Add((ResourceDictionary)XamlReader.Load(input));
        }

        public InventoryWindow(string baseDirectory, bool previewMode)
        {
            EnsureTheme(); directory = Path.GetFullPath(baseDirectory); preview = previewMode;
            try {
                string preference = Path.Combine(directory, "数据", "当前台账.txt");
                if (File.Exists(preference)) {
                    string remembered = File.ReadAllText(preference).Trim();
                    if (remembered == Path.GetFileName(remembered) && remembered.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) && !remembered.StartsWith("~$", StringComparison.Ordinal) && File.Exists(Path.Combine(directory, remembered))) selectedWorkbook = remembered;
                }
            } catch { }
            store = new InventoryStore(directory, selectedWorkbook);
            Title = "电脑台账采集 · CBY"; FontFamily = new FontFamily("Microsoft YaHei UI"); FontSize = 13;
            var area = SystemParameters.WorkArea;
            Width = Math.Min(1220, area.Width - 30); Height = Math.Min(900, area.Height - 30);
            MinWidth = 840; MinHeight = 560; WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = Brush("#F5F5F7"); WindowStyle = WindowStyle.None; ResizeMode = ResizeMode.CanResize;
            UseLayoutRounding = true; SnapsToDevicePixels = true;
            WindowChrome.SetWindowChrome(this, new WindowChrome { CaptionHeight = 38, ResizeBorderThickness = new Thickness(6), CornerRadius = new CornerRadius(12), GlassFrameThickness = new Thickness(0), UseAeroCaptionButtons = false });
            using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("IosWindow.xaml")) view = (Grid)XamlReader.Load(input);
            Content = view; Icon = MakeIcon();
            Foreground = Brush("#1D1D1F");
            view.SizeChanged += delegate
            {
                bool compact = view.ActualWidth < 1080;
                Find<Grid>("Workspace").ColumnDefinitions[0].Width = new GridLength(compact ? 168 : 184);
                Find<Grid>("MainWorkspace").Margin = compact ? new Thickness(20, 12, 20, 12) : new Thickness(28, 14, 28, 14);
                bool stackCards = view.ActualWidth < 960;
                var cards = Find<Grid>("CollectionCards");
                cards.ColumnDefinitions[0].Width = new GridLength(stackCards ? 1 : 1.12, GridUnitType.Star);
                cards.ColumnDefinitions[1].Width = stackCards ? new GridLength(0) : new GridLength(1, GridUnitType.Star);
                Grid.SetColumn(Find<Border>("HardwareCard"), stackCards ? 0 : 1);
                Grid.SetRow(Find<Border>("HardwareCard"), stackCards ? 1 : 0);
                Find<Border>("ManualCard").Margin = stackCards ? new Thickness(0, 0, 0, 14) : new Thickness(0, 0, 14, 0);
                ApplyCompactHeight();
            };
            foreach (string n in new[] { "Minimize", "Maximize", "CloseWindow" }) WindowChrome.SetIsHitTestVisibleInChrome(Find<Button>(n), true);
            Find<Button>("Minimize").Click += delegate { WindowState = WindowState.Minimized; };
            Find<Button>("Maximize").Click += delegate { WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized; };
            Find<Button>("CloseWindow").Click += delegate { Close(); };
            StateChanged += delegate { view.Margin = WindowState == WindowState.Maximized ? new Thickness(7) : new Thickness(0); };
            Find<ToggleButton>("NavCollect").Click += delegate { SwitchPage(0, true); };
            Find<ToggleButton>("NavHistory").Click += delegate { SwitchPage(1, true); };
            Find<ToggleButton>("NavBackup").Click += delegate { SwitchPage(2, true); };
            Find<ToggleButton>("NavAbout").Click += delegate { SwitchPage(3, true); };
            Find<Image>("DeveloperAvatar").Source = LoadImage("developer-avatar.jpg");
            Find<Image>("AlipayQr").Source = LoadImage("donate-alipay.jpg");
            Find<Image>("WechatQr").Source = LoadImage("donate-wechat.jpg");
            Find<Button>("GitHubLink").Click += delegate { OpenContact(DeveloperGitHubUrl); };
            Find<Button>("EmailLink").Click += delegate { OpenContact(DeveloperEmailUri); };
            Find<Button>("AlipayPreview").Click += delegate { ShowDonation("支付宝打赏", Find<Image>("AlipayQr").Source); };
            Find<Button>("WechatPreview").Click += delegate { ShowDonation("微信打赏", Find<Image>("WechatQr").Source); };
            Find<ComboBox>("DeviceType").ItemsSource = new[] { "台式机", "笔记本", "一体机", "其他", "待核实" };
            Find<ComboBox>("DeviceState").ItemsSource = new[] { "在用", "闲置", "待报废", "待核实" };
            Find<ComboBox>("DeviceType").SelectedItem = "待核实"; Find<ComboBox>("DeviceState").SelectedItem = "待核实";
            Find<Button>("Save").Click += async delegate { await SaveComputerAsync(); };
            Find<Button>("Refresh").Click += async delegate { await ReadComputerAsync(); };
            Find<Button>("Summary").Click += async delegate { await OpenSummaryAsync(); };
            Find<Button>("HistorySummary").Click += async delegate { await OpenSummaryAsync(); };
            Find<Button>("Backup").Click += async delegate { await BackupAsync(); };
            Find<Button>("Restore").Click += async delegate { await RestoreAsync(); };
            Find<Button>("OpenBackups").Click += delegate { OpenFolder(store.BackupDirectory); };
            Find<TextBox>("HistoryFilter").TextChanged += delegate { RefreshHistory(); };
            Find<Button>("ClearHistoryFilter").Click += delegate { Find<TextBox>("HistoryFilter").Clear(); Find<TextBox>("HistoryFilter").Focus(); };
            Find<TextBox>("HistoryFilter").PreviewKeyDown += delegate(object sender, KeyEventArgs e) {
                if (e.Key == Key.Escape && Find<TextBox>("HistoryFilter").Text.Length > 0) { Find<TextBox>("HistoryFilter").Clear(); e.Handled = true; }
            };
            CommandBindings.Add(new CommandBinding(ApplicationCommands.Find, delegate {
                SwitchPage(1, false); Find<TextBox>("HistoryFilter").Focus(); Find<TextBox>("HistoryFilter").SelectAll();
            }, delegate(object sender, CanExecuteRoutedEventArgs e) { e.CanExecute = !busy; e.Handled = true; }));
            InputBindings.Add(new KeyBinding(ApplicationCommands.Find, Key.F, ModifierKeys.Control));
            Find<DataGrid>("History").SelectionChanged += delegate { UpdateHistoryActions(); };
            Find<DataGrid>("History").PreviewKeyDown += delegate(object sender, KeyEventArgs e) {
                var row = Find<DataGrid>("History").SelectedItem as HistoryLine;
                if (e.Key == Key.Enter && !busy && row != null) { e.Handled = true; ShowRecord(row.Record); }
            };
            Find<DataGrid>("History").MouseDoubleClick += delegate(object sender, MouseButtonEventArgs e) {
                if (busy) return;
                var origin = e.OriginalSource as DependencyObject;
                var row = origin == null ? null : ItemsControl.ContainerFromElement(Find<DataGrid>("History"), origin) as DataGridRow;
                if (row != null && row.Item is HistoryLine) ShowRecord(((HistoryLine)row.Item).Record);
            };
            Find<Button>("PreviewRecord").Click += delegate { var row = Find<DataGrid>("History").SelectedItem as HistoryLine; if (row != null) ShowRecord(row.Record); };
            Find<Button>("DeleteRecord").Click += async delegate { await DeleteSelectedRecordAsync(true); };
            Find<Button>("RefreshHistory").Click += async delegate { await RefreshRecordsAsync(); };
            Find<Button>("StatusDetails").Click += delegate { ShowDetails("操作详情", fullStatus); };
            Find<ComboBox>("WorkbookPicker").SelectionChanged += async delegate {
                if (changingWorkbook) return;
                var file = Find<ComboBox>("WorkbookPicker").SelectedItem as string;
                if (!String.IsNullOrEmpty(file) && file != selectedWorkbook) await SelectWorkbookAsync(file);
            };
            Find<Button>("RefreshWorkbooks").Click += delegate { RefreshWorkbookList(); };
            Find<ComboBox>("WorkbookPicker").DropDownOpened += delegate { if (!busy) RefreshWorkbookList(); };
            RefreshWorkbookList();
            SetStatus("打开后自动读取，补充现场信息，再点击保存。", "idle"); SetBusy(false);
            historyTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
            historyTimer.Tick += async delegate { await RefreshRecordsAsync(); };
            Loaded += async delegate { if (!preview) { await ReadComputerAsync(); historyTimer.Start(); } };
            Activated += async delegate { if (!preview) await RefreshRecordsAsync(); };
            Closing += delegate(object sender, System.ComponentModel.CancelEventArgs e) { if (busy || syncing) { e.Cancel = true; SetStatus("正在处理，请等待完成后再关闭。", "idle"); } };
            Closed += delegate { closed = true; historyTimer.Stop(); Find<System.Windows.Shapes.Path>("Spinner").RenderTransform.BeginAnimation(RotateTransform.AngleProperty, null); };
        }

        public T Find<T>(string name) where T : FrameworkElement { return (T)view.FindName(name); }
        private static SolidColorBrush Brush(string color) { return (SolidColorBrush)new BrushConverter().ConvertFromString(color); }
        private static BitmapImage LoadImage(string resource)
        {
            using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("Assets." + resource))
            {
                var image = new BitmapImage(); image.BeginInit(); image.CacheOption = BitmapCacheOption.OnLoad;
                image.StreamSource = input; image.EndInit(); image.Freeze(); return image;
            }
        }
        private void OpenContact(string address)
        {
            try { Process.Start(new ProcessStartInfo(address) { UseShellExecute = true }); }
            catch { ShowDetails("联系开发者", "暂时无法打开默认浏览器或邮件应用，请复制以下地址：\n\n" + address); }
        }
        private void ShowDonation(string title, ImageSource source)
        {
            var area = SystemParameters.WorkArea;
            var dialog = new Window { Title = title, Owner = this, Width = Math.Min(580, area.Width - 40), Height = Math.Min(830, area.Height - 40), MinWidth = 320, MinHeight = 400, WindowStartupLocation = WindowStartupLocation.CenterOwner, Background = Brushes.White, FontFamily = FontFamily, ShowInTaskbar = false };
            var panel = new DockPanel { Margin = new Thickness(18) };
            var close = new Button { Content = "关闭", Width = 116, Margin = new Thickness(0, 12, 0, 0), IsCancel = true, IsDefault = true };
            close.Click += delegate { dialog.Close(); }; DockPanel.SetDock(close, Dock.Bottom); panel.Children.Add(close);
            panel.Children.Add(new Image { Source = source, Stretch = Stretch.Uniform }); dialog.Content = panel; dialog.ShowDialog();
        }
        private static ImageSource MakeIcon()
        {
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                dc.DrawRoundedRectangle(Brush("#007AFF"), null, new Rect(0, 0, 64, 64), 16, 16);
                var pen = new Pen(Brushes.White, 2.6);
                dc.DrawRoundedRectangle(null, pen, new Rect(13, 16, 38, 27), 3, 3);
                dc.DrawLine(pen, new Point(32, 43), new Point(32, 52)); dc.DrawLine(pen, new Point(23, 52), new Point(41, 52));
            }
            var bitmap = new RenderTargetBitmap(64, 64, 96, 96, PixelFormats.Pbgra32); bitmap.Render(visual); bitmap.Freeze(); return bitmap;
        }

        public void SwitchPage(int index, bool animate)
        {
            if (index < 0 || index > 3) return; activePage = index;
            string[] pageNames = { "CollectPage", "HistoryPage", "BackupPage", "AboutPage" }, navNames = { "NavCollect", "NavHistory", "NavBackup", "NavAbout" };
            for (int i = 0; i < pageNames.Length; i++)
            {
                var page = Find<FrameworkElement>(pageNames[i]); page.BeginAnimation(OpacityProperty, null); page.Opacity = 1;
                page.RenderTransform = Transform.Identity; page.Visibility = i == index ? Visibility.Visible : Visibility.Collapsed;
                Find<ToggleButton>(navNames[i]).IsChecked = i == index;
            }
            Find<TextBlock>("PageTitle").Text = new[] { "设备采集", "历史记录", "备份与恢复", "关于" }[index];
            Find<TextBlock>("PageDescription").Text = new[] { "读取电脑信息，补充现场情况，然后保存。", "查看现有记录，选中后可预览或删除。", "自动保留每次备份，重要资料另存一份。", "认识开发者，交流反馈，支持工具继续完善。" }[index];
            Find<Grid>("WorkbookBar").Visibility = index == 3 ? Visibility.Collapsed : Visibility.Visible;
            Find<Border>("StatusPanel").Visibility = index == 3 ? Visibility.Collapsed : Visibility.Visible;
            if (index == 1) RefreshHistory(true);
            ApplyCompactHeight();
            if (animate && SystemParameters.ClientAreaAnimation)
            {
                var page = Find<FrameworkElement>(pageNames[index]); var move = new TranslateTransform(0, 9); page.RenderTransform = move;
                move.BeginAnimation(TranslateTransform.YProperty, new DoubleAnimation(9, 0, TimeSpan.FromMilliseconds(220)) { EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut } });
                page.BeginAnimation(OpacityProperty, new DoubleAnimation(.45, 1, TimeSpan.FromMilliseconds(180)));
            }
        }

        private void ApplyCompactHeight()
        {
            bool compact = view.ActualHeight > 0 && view.ActualHeight < 680;
            Find<Grid>("GlobalStats").Visibility = activePage == 3 || (compact && activePage != 0) ? Visibility.Collapsed : Visibility.Visible;
            Find<TextBlock>("PageDescription").Visibility = compact ? Visibility.Collapsed : Visibility.Visible;
            Find<Border>("HistoryPage").Padding = compact ? new Thickness(14) : new Thickness(20);
        }

        private void SetBusy(bool value)
        {
            busy = value;
            foreach (var n in new[] { "Location", "Person", "Number", "Remarks", "DeviceType", "DeviceState", "Acquired", "Refresh", "Summary", "Backup", "Restore", "WorkbookPicker", "RefreshWorkbooks" }) Find<FrameworkElement>(n).IsEnabled = !value;
            Find<Button>("Save").IsEnabled = !value && current != null;
            Find<TextBlock>("SaveText").Text = value ? "请稍候…" : "保存并追加一条";
            UpdateHistoryActions();
        }

        private void SetStatus(string text, string kind)
        {
            fullStatus = text;
            Find<TextBlock>("Status").Text = text.Length > 88 ? text.Substring(0, 85) + "…" : text;
            Find<Button>("StatusDetails").Visibility = kind == "warning" || text.Length > 88 ? Visibility.Visible : Visibility.Collapsed;
            Find<Border>("StatusPanel").Background = Brush(kind == "warning" ? "#FFF4E3" : kind == "success" ? "#EAF6EE" : "#EAF3FF");
            Find<TextBlock>("Status").Foreground = Brush(kind == "warning" ? "#946116" : kind == "success" ? "#24773A" : "#245C98");
            var spinner = Find<System.Windows.Shapes.Path>("Spinner"); var spin = (RotateTransform)spinner.RenderTransform;
            spin.BeginAnimation(RotateTransform.AngleProperty, null);
            spinner.Visibility = kind == "working" ? Visibility.Visible : Visibility.Collapsed;
            Find<System.Windows.Shapes.Ellipse>("StatusDot").Visibility = kind == "working" ? Visibility.Collapsed : Visibility.Visible;
            Find<System.Windows.Shapes.Ellipse>("StatusDot").Fill = Brush(kind == "warning" ? "#D89024" : kind == "success" ? "#34A853" : "#007AFF");
            if (kind == "working" && SystemParameters.ClientAreaAnimation) spin.BeginAnimation(RotateTransform.AngleProperty, new DoubleAnimation(0, 360, TimeSpan.FromMilliseconds(850)) { RepeatBehavior = RepeatBehavior.Forever });
        }

        private async Task ReadComputerAsync()
        {
            if (busy) return; SetBusy(true); SetStatus("正在读取本机硬件，请稍候…", "working");
            try
            {
                await syncTask;
                var sync = await Task.Run(() => store.SynchronizeSummary());
                var loaded = sync.Records; string note = sync.SummaryReadable ? "" : " " + sync.Message;
                SetHistorySync(sync.Message, !sync.SummaryReadable);
                var read = await Task.Run(() => HardwareCollector.Collect());
                var previous = loaded.LastOrDefault(r => !String.IsNullOrEmpty(read.DeviceKey) && r.DeviceKey == read.DeviceKey);
                bool keepInput = current != null; current = read; records = loaded;
                if (!keepInput) { if (previous != null) FillManual(previous); else { Find<ComboBox>("DeviceType").SelectedItem = read.DeviceType ?? "待核实"; Find<ComboBox>("DeviceState").SelectedItem = "待核实"; } }
                ShowHardware(); RefreshHistory(); UpdateCount();
                SetStatus((previous == null ? "读取完成。补充现场信息，即可保存。" : "已带出上次填写的信息。本次保存会追加新记录。") + note, note.Length > 0 ? "warning" : "success");
            }
            catch (Exception e) { SetStatus("读取未完成：" + e.Message, "warning"); }
            finally { SetBusy(false); }
        }

        private void FillManual(InventoryRecord r)
        {
            Find<TextBox>("Location").Text = InventoryPresentation.LocationText(r); Find<TextBox>("Person").Text = r.AssignedTo ?? "";
            Find<TextBox>("Number").Text = r.DeviceNumber ?? ""; Find<TextBox>("Remarks").Text = r.Remarks ?? ""; Find<IosDateField>("Acquired").DateValue = r.AcquiredDate ?? "";
            Find<ComboBox>("DeviceType").SelectedItem = r.DeviceType ?? "待核实"; Find<ComboBox>("DeviceState").SelectedItem = r.DeviceStatus ?? "待核实";
        }
        private void ShowHardware()
        {
            Find<TextBlock>("ModelTitle").Text = current.ManufacturerModel ?? "未读到型号，待核实";
            Find<TextBlock>("ModelHint").Text = (current.DeviceType ?? "待核实") + " · 数量 1";
            string[][] pairs = { new[] { "整机序列号", InventoryPresentation.SerialText(current.BiosSerial) }, new[] { "硬盘", InventoryPresentation.DiskText(current) }, new[] { "物理MAC", InventoryPresentation.MacText(current) }, new[] { "CPU", current.Cpu }, new[] { "内存", current.Memory }, new[] { "操作系统", current.OperatingSystem }, new[] { "待核对", InventoryPresentation.VerificationText(current) } };
            Find<ItemsControl>("HardwareRows").ItemsSource = pairs.Select(p => new HardwareLine { Label = p[0], Value = String.IsNullOrWhiteSpace(p[1]) ? "待核对" : p[1] }).ToList();
        }

        public async Task SaveComputerAsync()
        {
            if (busy || current == null) return;
            foreach (var n in new[] { "Location", "Person" }) if (String.IsNullOrWhiteSpace(Find<TextBox>(n).Text))
            { SwitchPage(0, false); Find<TextBox>(n).BringIntoView(); Find<TextBox>(n).Focus(); SetStatus("请填写位置和实际使用人。不清楚填“待核实”，公用电脑填“公用”。", "warning"); return; }
            current.Site = ""; current.Location = Find<TextBox>("Location").Text.Trim(); current.AssignedTo = Find<TextBox>("Person").Text.Trim();
            current.DeviceNumber = Find<TextBox>("Number").Text.Trim(); current.Remarks = Find<TextBox>("Remarks").Text.Trim(); current.AcquiredDate = Find<IosDateField>("Acquired").DateValue;
            current.DeviceType = Convert.ToString(Find<ComboBox>("DeviceType").SelectedItem); current.DeviceStatus = Convert.ToString(Find<ComboBox>("DeviceState").SelectedItem);
            SetBusy(true); SetStatus("正在追加记录、更新总表并生成备份，请不要拔出 U 盘…", "working");
            try
            {
                await syncTask;
                var result = await Task.Run(() => store.Save(current));
                Find<TextBlock>("RecordCount").Text = result.Count + " 条";
                Find<TextBlock>("BackupState").Text = result.BackupUpdated ? "本次备份已完成" : "本次备份待完成";
                Find<TextBlock>("BackupState").Foreground = Brush(result.BackupUpdated ? "#248A3D" : "#946116");
                SetStatus("已追加第 " + result.Count + " 条，旧记录保留。" + (result.SummaryUpdated ? "总表已更新；" : "总表待更新；") + (result.BackupUpdated ? "完整备份已保存。" : "请另存备份。"), result.SummaryUpdated && result.BackupUpdated ? "success" : "warning");
                fullStatus = result.Message;
                try { records = await Task.Run(() => store.LoadRecords()); RefreshHistory(); } catch { SetStatus(result.Message + " 历史列表暂未刷新，请重新打开程序。", "warning"); }
            }
            catch (Exception e) { SetStatus("保存未完成，请保留窗口并检查磁盘。" + e.Message, "warning"); }
            finally { SetBusy(false); }
        }

        private void UpdateCount() { Find<TextBlock>("RecordCount").Text = records.Count + " 条"; }

        private void RefreshWorkbookList()
        {
            changingWorkbook = true;
            try
            {
                var files = Directory.Exists(directory) ? Directory.GetFiles(directory, "*.xlsx", SearchOption.TopDirectoryOnly).Select(Path.GetFileName).Where(n => !n.StartsWith("~$", StringComparison.Ordinal)).OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase).ToList() : new List<string>();
                if (!files.Contains(selectedWorkbook, StringComparer.OrdinalIgnoreCase)) files.Insert(0, selectedWorkbook);
                Find<ComboBox>("WorkbookPicker").ItemsSource = files;
                Find<ComboBox>("WorkbookPicker").SelectedItem = files.First(n => String.Equals(n, selectedWorkbook, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception e) { SetStatus("表格列表暂未刷新：" + e.Message, "warning"); }
            finally { changingWorkbook = false; }
        }

        public async Task SelectWorkbookAsync(string fileName)
        {
            if (busy || String.IsNullOrWhiteSpace(fileName) || String.Equals(fileName, selectedWorkbook, StringComparison.OrdinalIgnoreCase)) return;
            SetBusy(true); SetStatus("正在读取所选台账…", "working");
            try
            {
                await syncTask;
                var nextStore = new InventoryStore(directory, fileName);
                var result = await Task.Run(() => nextStore.SynchronizeSummary());
                // An unreadable selection must never silently become a writable, empty ledger.
                if (!result.SummaryReadable && result.Records.Count == 0) throw new IOException(result.Message);
                store = nextStore; selectedWorkbook = fileName; records = result.Records;
                try {
                    var preferences = Path.Combine(directory, "数据"); Directory.CreateDirectory(preferences);
                    string destination = Path.Combine(preferences, "当前台账.txt"); string temporary = destination + "." + Guid.NewGuid().ToString("N") + ".tmp";
                    File.WriteAllText(temporary, selectedWorkbook, new System.Text.UTF8Encoding(false));
                    if (File.Exists(destination)) File.Replace(temporary, destination, null); else File.Move(temporary, destination);
                } catch { /* Selection remains usable when the preference file is read-only. */ }
                Find<TextBox>("HistoryFilter").Clear(); RefreshHistory(); UpdateCount();
                SetHistorySync(result.Message, !result.SummaryReadable);
                Find<TextBlock>("BackupState").Text = "当前台账单独备份";
                SetStatus("已切换到“" + fileName + "”，当前共 " + records.Count + " 条。之后保存到这份台账。", "success");
                SwitchPage(1, true);
            }
            catch (Exception e) { SetStatus("台账未切换：" + e.Message, "warning"); }
            finally { RefreshWorkbookList(); SetBusy(false); }
        }
        private void RefreshHistory(bool preserveScroll = false)
        {
            var grid = Find<DataGrid>("History"); var selected = grid.SelectedItem as HistoryLine;
            var selectedId = selected == null ? null : selected.Record.StorageId;
            var previousView = grid.ItemsSource == null ? null : CollectionViewSource.GetDefaultView(grid.ItemsSource);
            var sorting = previousView == null ? new List<SortDescription>() : previousView.SortDescriptions.ToList();
            var directions = grid.Columns.Select(c => c.SortDirection).ToList();
            var scroll = FindVisualChild<ScrollViewer>(grid);
            double vertical = preserveScroll && scroll != null ? scroll.VerticalOffset : 0;
            double horizontal = scroll == null ? 0 : scroll.HorizontalOffset;
            var query = Find<TextBox>("HistoryFilter").Text.Trim(); var rows = new List<HistoryLine>();
            Find<Button>("ClearHistoryFilter").Visibility = Find<TextBox>("HistoryFilter").Text.Length == 0 ? Visibility.Collapsed : Visibility.Visible;
            for (int i = 0; i < records.Count; i++)
            {
                var r = records[i]; var text = r.DeviceNumber + " " + r.BiosSerial + " " + r.Site + " " + r.Location + " " + r.AssignedTo + " " + r.ManufacturerModel;
                if (query.Length > 0 && text.IndexOf(query, StringComparison.OrdinalIgnoreCase) < 0) continue;
                DateTimeOffset time; rows.Add(new HistoryLine { Index = i + 1, Number = r.DeviceNumber, Serial = InventoryPresentation.SerialText(r.BiosSerial), Site = r.Site, Location = InventoryPresentation.LocationText(r), Person = r.AssignedTo, Model = r.ManufacturerModel, Time = DateTimeOffset.TryParse(r.CollectedAt, out time) ? time.ToString("yyyy-MM-dd HH:mm") : r.CollectedAt, Record = r });
            }
            grid.ItemsSource = rows;
            if (sorting.Count > 0)
            {
                var sortedView = CollectionViewSource.GetDefaultView(rows);
                using (sortedView.DeferRefresh()) foreach (var item in sorting) sortedView.SortDescriptions.Add(item);
                for (int i = 0; i < grid.Columns.Count; i++) grid.Columns[i].SortDirection = directions[i];
            }
            if (!String.IsNullOrEmpty(selectedId)) grid.SelectedItem = rows.FirstOrDefault(r => r.Record.StorageId == selectedId);
            if (scroll != null) { scroll.ScrollToVerticalOffset(vertical); scroll.ScrollToHorizontalOffset(horizontal); }
            Find<TextBlock>("HistoryCount").Text = query.Length == 0 ? "当前共有 " + records.Count + " 条记录" : "找到 " + rows.Count + " 条 · 当前共 " + records.Count + " 条";
            Find<FrameworkElement>("HistoryEmpty").Visibility = rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            Find<TextBlock>("HistoryEmptyText").Text = query.Length == 0 ? "还没有记录，采集保存后会显示在这里。" : "没有找到匹配的记录。";
            UpdateHistoryActions();
        }

        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i); var result = child as T ?? FindVisualChild<T>(child);
                if (result != null) return result;
            }
            return null;
        }

        private void UpdateHistoryActions()
        {
            bool selected = Find<DataGrid>("History").SelectedItem is HistoryLine;
            Find<Button>("PreviewRecord").IsEnabled = selected && !busy;
            Find<Button>("DeleteRecord").IsEnabled = selected && !busy;
            Find<Button>("RefreshHistory").IsEnabled = !busy;
            Find<Button>("HistorySummary").IsEnabled = !busy;
        }

        private void SetHistorySync(string message, bool warning)
        {
            Find<TextBlock>("HistorySync").Text = String.IsNullOrWhiteSpace(message) ? "已同步 · 表格删行并保存后会自动更新" : message;
            Find<TextBlock>("HistorySync").Foreground = Brush(warning ? "#946116" : "#86868B");
            Find<TextBlock>("RecordCount").ToolTip = "当前台账：" + selectedWorkbook + "\n" + Find<TextBlock>("HistorySync").Text;
        }

        public Task RefreshRecordsAsync()
        {
            if (closed || busy || syncing) return syncTask;
            syncTask = SynchronizeRecordsCoreAsync(); return syncTask;
        }

        private async Task SynchronizeRecordsCoreAsync()
        {
            syncing = true; UpdateHistoryActions();
            try
            {
                var result = await Task.Run(() => store.SynchronizeSummary());
                if (closed) return;
                bool different = records.Count != result.Records.Count || !records.Select(r => r.StorageId).SequenceEqual(result.Records.Select(r => r.StorageId));
                records = result.Records; UpdateCount();
                if (different) RefreshHistory(true);
                SetHistorySync(result.Message, !result.SummaryReadable);
                if (result.Changed) {
                    Find<TextBlock>("BackupState").Text = "删除前已自动备份";
                    if (!busy) SetStatus("记录已同步，当前共 " + records.Count + " 条。删除前已保留备份。", "success");
                }
            }
            catch (Exception e) { if (!closed) SetHistorySync("暂未同步：" + e.Message, true); }
            finally { syncing = false; if (!closed) UpdateHistoryActions(); }
        }

        public async Task DeleteSelectedRecordAsync(bool confirm)
        {
            if (busy) return;
            var row = Find<DataGrid>("History").SelectedItem as HistoryLine; if (row == null) return;
            if (confirm && !ConfirmDelete(row)) return;
            SetBusy(true); SetStatus("正在备份并删除所选记录…", "working");
            try
            {
                await syncTask;
                var result = await Task.Run(() => store.DeleteRecords(new[] { row.Record.StorageId }));
                records = await Task.Run(() => store.LoadRecords()); RefreshHistory(true); UpdateCount();
                Find<TextBlock>("BackupState").Text = result.BackupUpdated ? "删除前已自动备份" : "请检查备份";
                SetStatus(result.SummaryUpdated ? "已删除所选记录，当前共 " + records.Count + " 条。需要恢复时可选择删除前的备份。" : "记录已删除并备份。请关闭 Excel / WPS 后更新总表。", result.SummaryUpdated ? "success" : "warning");
                fullStatus = result.Message;
            }
            catch (Exception e) { SetStatus("删除未完成：" + e.Message, "warning"); }
            finally { SetBusy(false); }
        }

        private bool ConfirmDelete(HistoryLine row)
        {
            var dialog = new Window { Title = "删除记录", Owner = this, Width = 430, SizeToContent = SizeToContent.Height, ResizeMode = ResizeMode.NoResize, WindowStartupLocation = WindowStartupLocation.CenterOwner, Background = Brush("#F5F5F7"), FontFamily = FontFamily, ShowInTaskbar = false };
            var panel = new StackPanel { Margin = new Thickness(26) };
            panel.Children.Add(new TextBlock { Text = "删除这条记录？", FontSize = 21, FontWeight = FontWeights.SemiBold });
            panel.Children.Add(new TextBlock { Text = (String.IsNullOrWhiteSpace(row.Number) ? row.Model : row.Number + " · " + row.Model) + "\n" + row.Location + " · " + row.Person, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 15, 0, 12), Foreground = Brush("#6E6E73") });
            panel.Children.Add(new TextBlock { Text = "删除前会自动备份，需要时可以恢复。", TextWrapping = TextWrapping.Wrap, FontSize = 12, Foreground = Brush("#86868B") });
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 23, 0, 0) };
            var cancel = new Button { Content = "保留", Width = 100, IsCancel = true, IsDefault = true, Margin = new Thickness(0, 0, 10, 0) };
            var remove = new Button { Content = "确认删除", Width = 110, Foreground = Brush("#D92D20"), Background = Brush("#FFF0EF") };
            cancel.Click += delegate { dialog.DialogResult = false; }; remove.Click += delegate { dialog.DialogResult = true; };
            buttons.Children.Add(cancel); buttons.Children.Add(remove); panel.Children.Add(buttons); dialog.Content = panel;
            return dialog.ShowDialog() == true;
        }

        private async Task OpenSummaryAsync()
        {
            if (busy) return; SetBusy(true); SetStatus("正在更新当前记录总表…", "working");
            try { await syncTask; string path = await Task.Run(() => store.RebuildSummary()); records = await Task.Run(() => store.LoadRecords()); UpdateCount(); RefreshHistory(); SetStatus("总表已打开。删除整行并保存后，记录数会自动同步。", "success"); Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
            catch (Exception e) { SetStatus("总表暂时无法更新或打开。已保存的原始记录仍保留，请关闭 Excel 后重试。" + e.Message, "warning"); }
            finally { SetBusy(false); }
        }

        private async Task BackupAsync()
        {
            if (busy) return;
            using (var picker = new System.Windows.Forms.FolderBrowserDialog { Description = "选择备份位置，建议选单位电脑或另一块盘。", ShowNewFolderButton = true })
            {
                if (picker.ShowDialog(new WindowOwner(this)) != System.Windows.Forms.DialogResult.OK) return;
                SetBusy(true); SetStatus("正在创建完整备份…", "working");
                try { await syncTask; string path = await Task.Run(() => store.BackupNow(picker.SelectedPath)); Find<TextBlock>("BackupState").Text = "手动备份已完成"; Find<TextBlock>("BackupState").Foreground = Brush("#248A3D"); SetStatus("备份成功：" + path, "success"); }
                catch (Exception e) { SetStatus("备份未完成，原始记录没有改动。" + e.Message, "warning"); }
                finally { SetBusy(false); }
            }
        }
        private async Task RestoreAsync()
        {
            if (busy) return; var picker = new OpenFileDialog { Title = "选择本工具生成的备份", Filter = "台账备份 (*.zip)|*.zip", InitialDirectory = Directory.Exists(store.BackupDirectory) ? store.BackupDirectory : directory };
            if (picker.ShowDialog(this) != true) return;
            using (var target = new System.Windows.Forms.FolderBrowserDialog { Description = "选择恢复位置，将新建文件夹并保留现有记录。", ShowNewFolderButton = true })
            {
                if (target.ShowDialog(new WindowOwner(this)) != System.Windows.Forms.DialogResult.OK) return;
                SetBusy(true); SetStatus("正在校验备份并恢复到新文件夹…", "working");
                try
                {
                    await syncTask; string restored = await Task.Run(() => store.RestoreBackupToNewFolder(picker.FileName, target.SelectedPath));
                    try { File.Copy(Assembly.GetExecutingAssembly().Location, Path.Combine(restored, "电脑台账采集.exe"), false); }
                    catch { SetStatus("数据已恢复至 " + restored + "，请把本程序复制进去后继续使用。", "warning"); OpenFolder(restored); return; }
                    SetStatus("恢复成功，请从新文件夹打开程序继续使用：" + restored, "success"); OpenFolder(restored);
                }
                catch (Exception e) { SetStatus("恢复未完成，现有记录没有改动。" + e.Message, "warning"); }
                finally { SetBusy(false); }
            }
        }
        private void OpenFolder(string path) { try { Directory.CreateDirectory(path); Process.Start("explorer.exe", "\"" + path + "\""); } catch (Exception e) { SetStatus("暂时无法打开文件夹。" + e.Message, "warning"); } }
        private void ShowRecord(InventoryRecord r)
        {
            ShowDetails("采集记录详情", "资产编号：" + r.DeviceNumber + "\n品牌型号：" + r.ManufacturerModel + "\n整机序列号：" + InventoryPresentation.SerialText(r.BiosSerial) + "\n\n硬盘：\n" + InventoryPresentation.DiskText(r) + "\n\n物理MAC：\n" + InventoryPresentation.MacText(r) + "\n\nCPU：" + r.Cpu + "\n内存：" + r.Memory + "\n系统：" + r.OperatingSystem + "\n\n位置：" + InventoryPresentation.LocationText(r) + "\n使用人：" + r.AssignedTo + "\n类型 / 状态：" + r.DeviceType + " / " + r.DeviceStatus + "\n取得日期：" + r.AcquiredDate + "\n备注：" + r.Remarks + "\n采集时间：" + InventoryPresentation.CollectedTime(r.CollectedAt) + "\n待核对：" + InventoryPresentation.VerificationText(r));
        }
        private void ShowDetails(string title, string text)
        {
            var dialog = new Window { Title = title, Owner = this, Width = 690, Height = 570, MinWidth = 440, MinHeight = 340, WindowStartupLocation = WindowStartupLocation.CenterOwner, Background = Brush("#F5F5F7"), FontFamily = FontFamily };
            var panel = new DockPanel { Margin = new Thickness(24) };
            var heading = new TextBlock { Text = title, FontSize = 22, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 20) }; DockPanel.SetDock(heading, Dock.Top); panel.Children.Add(heading);
            var footer = new DockPanel { Margin = new Thickness(0, 18, 0, 0), LastChildFill = false };
            var close = new Button { Content = "关闭", Style = (Style)FindResource("PrimaryButton"), Width = 116, IsCancel = true, IsDefault = true }; close.Click += delegate { dialog.Close(); }; DockPanel.SetDock(close, Dock.Right); footer.Children.Add(close);
            DockPanel.SetDock(footer, Dock.Bottom); panel.Children.Add(footer);
            panel.Children.Add(new TextBox { Text = text, IsReadOnly = true, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Height = Double.NaN, VerticalContentAlignment = VerticalAlignment.Top, Padding = new Thickness(16), AcceptsReturn = true }); dialog.Content = panel; dialog.ShowDialog();
        }
        public void SetPreview()
        {
            current = new InventoryRecord { HardwareDataVersion = 3, DeviceKey = "DEMO-IOS", CollectedAt = "2026-10-12T09:30:00+08:00", Site = "示例管理站", Location = "综合办公室", AssignedTo = "公用", DeviceNumber = "PC-001", DeviceType = "台式机", DeviceStatus = "在用", AcquiredDate = "2023-08-15", ManufacturerModel = "示例办公电脑", Cpu = "示例处理器 · 6 核", Memory = "16 GiB", OperatingSystem = "Windows 11 · 64 位", BiosSerial = "DEMO-001", DiskSerials = "硬盘0 — 示例固态硬盘 — DEMO-DISK-001", MacAddresses = "02:00:00:00:00:01", IPAddresses = "", AdapterPairs = "", Verification = "无", Remarks = "" };
            FillManual(current); ShowHardware(); SetBusy(false); SetStatus("读取完成。核对现场信息，保存后会追加一条记录。", "success");
        }
        private sealed class WindowOwner : System.Windows.Forms.IWin32Window { private readonly IntPtr handle; public WindowOwner(Window w) { handle = new System.Windows.Interop.WindowInteropHelper(w).Handle; } public IntPtr Handle { get { return handle; } } }
    }
}
