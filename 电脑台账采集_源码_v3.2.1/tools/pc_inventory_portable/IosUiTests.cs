using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Xml;
using PcInventory;
using PcInventory.Ios;

// Synthetic data only. This harness targets only the HWND it creates itself.
// It never moves the global pointer, uses SendInput, opens Excel or reads WMI.
internal static class IosUiTests
{
    private static readonly List<string> report = new List<string>();
    private static string qaRoot;
    private static int checks, nativePointChecks, clickCount;
    private static bool? nativeMouseAvailable;
    private static Exception failure;
    private static InventoryWindow window;
    private static IntPtr ownHwnd;
    private static string ownData;
    private const int WM_MOUSEMOVE = 0x0200, WM_LBUTTONDOWN = 0x0201, WM_LBUTTONUP = 0x0202, WM_CHAR = 0x0102;

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private static void Log(string message)
    {
        report.Add(message); Console.WriteLine(message);
        if (!String.IsNullOrEmpty(qaRoot)) File.WriteAllLines(Path.Combine(qaRoot, "交互验证日志.txt"), report.ToArray(), System.Text.Encoding.UTF8);
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException("FAILED: " + message);
        checks++; Log("PASS: " + message);
    }

    private static async Task Layout()
    {
        window.UpdateLayout();
        await window.Dispatcher.InvokeAsync(delegate { window.UpdateLayout(); }, DispatcherPriority.Render);
        await window.Dispatcher.InvokeAsync(delegate { }, DispatcherPriority.ContextIdle);
    }

    private static async Task Until(Func<bool> predicate, string description)
    {
        var timer = Stopwatch.StartNew();
        while (!predicate())
        {
            if (timer.ElapsedMilliseconds > 15000) throw new TimeoutException(description);
            await Task.Delay(15);
        }
        await Layout();
    }

    private static T Find<T>(string name) where T : FrameworkElement
    {
        T element = window.Find<T>(name);
        if (element == null) throw new InvalidOperationException("Missing named element: " + name);
        return element;
    }

    private static DependencyObject Parent(DependencyObject node)
    {
        if (node == null) return null;
        if (node is Visual || node is System.Windows.Media.Media3D.Visual3D)
        {
            var parent = VisualTreeHelper.GetParent(node);
            if (parent != null) return parent;
        }
        var content = node as FrameworkContentElement;
        if (content != null) return content.Parent ?? content.TemplatedParent;
        var element = node as FrameworkElement;
        if (element != null) return element.Parent ?? element.TemplatedParent;
        return LogicalTreeHelper.GetParent(node);
    }

    private static bool BelongsTo(DependencyObject node, DependencyObject owner)
    {
        for (int i = 0; node != null && i < 80; i++, node = Parent(node))
            if (Object.ReferenceEquals(node, owner)) return true;
        return false;
    }

    private static string Chain(DependencyObject node)
    {
        var items = new List<string>();
        for (int i = 0; node != null && i < 15; i++, node = Parent(node))
        { var element = node as FrameworkElement; items.Add(node.GetType().Name + (element == null ? "" : "/" + element.Name)); }
        return String.Join(" > ", items.ToArray());
    }

    private static Point WindowPoint(FrameworkElement element, Point point)
    { return element.TranslatePoint(point, window); }

    private static IntPtr Position(Point point)
    {
        var source = HwndSource.FromHwnd(ownHwnd);
        if (source == null || source.CompositionTarget == null) throw new InvalidOperationException("Own HWND has no composition target.");
        Point pixels = source.CompositionTarget.TransformToDevice.Transform(point);
        int x = (int)Math.Round(pixels.X), y = (int)Math.Round(pixels.Y);
        if (x < 0 || y < 0 || x > 32767 || y > 32767) throw new InvalidOperationException("Input point outside own client coordinate range.");
        return new IntPtr((y << 16) | (x & 0xffff));
    }

    private static void OwnMouse(int message, FrameworkElement element, Point localPoint)
    {
        uint processId;
        GetWindowThreadProcessId(ownHwnd, out processId);
        if (processId != (uint)Process.GetCurrentProcess().Id || ownHwnd != new WindowInteropHelper(window).Handle)
            throw new InvalidOperationException("Refusing to send input outside the harness window.");
        SendMessage(ownHwnd, message, message == WM_LBUTTONDOWN ? new IntPtr(1) : IntPtr.Zero, Position(WindowPoint(element, localPoint)));
    }

    private static async Task Click(FrameworkElement element, Point localPoint)
    {
        OwnMouse(WM_MOUSEMOVE, element, localPoint);
        OwnMouse(WM_LBUTTONDOWN, element, localPoint);
        OwnMouse(WM_LBUTTONUP, element, localPoint);
        await Layout();
    }

    private static void TypeUnicode(string value)
    {
        foreach (char letter in value) SendMessage(ownHwnd, WM_CHAR, new IntPtr(letter), new IntPtr(1));
    }

    private static Point[] InputPoints(FrameworkElement element)
    {
        double x = element.ActualWidth, y = element.ActualHeight;
        double edge = Math.Min(8, Math.Min(x, y) / 4);
        return new[] { new Point(x / 2, y / 2), new Point(x / 2, edge), new Point(x / 2, y - edge),
            new Point(edge, y / 2), new Point(x - edge, y / 2), new Point(edge, edge),
            new Point(x - edge, edge), new Point(edge, y - edge), new Point(x - edge, y - edge) };
    }

    private static HitTestResult VisibleHit(Point point)
    {
        HitTestResult found = null;
        VisualTreeHelper.HitTest(window, delegate(DependencyObject candidate)
        {
            var ui = candidate as UIElement;
            if (ui != null && (!ui.IsVisible || !ui.IsHitTestVisible)) return HitTestFilterBehavior.ContinueSkipSelfAndChildren;
            return HitTestFilterBehavior.Continue;
        }, delegate(HitTestResult result) { found = result; return HitTestResultBehavior.Stop; }, new PointHitTestParameters(point));
        return found;
    }

    private static async Task InputHitChecks(string label, bool sendCharacters)
    {
        window.SwitchPage(0, false);
        foreach (string name in new[] { "Location", "Person", "Number", "Remarks" })
        {
            TextBox box = Find<TextBox>(name);
            box.BringIntoView(); await Layout();
            Require(box.IsVisible && box.IsEnabled && box.ActualWidth > 30 && box.ActualHeight > 20, label + " " + name + " has a usable visible field");
            int pointIndex = 0;
            foreach (Point point in InputPoints(box))
            {
                Point rootPoint = WindowPoint(box, point);
                HitTestResult visual = VisibleHit(rootPoint);
                IInputElement routed = window.InputHitTest(rootPoint);
                if (visual == null || !BelongsTo(visual.VisualHit, box) || !BelongsTo(routed as DependencyObject, box))
                {
                    var scroll = Find<ScrollViewer>("CollectionScroll");
                    var actualHit = routed as FrameworkElement;
                    Log("HIT FAILURE: field=" + name + "; point=" + rootPoint + "; actual=" + (actualHit == null ? (routed == null ? "null" : routed.GetType().Name) : actualHit.GetType().Name + "/" + actualHit.Name) + "; scroll offset=" + scroll.VerticalOffset + "; viewport=" + scroll.ViewportHeight + "; field size=" + box.ActualWidth + "x" + box.ActualHeight);
                    Log("VISUAL CHAIN: " + Chain(visual == null ? null : visual.VisualHit)); Log("INPUT CHAIN: " + Chain(routed as DependencyObject));
                    Screenshot("FAIL_输入命中.png");
                }
                Require(visual != null && BelongsTo(visual.VisualHit, box) && BelongsTo(routed as DependencyObject, box),
                    label + " " + name + " hit target " + (++pointIndex) + "/9 belongs to the TextBox, including visual padding");
                Keyboard.Focus(Find<Button>("Save"));
                await Click(box, point);
                bool nativeFocused = box.IsKeyboardFocusWithin && BelongsTo(Keyboard.FocusedElement as DependencyObject, box);
                if (!nativeMouseAvailable.HasValue)
                {
                    nativeMouseAvailable = nativeFocused;
                    if (!nativeFocused) Log("LIMITATION: own-HWND SendMessage did not assign focus to an offscreen inactive window. IsActive=" + window.IsActive + "; keyboard focus=" + (Keyboard.FocusedElement == null ? "null" : Keyboard.FocusedElement.GetType().FullName) + "; focusable=" + box.Focusable + ". Physical click and WM_CHAR typing are NOT claimed. Continuing with actual visual hit targets plus explicit UI Automation focus/value checks.");
                }
                if (nativeMouseAvailable.Value)
                {
                    Require(nativeFocused, label + " " + name + " own-HWND point click " + pointIndex + " focuses the editor");
                    nativePointChecks++;
                }
                else
                {
                    UIElementAutomationPeer.CreatePeerForElement(box).SetFocus();
                    await Layout();
                    Require(box.IsKeyboardFocusWithin, label + " " + name + " explicit AutomationPeer focus succeeds after hit target " + pointIndex + " (not a physical click claim)");
                }
                if (sendCharacters)
                {
                    box.Clear();
                    if (nativeMouseAvailable.Value) TypeUnicode("Q中文7");
                    else ((IValueProvider)UIElementAutomationPeer.CreatePeerForElement(box).GetPattern(PatternInterface.Value)).SetValue("Q中文7");
                    await Layout();
                    Require(box.Text == "Q中文7", label + " " + name + " point " + pointIndex + (nativeMouseAvailable.Value ? " receives WM_CHAR Unicode input" : " accepts Automation ValuePattern Unicode input"));
                }
            }
        }
    }

    private static void DateValue(string value)
    {
        var date = Find<FrameworkElement>("Acquired");
        PropertyInfo property = date.GetType().GetProperty("DateValue", BindingFlags.Instance | BindingFlags.Public);
        if (property == null) throw new InvalidOperationException("Acquired must expose public DateValue.");
        property.SetValue(date, value, null);
    }

    private static string DateValue()
    {
        var date = Find<FrameworkElement>("Acquired");
        return Convert.ToString(date.GetType().GetProperty("DateValue").GetValue(date, null));
    }

    private static int WorkbookRows(string path)
    {
        using (var file = File.OpenRead(path))
        using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
        using (var source = zip.GetEntry("xl/worksheets/sheet1.xml").Open())
        {
            var xml = new XmlDocument(); xml.Load(source);
            var ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            return xml.SelectNodes("//s:sheetData/s:row", ns).Count;
        }
    }

    private static void SetManual(string person)
    {
        Find<TextBox>("Location").Text = "TEST_ONLY示例站办公室";
        Find<TextBox>("Person").Text = person;
        Find<TextBox>("Number").Text = "0007";
        Find<TextBox>("Remarks").Text = "TEST_ONLY虚构交互验证";
    }

    private static IInvokeProvider InvokeProvider(Button button)
    {
        var peer = UIElementAutomationPeer.CreatePeerForElement(button);
        Require(peer != null, button.Name + " exposes an automation peer");
        var invoke = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
        Require(invoke != null, button.Name + " exposes InvokePattern");
        return invoke;
    }

    private static async Task KeyboardAndAutomation()
    {
        window.SwitchPage(0, false);
        Require(window.Find<TextBox>("Site") == null, "collection exposes one location editor without a separate site field");
        TextBox location = Find<TextBox>("Location"); location.BringIntoView(); await Layout();
        var peer = UIElementAutomationPeer.CreatePeerForElement(location);
        Require(peer != null && !String.IsNullOrWhiteSpace(peer.GetName()), "location editor has a meaningful accessibility name");
        var value = peer.GetPattern(PatternInterface.Value) as IValueProvider;
        Require(value != null && !value.IsReadOnly, "location editor exposes editable ValuePattern");
        value.SetValue("TEST_ONLY自动化值"); peer.SetFocus(); await Layout();
        Require(value.Value == "TEST_ONLY自动化值" && location.IsKeyboardFocusWithin, "ValuePattern and peer focus reach the native WPF editor");
        string[] expected = { "Location", "Person", "Number" };
        for (int i = 0; i < expected.Length - 1; i++)
        {
            var current = Find<TextBox>(expected[i]); current.Focus();
            current.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next)); await Layout();
            Require(Find<TextBox>(expected[i + 1]).IsKeyboardFocusWithin, "Tab traversal " + expected[i] + " -> " + expected[i + 1]);
        }
        var person = Find<TextBox>("Person"); person.Focus();
        person.MoveFocus(new TraversalRequest(FocusNavigationDirection.Previous)); await Layout();
        Require(location.IsKeyboardFocusWithin, "Shift+Tab traversal Person -> Location");
    }

    private static void Screenshot(string name)
    {
        FrameworkElement root = window.Content as FrameworkElement;
        if (root == null) throw new InvalidOperationException("Window content has no renderable root.");
        root.UpdateLayout();
        Rect extent = root.TransformToAncestor(window).TransformBounds(new Rect(0, 0, root.ActualWidth, root.ActualHeight));
        var bitmap = new RenderTargetBitmap(Math.Max(1, (int)Math.Ceiling(extent.Right)), Math.Max(1, (int)Math.Ceiling(extent.Bottom)), 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(root);
        var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using (var stream = File.Create(Path.Combine(qaRoot, name))) encoder.Save(stream);
        Log("SCREENSHOT: " + name);
    }

    private static async Task ButtonStateScreenshots()
    {
        window.SwitchPage(0, false); Find<ScrollViewer>("CollectionScroll").ScrollToTop(); await Layout();
        Button save = Find<Button>("Save");
        Point middle = new Point(save.ActualWidth / 2, save.ActualHeight / 2);
        Screenshot("01_采集界面_正常.png");
        OwnMouse(WM_MOUSEMOVE, save, middle); await Task.Delay(220); await Layout(); Screenshot("02_按钮悬停消息后.png");
        Log("VISUAL STATE OBSERVED: IsMouseOver=" + save.IsMouseOver + " (message injection is not a physical-pointer assertion).");
        OwnMouse(WM_LBUTTONDOWN, save, middle); await Task.Delay(100); await Layout(); Screenshot("03_按钮按下消息后.png");
        Log("VISUAL STATE OBSERVED: IsPressed=" + save.IsPressed + " (message injection is not a physical-pointer assertion).");
        // Release outside the button so screenshot testing does not save a record.
        OwnMouse(WM_MOUSEMOVE, window, new Point(4, 4)); OwnMouse(WM_LBUTTONUP, window, new Point(4, 4));
        await Layout();
        save.Focus(); await Layout(); Screenshot("04_按钮键盘焦点.png");
        save.IsEnabled = false; await Layout(); Screenshot("05_按钮禁用.png");
        save.IsEnabled = true;
        Require(new InventoryStore(ownData).LoadRecords().Count == 0, "press then release outside button does not save");
    }

    private static async Task SaveRegression()
    {
        var store = new InventoryStore(ownData);
        var button = Find<Button>("Save"); button.Click += delegate { clickCount++; };
        var provider = InvokeProvider(button);
        SetManual("TEST_ONLY第一次"); DateValue(""); Find<TextBox>("Location").Clear();
        await window.SaveComputerAsync();
        Require(store.LoadRecords().Count == 0 && !window.IsBusy, "blank location prevents persistence and clears busy state");
        Require(!String.IsNullOrWhiteSpace(Find<TextBlock>("Status").Text), "validation feedback is visible");
        SetManual("");
        await window.SaveComputerAsync();
        Require(store.LoadRecords().Count == 0 && !window.IsBusy, "blank person prevents persistence even with location filled");
        SetManual("TEST_ONLY第一次");
        provider.Invoke();
        await Until(() => clickCount == 1 && !window.IsBusy && store.LoadRecords().Count == 1, "first UI Automation save did not finish");
        Require(clickCount == 1 && store.LoadRecords().Count == 1, "one automation Invoke causes exactly one click and one saved row");
        Require(store.LoadRecords()[0].Site == "" && store.LoadRecords()[0].Location == "TEST_ONLY示例站办公室", "location alone saves without a site value");
        Require(store.LoadRecords()[0].AcquiredDate == "" && store.LoadRecords()[0].DeviceNumber == "0007", "unknown date and leading zero asset number persist");

        SetManual("TEST_ONLY第二次"); DateValue("2023-08");
        Task second = window.SaveComputerAsync();
        Require(window.IsBusy && !button.IsEnabled && !Find<TextBox>("Location").IsEnabled, "saving immediately disables editing and action");
        bool invokeBlocked = false;
        try { provider.Invoke(); }
        catch (ElementNotEnabledException) { invokeBlocked = true; }
        await window.SaveComputerAsync();
        await second; await Layout();
        Require(invokeBlocked && clickCount == 1 && store.LoadRecords().Count == 2, "busy automation Invoke and direct reentry cannot append again");
        Require(!window.IsBusy && button.IsEnabled && Find<TextBox>("Location").IsEnabled, "all editing and save controls reenable after save");

        SetManual("TEST_ONLY第三次"); DateValue("2023-08-15");
        if (nativeMouseAvailable == true) await Click(button, new Point(button.ActualWidth / 2, button.ActualHeight / 2));
        else provider.Invoke();
        await Until(() => clickCount == 2 && !window.IsBusy && store.LoadRecords().Count == 3, "own-HWND button click did not save");
        var saved = store.LoadRecords();
        Require(saved.Count == 3 && saved[0].AssignedTo == "TEST_ONLY第一次" && saved[1].AssignedTo == "TEST_ONLY第二次" && saved[2].AssignedTo == "TEST_ONLY第三次", "same device appends three rows without overwriting earlier values");
        Require(new HashSet<string>(saved.Select(record => record.RecordId)).Count == 3, "three UI saves have three distinct record IDs");
        Require(saved[1].AcquiredDate == "2023-08" && saved[2].AcquiredDate == "2023-08-15", "legacy month and full date precision remain intact");
        Require(Directory.GetFiles(store.BackupDirectory, "*.zip").Length == 3 && WorkbookRows(store.SummaryPath) == 4, "three UI saves produce three full backups and three workbook rows");

        using (var locked = File.Open(store.SummaryPath, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            SetManual("TEST_ONLY第四次"); await window.SaveComputerAsync();
            Require(store.LoadRecords().Count == 4 && Directory.GetFiles(store.BackupDirectory, "*.zip").Length == 4, "exclusive workbook lock preserves fourth raw record and full backup");
            string message = Find<TextBlock>("Status").Text;
            Require(message.Contains("总表") && (message.Contains("待") || message.Contains("未") || message.Contains("占用")), "workbook lock gives a visible workbook-pending status");
            Screenshot("06_总表占用时状态.png");
        }
        store.RebuildSummary();
        Require(WorkbookRows(store.SummaryPath) == 5 && !window.IsBusy, "workbook rebuild after release includes all four saved rows");
        window.SwitchPage(1, false); await Layout();
        Require(Find<DataGrid>("History").Items.Count == 4, "history lists all four same-device records");
        Find<TextBox>("HistoryFilter").Text = "第三次"; await Layout();
        Require(Find<DataGrid>("History").Items.Count == 1, "history filtering finds the requested record");
        Find<TextBox>("HistoryFilter").Clear(); await Layout(); Screenshot("07_历史记录.png");
        window.SwitchPage(2, false); await Layout();
        Require(Find<Button>("Backup").IsVisible && Find<Button>("Restore").IsVisible, "backup and restore actions are reachable on backup page");
        Screenshot("08_备份恢复.png");
    }

    private static bool IsInsideWindow(FrameworkElement element)
    {
        Rect bounds = element.TransformToAncestor(window).TransformBounds(new Rect(0, 0, element.ActualWidth, element.ActualHeight));
        return element.IsVisible && bounds.Width > 0 && bounds.Height > 0 && bounds.Left >= -1 && bounds.Top >= -1 && bounds.Right <= window.ActualWidth + 1 && bounds.Bottom <= window.ActualHeight + 1;
    }

    private static byte[] ImagePixels(BitmapSource bitmap)
    {
        var converted = new FormatConvertedBitmap(bitmap, PixelFormats.Bgra32, null, 0);
        int stride = checked(converted.PixelWidth * 4);
        var pixels = new byte[checked(stride * converted.PixelHeight)];
        converted.CopyPixels(pixels, stride, 0);
        return pixels;
    }

    private static void VerifyEmbeddedImages()
    {
        var embedded = new List<BitmapSource>();
        var assembly = typeof(InventoryWindow).Assembly;
        foreach (string name in assembly.GetManifestResourceNames())
        {
            if (!name.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) &&
                !name.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) &&
                !name.EndsWith(".png", StringComparison.OrdinalIgnoreCase)) continue;
            using (var stream = assembly.GetManifestResourceStream(name))
                embedded.Add(BitmapDecoder.Create(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad).Frames[0]);
        }
        Require(embedded.Count >= 3, "avatar and two donation images are embedded in the executable");
        foreach (string name in new[] { "DeveloperAvatar", "AlipayQr", "WechatQr" })
        {
            Image image = Find<Image>(name);
            var bitmap = image.Source as BitmapSource;
            Require(bitmap != null && bitmap.PixelWidth > 64 && bitmap.PixelHeight > 64 && image.ActualWidth > 20 && image.ActualHeight > 20,
                name + " decodes and has a nonempty rendered size");
            byte[] pixels = ImagePixels(bitmap);
            Require(embedded.Any(candidate => candidate.PixelWidth == bitmap.PixelWidth && candidate.PixelHeight == bitmap.PixelHeight && ImagePixels(candidate).SequenceEqual(pixels)),
                name + " displays the original embedded image pixels without a sidecar image file");
        }
    }

    private static IEnumerable<T> VisualDescendants<T>(DependencyObject root) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
        {
            DependencyObject child = VisualTreeHelper.GetChild(root, i);
            if (child is T) yield return (T)child;
            foreach (T descendant in VisualDescendants<T>(child)) yield return descendant;
        }
    }

    private static async Task DonationPreview()
    {
        foreach (var preview in new[] {
            new KeyValuePair<string, string>("AlipayPreview", "AlipayQr"),
            new KeyValuePair<string, string>("WechatPreview", "WechatQr") })
        {
            Require(window.OwnedWindows.Count == 0, preview.Key + " begins without leftover dialogs");
            InvokeProvider(Find<Button>(preview.Key)).Invoke();
            await Until(() => window.OwnedWindows.Count == 1, "donation preview dialog did not open");
            Window dialog = window.OwnedWindows[0];
            try
            {
                dialog.UpdateLayout();
                var expanded = VisualDescendants<Image>(dialog).Single();
                Require(Object.ReferenceEquals(expanded.Source, Find<Image>(preview.Value).Source), preview.Key + " opens the same original image in its own dialog");
                Require(expanded.ActualHeight > Find<Image>(preview.Value).ActualHeight && expanded.Stretch == Stretch.Uniform, preview.Key + " enlarges the image while preserving its aspect ratio");
                var close = VisualDescendants<Button>(dialog).Single(button => button.IsCancel);
                Require(close.IsEnabled, preview.Key + " offers an enabled cancel button for Escape dismissal");
                InvokeProvider(close).Invoke();
                await Until(() => window.OwnedWindows.Count == 0, "donation preview dialog did not close");
                Require(window.IsEnabled && Find<ScrollViewer>("AboutPage").IsVisible, preview.Key + " returns to the usable About page after closing");
            }
            finally { if (dialog.IsVisible) dialog.Close(); }
        }
    }

    private static async Task AboutAndNavigation()
    {
        string[] pages = { "CollectPage", "HistoryPage", "BackupPage", "AboutPage" };
        string[] navigation = { "NavCollect", "NavHistory", "NavBackup", "NavAbout" };
        string[] titles = { "设备采集", "历史记录", "备份与恢复", "关于" };
        window.Width = 1216; window.Height = 820;
        for (int index = 0; index < pages.Length; index++)
        {
            Find<ToggleButton>(navigation[index]).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            await Layout();
            Require(Find<TextBlock>("PageTitle").Text == titles[index], navigation[index] + " click selects the corresponding page title");
            for (int other = 0; other < pages.Length; other++)
            {
                Require((Find<FrameworkElement>(pages[other]).Visibility == Visibility.Visible) == (other == index), pages[other] + " visibility is exclusive after selecting " + titles[index]);
                Require(Find<ToggleButton>(navigation[other]).IsChecked == (other == index), navigation[other] + " checked state follows the selected page");
            }
        }
        window.SwitchPage(3, false); await Layout();
        var about = Find<ScrollViewer>("AboutPage");
        Require(BelongsTo(Find<TextBlock>("DeveloperMark"), about) && Find<TextBlock>("DeveloperMark").Text.Contains("CBY"), "developer attribution is inside the About page");
        foreach (var link in new[] {
            new KeyValuePair<string, string>("GitHubLink", "https://github.com/CBY-Aluuu"),
            new KeyValuePair<string, string>("EmailLink", "mailto:1140357618@qq.com") })
        {
            Button button = Find<Button>(link.Key);
            Require(button.IsEnabled && BelongsTo(button, about), link.Key + " is an enabled button in About");
            Require(String.Equals(Convert.ToString(button.Tag), link.Value, StringComparison.Ordinal), link.Key + " declares the exact requested external destination");
            Require(!String.IsNullOrWhiteSpace(UIElementAutomationPeer.CreatePeerForElement(button).GetName()), link.Key + " has a meaningful accessibility name");
            InvokeProvider(button);
        }
        Log("NOTE: external links are inspected but never invoked; no browser or mail application is opened.");
        VerifyEmbeddedImages();
        about.ScrollToTop(); await Layout(); Screenshot("13_关于页面.png");
        window.Width = 840; window.Height = 560; await Layout();
        Require(about.ScrollableHeight > 0 && about.VerticalScrollBarVisibility == ScrollBarVisibility.Auto, "840x560 About content has automatic vertical scrolling");
        about.ScrollToTop(); await Layout();
        Require(IsInsideWindow(Find<Image>("DeveloperAvatar")), "840x560 About avatar is visible at the top");
        Screenshot("14_关于小窗口顶部.png");
        about.ScrollToBottom(); await Layout();
        Require(about.VerticalOffset > 0 && Math.Abs(about.VerticalOffset - about.ScrollableHeight) < 1, "840x560 About can reach its bottom content");
        foreach (string name in new[] { "AlipayQr", "WechatQr" })
        {
            Image image = Find<Image>(name);
            Rect bounds = image.TransformToAncestor(about).TransformBounds(new Rect(0, 0, image.ActualWidth, image.ActualHeight));
            Require(bounds.Left >= -1 && bounds.Right <= about.ActualWidth + 1, "840x560 " + name + " fits horizontally without clipping");
        }
        Screenshot("15_关于小窗口底部.png");
        await DonationPreview();
        window.SwitchPage(0, false); await Layout();
        Require(!Find<TextBlock>("DeveloperMark").IsVisible, "developer attribution is hidden outside About");
        window.Width = 1216; window.Height = 820; await Layout();
    }

    private static async Task LayoutAndNavigation()
    {
        window.SwitchPage(0, false); window.Width = 1024; window.Height = 700; await Layout();
        Require(IsInsideWindow(Find<Button>("Save")) && IsInsideWindow(Find<TextBlock>("Status")), "1024x700 retains save action and status within window");
        await InputHitChecks("1024x700", false); Screenshot("09_小窗口.png");
        window.Width = 840; window.Height = 560; await Layout();
        Require(IsInsideWindow(Find<Button>("Save")) && IsInsideWindow(Find<TextBlock>("Status")), "840x560 retains save action and status within window");
        await InputHitChecks("840x560", false);
        var date = Find<FrameworkElement>("Acquired"); var viewport = Find<ScrollViewer>("CollectionScroll");
        date.BringIntoView(); await Layout();
        Rect dateInViewport = date.TransformToAncestor(viewport).TransformBounds(new Rect(0, 0, date.ActualWidth, date.ActualHeight));
        Require(new Rect(0, 0, viewport.ActualWidth + 1, viewport.ActualHeight + 1).Contains(dateInViewport), "840x560 calendar scrolls fully into collection viewport");
        Screenshot("09b_最小窗口840x560.png");
        window.Width = 1216; window.Height = 820;
        var root = (FrameworkElement)window.Content;
        Transform old = root.LayoutTransform;
        foreach (double scale in new[] { 1.25, 1.50 })
        {
            root.LayoutTransform = new ScaleTransform(scale, scale); await Layout();
            Require(IsInsideWindow(Find<Button>("Save")), "LayoutTransform " + scale.ToString("P0", CultureInfo.InvariantCulture) + " simulation retains save button");
            await InputHitChecks("layout scale " + scale, false);
            Screenshot(scale == 1.25 ? "10_布局缩放模拟125.png" : "11_布局缩放模拟150.png");
        }
        root.LayoutTransform = old; await Layout();
        SetManual("TEST_ONLY焦点往返");
        for (int i = 0; i < 32; i++) { window.SwitchPage(i % 4, true); await Task.Delay(12); }
        window.SwitchPage(0, true); await Task.Delay(350); await Layout();
        Require(Find<TextBox>("Person").Text == "TEST_ONLY焦点往返", "32 rapid animated page switches across four pages preserve manual input");
        await InputHitChecks("after animated navigation", false);
        Require(!window.IsBusy, "navigation stress leaves no residual busy state");
        Screenshot("12_最终采集界面.png");
        Log("NOTE: 125%/150% checks use root LayoutTransform and are layout simulations, not physical monitor DPI verification.");
    }

    private static void CopyFolder(string source, string target)
    {
        Directory.CreateDirectory(target);
        foreach (string file in Directory.GetFiles(source)) File.Copy(file, Path.Combine(target, Path.GetFileName(file)), false);
        foreach (string child in Directory.GetDirectories(source)) CopyFolder(child, Path.Combine(target, Path.GetFileName(child)));
    }

    private static void Portability()
    {
        string relocated = Path.Combine(qaRoot, "TEST_ONLY换电脑目录"); CopyFolder(ownData, relocated);
        var moved = new InventoryStore(relocated);
        Require(moved.LoadRecords().Count == 4, "copied whole directory retains all four original records");
        var fifth = moved.Save(new InventoryRecord { DeviceKey = "TEST_ONLY_NEW_DEVICE", Site = "TEST_ONLY另一站", Location = "TEST_ONLY房间", AssignedTo = "TEST_ONLY公用", ManufacturerModel = "TEST_ONLY另一电脑" });
        Require(fifth.Count == 5 && fifth.BackupUpdated && WorkbookRows(moved.SummaryPath) == 6, "next-device directory relocation appends fifth record and backup");
        Require(new InventoryStore(ownData).LoadRecords().Count == 4, "relocation test leaves original QA records unchanged");
    }

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length != 1) { Console.Error.WriteLine("Usage: IosUiTests.exe <empty QA output directory>"); return 2; }
        qaRoot = Path.GetFullPath(args[0]);
        if (Directory.Exists(qaRoot) && Directory.GetFileSystemEntries(qaRoot).Length > 0) { Console.Error.WriteLine("QA directory must be empty."); return 2; }
        Directory.CreateDirectory(qaRoot); ownData = Path.Combine(qaRoot, "TEST_ONLY应用数据"); Directory.CreateDirectory(ownData);
        Log("WPF interaction harness: fabricated records, own HWND only, no WMI/network/external application control.");
        Log("NOTE: WM_CHAR checks Unicode routed text, not a physical Chinese IME or system clipboard.");
        var application = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        try
        {
            window = new InventoryWindow(ownData, true);
            window.ShowInTaskbar = false; window.ShowActivated = false; window.Opacity = 1;
            window.WindowStartupLocation = WindowStartupLocation.Manual; window.Left = -32000; window.Top = -32000;
            window.Loaded += async delegate
            {
                try
                {
                    ownHwnd = new WindowInteropHelper(window).Handle;
                    Require(ownHwnd != IntPtr.Zero, "own QA HWND created");
                    Require(DateValue() == "", "new date field defaults to unknown");
                    window.SetPreview(); await Layout();
                    Require(DateValue() == "2023-08-15", "preview full date is reflected in calendar field");
                    DateValue("2023"); Require(DateValue() == "2023", "legacy year date round-trips without inventing month/day");
                    DateValue("2023-08"); Require(DateValue() == "2023-08", "legacy month date round-trips without inventing a day");
                    DateValue(""); Require(DateValue() == "", "date field can clear to unknown");
                    window.Width = 1216; window.Height = 820; await Layout();
                    await InputHitChecks("normal", true);
                    await KeyboardAndAutomation();
                    await ButtonStateScreenshots();
                    await SaveRegression();
                    await LayoutAndNavigation();
                    await AboutAndNavigation();
                    Portability();
                    Log("ALL " + checks + " CHECKS PASSED; native point-click focus checks: " + nativePointChecks + ".");
                    Log("Screenshots require separate human/model visual review; their existence is not a pixel-perfect quality assertion.");
                }
                catch (Exception ex) { failure = ex; Log(ex.ToString()); }
                finally { application.Shutdown(); }
            };
            application.Run(window);
        }
        catch (Exception ex) { failure = ex; Log(ex.ToString()); }
        return failure == null ? 0 : 1;
    }
}
