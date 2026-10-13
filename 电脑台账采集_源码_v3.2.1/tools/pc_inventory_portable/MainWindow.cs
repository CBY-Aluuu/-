using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PcInventory
{
    public sealed class InventoryForm : Form
    {
        private readonly InventoryStore store;
        private readonly string baseDirectory;
        private readonly bool preview;
        private InventoryRecord current;
        private List<InventoryRecord> records = new List<InventoryRecord>();
        private readonly TextBox site = new TextBox(), location = new TextBox(), person = new TextBox(), number = new TextBox(), remarks = new TextBox();
        private readonly CalendarField acquired = new CalendarField();
        private readonly ComboBox type = new ComboBox(), state = new ComboBox();
        private readonly Label count = new Label(), status = new Label(), pageTitle = new Label(), pageDescription = new Label();
        private readonly Label modelTitle = new Label(), modelHint = new Label(), backupState = new Label();
        private readonly ActionButton save = new ActionButton(), refresh = new ActionButton(), summary = new ActionButton();
        private readonly ActionButton backup = new ActionButton(), restore = new ActionButton();
        private readonly StatusGlyph statusGlyph = new StatusGlyph();
        private readonly DataGridView hardware = new DataGridView(), history = new DataGridView();
        private readonly TextBox historyFilter = new TextBox();
        private readonly Label historyCount = new Label();
        private readonly LinkLabel statusDetails = new LinkLabel();
        private string fullStatus = "";
        private readonly FlowLayoutPanel navigation = new FlowLayoutPanel();
        private readonly Panel pageHost = new BufferedPanel();
        private readonly List<Panel> pages = new List<Panel>();
        private readonly List<ActionButton> navButtons = new List<ActionButton>();
        private readonly System.Windows.Forms.Timer transition = new System.Windows.Forms.Timer();
        private readonly ToolTip tips = new ToolTip();
        private DateTime transitionStarted;
        private int activePage;
        private bool busy, rebuiltOnStart;

        public InventoryForm(string directory, bool previewMode)
        {
            baseDirectory = Path.GetFullPath(directory); store = new InventoryStore(baseDirectory); preview = previewMode;
            Text = "电脑台账采集 · 追加备份版"; Font = new Font("Microsoft YaHei UI", 10F);
            var workArea = Screen.PrimaryScreen.WorkingArea;
            ClientSize = new Size(Math.Min(1216, Math.Max(1008, workArea.Width - 50)), Math.Min(820, Math.Max(650, workArea.Height - 70)));
            MinimumSize = new Size(1024, 700);
            StartPosition = FormStartPosition.CenterScreen; BackColor = AppTheme.Background; AutoScaleMode = AutoScaleMode.Font;
            DoubleBuffered = true;
            BuildLayout();
            transition.Interval = 15;
            transition.Tick += delegate
            {
                double t = Math.Min(1, (DateTime.UtcNow - transitionStarted).TotalMilliseconds / 170.0);
                double ease = 1 - Math.Pow(1 - t, 3);
                pages[activePage].SetBounds((int)(18 * (1 - ease)), 0, pageHost.ClientSize.Width, pageHost.ClientSize.Height);
                if (t >= 1) transition.Stop();
            };
            pageHost.Resize += delegate { foreach (var page in pages) page.SetBounds(0, 0, pageHost.ClientSize.Width, pageHost.ClientSize.Height); };
            if (!preview) Shown += async delegate { await ReadComputer(); };
            FormClosing += delegate(object sender, FormClosingEventArgs e)
            { if (busy) { e.Cancel = true; SetStatus("正在处理，请等待完成后再关闭程序。", "working"); } };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { transition.Dispose(); tips.Dispose(); }
            base.Dispose(disposing);
        }

        private static Label TextLabel(string text, float size, Color color, bool bold)
        {
            return new Label { Text = text, ForeColor = color, Font = new Font("Microsoft YaHei UI", size, bold ? FontStyle.Bold : FontStyle.Regular),
                AutoSize = false, BackColor = Color.Transparent, Margin = Padding.Empty, TextAlign = ContentAlignment.MiddleLeft };
        }

        private void BuildLayout()
        {
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 188)); layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(layout);
            var sidebar = new BufferedPanel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(19, 34, 53), Margin = Padding.Empty, Padding = new Padding(18, 26, 18, 20) };
            layout.Controls.Add(sidebar, 0, 0);
            var brand = new GlyphTile { Icon = IconKind.Computer, IconColor = Color.White, FillColor = Color.FromArgb(49, 107, 216), Size = new Size(44, 44), Location = new Point(22, 26) };
            sidebar.Controls.Add(brand);
            var brandTitle = TextLabel("电脑台账", 18, Color.White, true); brandTitle.SetBounds(22, 84, 144, 36); sidebar.Controls.Add(brandTitle);
            var brandSub = TextLabel("离线采集 · 随盘携带", 9, Color.FromArgb(153, 171, 194), false); brandSub.SetBounds(22, 121, 146, 25); sidebar.Controls.Add(brandSub);
            navigation.FlowDirection = FlowDirection.TopDown; navigation.WrapContents = false; navigation.SetBounds(18, 185, 152, 185); sidebar.Controls.Add(navigation);
            string[] navNames = { "设备采集", "历史记录", "备份与恢复" };
            IconKind[] navIcons = { IconKind.Computer, IconKind.Clipboard, IconKind.Shield };
            for (int i = 0; i < navNames.Length; i++)
            {
                int index = i; var button = new ActionButton { Text = navNames[i], Icon = navIcons[i], Size = new Size(152, 48), Margin = new Padding(0, 0, 0, 12), Primary = i == 0 };
                button.Click += delegate { SwitchPage(index, true); }; navButtons.Add(button); navigation.Controls.Add(button);
            }
            var offline = new Panel { Dock = DockStyle.Bottom, Height = 120, BackColor = Color.Transparent };
            var shield = new GlyphTile { Icon = IconKind.Shield, IconColor = Color.FromArgb(106, 211, 180), FillColor = Color.FromArgb(30, 52, 68), Location = new Point(4, 7), Size = new Size(28, 28) };
            offline.Controls.Add(shield);
            var offlineTitle = TextLabel("资料留在本地", 10, Color.FromArgb(218, 228, 240), true); offlineTitle.SetBounds(4, 42, 145, 26); offline.Controls.Add(offlineTitle);
            var offlineNote = TextLabel("本机读取，不联网\r\n整个文件夹一起带走", 8.5F, Color.FromArgb(153, 171, 194), false); offlineNote.SetBounds(4, 71, 145, 43); offline.Controls.Add(offlineNote); sidebar.Controls.Add(offline);

            var main = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(26, 20, 26, 14), Margin = Padding.Empty };
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 79)); main.RowStyles.Add(new RowStyle(SizeType.Absolute, 94));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); main.RowStyles.Add(new RowStyle(SizeType.Absolute, 74)); main.RowStyles.Add(new RowStyle(SizeType.Absolute, 23));
            layout.Controls.Add(main, 1, 0);
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Margin = Padding.Empty };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 78)); header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 41)); header.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            pageTitle.Font = new Font(Font.FontFamily, 22, FontStyle.Bold); pageTitle.ForeColor = AppTheme.Ink; pageTitle.Dock = DockStyle.Fill;
            pageDescription.Font = new Font(Font.FontFamily, 9.5F); pageDescription.ForeColor = AppTheme.Muted; pageDescription.Dock = DockStyle.Fill;
            header.Controls.Add(pageTitle, 0, 0); header.Controls.Add(pageDescription, 0, 1); header.SetColumnSpan(pageDescription, 2);
            var badge = TextLabel("●  离线运行", 9.5F, AppTheme.Success, true); badge.TextAlign = ContentAlignment.MiddleRight; badge.Dock = DockStyle.Fill; header.Controls.Add(badge, 1, 0);
            main.Controls.Add(header, 0, 0);

            var stats = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Margin = new Padding(0, 0, 0, 16) };
            stats.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            stats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28)); stats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32)); stats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            count.Text = "0 条"; count.Font = new Font(Font.FontFamily, 22, FontStyle.Bold); count.ForeColor = AppTheme.Ink;
            stats.Controls.Add(StatCard(IconKind.Database, "累计保存记录", count), 0, 0);
            stats.Controls.Add(StatCard(IconKind.Save, "每次点击保存", TextLabel("追加一条新记录", 12, AppTheme.Ink, true)), 1, 0);
            backupState.Text = "自动保留每次备份"; backupState.Font = new Font(Font.FontFamily, 11, FontStyle.Bold); backupState.ForeColor = AppTheme.Ink;
            stats.Controls.Add(StatCard(IconKind.Shield, "数据保护", backupState), 2, 0); main.Controls.Add(stats, 0, 1);

            pageHost.Dock = DockStyle.Fill; pageHost.Margin = Padding.Empty; main.Controls.Add(pageHost, 0, 2);
            pages.Add(BuildCollectionPage()); pages.Add(BuildHistoryPage()); pages.Add(BuildBackupPage());
            foreach (Panel page in pages) { pageHost.Controls.Add(page); page.Visible = false; }

            var feedback = new SurfacePanel { Dock = DockStyle.Fill, Radius = 12, Margin = new Padding(0, 12, 0, 6), Padding = new Padding(12, 5, 12, 5) };
            var feedbackGrid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Margin = Padding.Empty };
            feedbackGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 38)); feedbackGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); feedbackGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88));
            statusGlyph.Dock = DockStyle.Fill; statusGlyph.State = "idle"; status.Dock = DockStyle.Fill; status.Font = new Font(Font.FontFamily, 9.5F);
            status.TextAlign = ContentAlignment.MiddleLeft; status.ForeColor = AppTheme.Ink;
            status.AutoEllipsis = true;
            statusDetails.Text = "查看详情"; statusDetails.Dock = DockStyle.Fill; statusDetails.Margin = Padding.Empty; statusDetails.Font = new Font(Font.FontFamily, 9F); statusDetails.AutoEllipsis = true; statusDetails.TextAlign = ContentAlignment.MiddleRight; statusDetails.LinkColor = AppTheme.Accent; statusDetails.LinkBehavior = LinkBehavior.HoverUnderline;
            statusDetails.LinkClicked += delegate { MessageBox.Show(this, fullStatus, "操作详情", MessageBoxButtons.OK, statusGlyph.State == "warning" ? MessageBoxIcon.Warning : MessageBoxIcon.Information); };
            feedbackGrid.Controls.Add(statusGlyph, 0, 0); feedbackGrid.Controls.Add(status, 1, 0); feedbackGrid.Controls.Add(statusDetails, 2, 0); feedback.Controls.Add(feedbackGrid); main.Controls.Add(feedback, 0, 3);
            var path = TextLabel("保存位置：" + baseDirectory, 8, AppTheme.Muted, false); path.AutoEllipsis = true; path.Dock = DockStyle.Fill; tips.SetToolTip(path, baseDirectory); main.Controls.Add(path, 0, 4);
            SwitchPage(0, false); SetStatus("打开后自动读取。填写现场信息，保存后即可换下一台。", "idle");
        }

        private SurfacePanel StatCard(IconKind icon, string label, Label value)
        {
            var card = new SurfacePanel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 10, 0), Radius = 13, Padding = new Padding(12, 11, 12, 10) };
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Margin = Padding.Empty };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40)); grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var tile = new GlyphTile { Icon = icon, IconColor = AppTheme.Accent, FillColor = Color.FromArgb(235, 242, 254), Size = new Size(31, 31), Margin = new Padding(0, 5, 0, 0) };
            grid.Controls.Add(tile, 0, 0); grid.SetRowSpan(tile, 2); var title = TextLabel(label, 8.5F, AppTheme.Muted, false); title.Dock = DockStyle.Fill;
            grid.Controls.Add(title, 1, 0); value.Dock = DockStyle.Fill; grid.Controls.Add(value, 1, 1); card.Controls.Add(grid); return card;
        }

        private Panel BuildCollectionPage()
        {
            var page = new BufferedPanel();
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Margin = Padding.Empty };
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 58)); page.Controls.Add(grid);
            var viewport = new BufferedPanel { Dock = DockStyle.Fill, AutoScroll = true, Margin = Padding.Empty };
            var cards = new TableLayoutPanel { Dock = DockStyle.Top, Height = 420, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            viewport.Controls.Add(cards); viewport.Resize += delegate { cards.Height = Math.Max(420, viewport.ClientSize.Height); };
            cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 51)); cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 49)); grid.Controls.Add(viewport, 0, 0);
            var formCard = new SurfacePanel { Dock = DockStyle.Fill, Radius = 15, Padding = new Padding(18), Margin = new Padding(0, 0, 14, 0) };
            cards.Controls.Add(formCard, 0, 0);
            var formGrid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 6, Margin = Padding.Empty };
            formGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); formGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            formGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            for (int i = 0; i < 4; i++) formGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 25));
            formGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); formCard.Controls.Add(formGrid);
            var formTitle = SectionTitle(IconKind.Clipboard, "现场信息", "星号必填，不清楚可填“待核实”"); formGrid.Controls.Add(formTitle, 0, 0); formGrid.SetColumnSpan(formTitle, 2);
            formGrid.Controls.Add(InputField("站点 / 部门 *", site), 0, 1); formGrid.Controls.Add(InputField("房间 / 位置 *", location), 1, 1);
            formGrid.Controls.Add(InputField("实际使用人 *", person), 0, 2); formGrid.Controls.Add(InputField("资产编号 · 选填", number), 1, 2);
            type.DropDownStyle = ComboBoxStyle.DropDownList; type.Items.AddRange(new object[] { "台式机", "笔记本", "一体机", "其他", "待核实" });
            state.DropDownStyle = ComboBoxStyle.DropDownList; state.Items.AddRange(new object[] { "在用", "闲置", "待报废", "待核实" });
            formGrid.Controls.Add(InputField("电脑类型 · 核对", type), 0, 3); formGrid.Controls.Add(InputField("设备状态 · 核对", state), 1, 3);
            formGrid.Controls.Add(InputField("取得日期 · 日历选择", acquired), 0, 4); formGrid.Controls.Add(InputField("备注 · 选填", remarks), 1, 4);
            var hint = TextLabel("公用电脑填“公用”；日期未知就留空。", 8.5F, AppTheme.Muted, false); hint.Dock = DockStyle.Fill; formGrid.Controls.Add(hint, 0, 5); formGrid.SetColumnSpan(hint, 2);

            var autoCard = new SurfacePanel { Dock = DockStyle.Fill, Radius = 15, Padding = new Padding(18), Margin = Padding.Empty };
            cards.Controls.Add(autoCard, 1, 0);
            var autoGrid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = Padding.Empty };
            autoGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 52)); autoGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 76)); autoGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            autoGrid.Controls.Add(SectionTitle(IconKind.Computer, "本机硬件", "自动读取结果，保存前请核对"), 0, 0);
            var deviceHero = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(240, 245, 253), Padding = new Padding(10), Margin = new Padding(0, 0, 0, 10) };
            var heroGrid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Margin = Padding.Empty };
            heroGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 45)); heroGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            heroGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 62)); heroGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 38));
            var computer = new GlyphTile { Icon = IconKind.Computer, FillColor = Color.FromArgb(240, 245, 253), IconColor = AppTheme.Accent, Dock = DockStyle.Fill };
            heroGrid.Controls.Add(computer, 0, 0); heroGrid.SetRowSpan(computer, 2);
            modelTitle.Font = new Font(Font.FontFamily, 11, FontStyle.Bold); modelTitle.ForeColor = AppTheme.Ink; modelTitle.AutoEllipsis = true; modelTitle.Dock = DockStyle.Fill; modelTitle.Text = "等待读取这台电脑";
            modelHint.Font = new Font(Font.FontFamily, 8.5F); modelHint.ForeColor = AppTheme.Muted; modelHint.Dock = DockStyle.Fill; modelHint.Text = "只读采集 · 不修改电脑设置";
            heroGrid.Controls.Add(modelTitle, 1, 0); heroGrid.Controls.Add(modelHint, 1, 1); deviceHero.Controls.Add(heroGrid); autoGrid.Controls.Add(deviceHero, 0, 1);
            ConfigureGrid(hardware); hardware.ColumnHeadersVisible = false; hardware.Columns.Add(new DataGridViewTextBoxColumn { Name = "item", Width = 92 });
            hardware.Columns.Add(new DataGridViewTextBoxColumn { Name = "value", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill }); autoGrid.Controls.Add(hardware, 0, 2); autoCard.Controls.Add(autoGrid);

            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(0, 12, 0, 0), Margin = Padding.Empty };
            ConfigureAction(save, "保存并追加一条", IconKind.Save, 192, true); ConfigureAction(refresh, "重新读取", IconKind.Refresh, 130, false); ConfigureAction(summary, "打开 / 更新总表", IconKind.Clipboard, 178, false);
            save.Click += async delegate { await SaveComputer(); }; refresh.Click += async delegate { await ReadComputer(); }; summary.Click += async delegate { await OpenSummary(); };
            actions.Controls.Add(save); actions.Controls.Add(refresh); actions.Controls.Add(summary); grid.Controls.Add(actions, 0, 1); save.Enabled = false;
            return page;
        }

        private static void ConfigureGrid(DataGridView grid)
        {
            grid.Dock = DockStyle.Fill; grid.ReadOnly = true; grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false; grid.RowHeadersVisible = false; grid.BackgroundColor = Color.White; grid.BorderStyle = BorderStyle.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal; grid.GridColor = Color.FromArgb(235, 239, 245);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(232, 241, 255); grid.DefaultCellStyle.SelectionForeColor = AppTheme.Ink;
            grid.DefaultCellStyle.ForeColor = AppTheme.Ink; grid.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F); grid.DefaultCellStyle.Padding = new Padding(3, 7, 3, 7);
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells; grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.EnableHeadersVisualStyles = false; grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(244, 247, 251);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.Muted; grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            grid.ColumnHeadersHeight = 36; grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(grid, true, null);
        }

        private Panel InputField(string caption, Control input)
        {
            var wrapper = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Margin = new Padding(0, 0, 10, 5) };
            wrapper.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); wrapper.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            var label = TextLabel(caption, 8.7F, AppTheme.Muted, false); label.Dock = DockStyle.Fill; wrapper.Controls.Add(label, 0, 0);
            input.Dock = DockStyle.Fill; input.Margin = Padding.Empty; input.Font = new Font(Font.FontFamily, 9.5F);
            input.AccessibleName = caption;
            if (input is TextBox)
            {
                var box = new SurfacePanel { Dock = DockStyle.Fill, Radius = 7, Padding = new Padding(8, 7, 8, 5), Margin = Padding.Empty };
                ((TextBox)input).BorderStyle = BorderStyle.None; input.Enter += delegate { box.BorderColor = AppTheme.Accent; box.Invalidate(); };
                input.Leave += delegate { box.BorderColor = AppTheme.Border; box.Invalidate(); };
                box.Controls.Add(input); wrapper.Controls.Add(box, 0, 1);
            }
            else
            {
                if (input is ComboBox) { input.Margin = new Padding(0, 4, 0, 0); ((ComboBox)input).FlatStyle = FlatStyle.Flat; }
                wrapper.Controls.Add(input, 0, 1);
            }
            return wrapper;
        }

        private Panel SectionTitle(IconKind icon, string title, string subtitle)
        {
            var panel = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            var glyph = new GlyphTile { Icon = icon, IconColor = AppTheme.Accent, FillColor = Color.White, Location = new Point(0, 2), Size = new Size(25, 25) };
            var heading = TextLabel(title, 12, AppTheme.Ink, true); heading.SetBounds(32, 0, 250, 28); heading.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            var sub = TextLabel(subtitle, 8.5F, AppTheme.Muted, false); sub.SetBounds(0, 28, 410, 20); sub.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            panel.Controls.Add(glyph); panel.Controls.Add(heading); panel.Controls.Add(sub); return panel;
        }

        private static void ConfigureAction(ActionButton button, string text, IconKind icon, int width, bool primary)
        { button.Text = text; button.Icon = icon; button.Size = new Size(width, 44); button.Primary = primary; button.Margin = new Padding(0, 0, 12, 0); }

        private Panel BuildHistoryPage()
        {
            var page = new BufferedPanel();
            var card = new SurfacePanel { Dock = DockStyle.Fill, Radius = 15, Padding = new Padding(18), Margin = Padding.Empty }; page.Controls.Add(card);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = Padding.Empty };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 29)); card.Controls.Add(layout);
            var tools = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            tools.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65)); tools.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            tools.Controls.Add(SectionTitle(IconKind.Clock, "每次采集都有记录", "按保存顺序排列，同一电脑多次采集也逐条保留"), 0, 0);
            historyFilter.Dock = DockStyle.Fill; historyFilter.Margin = new Padding(0, 8, 0, 14); historyFilter.AccessibleName = "搜索站点、位置、使用人、型号"; tips.SetToolTip(historyFilter, "输入站点、位置、使用人或型号筛选"); historyFilter.TextChanged += delegate { RefreshHistory(); };
            tools.Controls.Add(historyFilter, 1, 0); layout.Controls.Add(tools, 0, 0);
            ConfigureGrid(history);
            string[] titles = { "序号", "站点", "位置", "使用人", "品牌型号", "采集时间" }; int[] widths = { 52, 110, 110, 85, 220, 155 };
            for (int i = 0; i < titles.Length; i++) history.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = titles[i], Width = widths[i], Name = "c" + i, SortMode = DataGridViewColumnSortMode.NotSortable });
            history.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            history.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                var record = history.Rows[e.RowIndex].Tag as InventoryRecord;
                if (record != null) ShowRecordDetails(record);
            };
            layout.Controls.Add(history, 0, 1); historyCount.Dock = DockStyle.Fill; historyCount.Font = new Font(Font.FontFamily, 8.5F); historyCount.ForeColor = AppTheme.Muted;
            layout.Controls.Add(historyCount, 0, 2); return page;
        }

        private Panel BuildBackupPage()
        {
            var page = new BufferedPanel();
            var card = new SurfacePanel { Dock = DockStyle.Fill, Radius = 15, Padding = new Padding(24), Margin = Padding.Empty }; page.Controls.Add(card);
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Margin = Padding.Empty };
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 59)); grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 100)); grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 69));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 67)); grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 63)); grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); card.Controls.Add(grid);
            grid.Controls.Add(SectionTitle(IconKind.Shield, "每次保存，留下一份独立备份", "原始记录、汇总表和版本备份分别保留"), 0, 0);
            var diagram = new BackupDiagram { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 12) }; grid.Controls.Add(diagram, 0, 1);
            var explanation = TextLabel("自动备份保存在本工具的“备份”文件夹，每份都有独立文件名。\r\n同一U盘上的备份无法防止U盘丢失或损坏，建议另备份到单位电脑或另一块盘。", 10, AppTheme.Muted, false);
            explanation.Dock = DockStyle.Fill; grid.Controls.Add(explanation, 0, 2);
            var row = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Margin = Padding.Empty, Padding = new Padding(0, 10, 0, 0) };
            ConfigureAction(backup, "另存备份到其他盘", IconKind.Backup, 226, true); ConfigureAction(restore, "从备份恢复", IconKind.Restore, 176, false);
            var open = new ActionButton(); ConfigureAction(open, "打开备份文件夹", IconKind.Folder, 185, false);
            backup.Click += async delegate { await BackupToFolder(); }; restore.Click += async delegate { await RestoreFromBackup(); };
            open.Click += delegate { OpenFolder(Path.Combine(baseDirectory, "备份")); };
            row.Controls.Add(backup); row.Controls.Add(restore); row.Controls.Add(open); grid.Controls.Add(row, 0, 3);
            var restoreHint = TextLabel("恢复时会生成一个新文件夹，现有记录不会被覆盖。\r\n总表是自动生成的；需要手工修改时，请先另存一份。", 9.3F, AppTheme.Muted, false); restoreHint.Dock = DockStyle.Fill; grid.Controls.Add(restoreHint, 0, 4);
            return page;
        }

        private void SwitchPage(int index, bool animate)
        {
            transition.Stop(); activePage = index;
            for (int i = 0; i < pages.Count; i++) { pages[i].Visible = i == index; navButtons[i].Primary = i == index; navButtons[i].Invalidate(); }
            string[] titles = { "设备采集", "历史记录", "备份与恢复" };
            string[] descriptions = { "自动读取，现场确认。每次保存都追加一条，旧记录始终保留。", "查看每一次保存的结果，双击记录可查看完整信息。", "日常自动留存，重要资料另存一份。" };
            pageTitle.Text = titles[index]; pageDescription.Text = descriptions[index];
            if (pages.Count > index)
            {
                pages[index].BringToFront(); pages[index].SetBounds(0, 0, pageHost.Width, pageHost.Height);
                if (animate && AppTheme.Animate) { transitionStarted = DateTime.UtcNow; transition.Start(); }
            }
            if (index == 1) RefreshHistory();
        }

        private void SetBusy(bool value, ActionButton active)
        {
            busy = value;
            foreach (Control control in new Control[] { site, location, person, number, remarks, acquired, type, state }) control.Enabled = !value;
            foreach (ActionButton button in new[] { save, refresh, summary, backup, restore }) { button.Busy = value && button == active; button.Enabled = !value; }
            save.Enabled = !value && current != null;
        }

        private void SetStatus(string message, string kind)
        {
            fullStatus = message;
            string compact = message.Replace("\r", " ").Replace("\n", " ");
            status.Text = compact.Length > 95 ? compact.Substring(0, 92) + "…" : compact;
            statusDetails.Visible = kind == "warning" || compact.Length > 95;
            status.ForeColor = kind == "warning" ? Color.FromArgb(155, 79, 30) : AppTheme.Ink; statusGlyph.State = kind; tips.SetToolTip(status, message);
        }

        private async Task ReadComputer()
        {
            if (busy) return; SetBusy(true, refresh); SetStatus("正在读取本机硬件，界面仍可切换。请稍候……", "working");
            try
            {
                var loaded = await Task.Run(() => store.LoadRecords());
                string rebuildProblem = "";
                if (!rebuiltOnStart && loaded.Count > 0)
                {
                    try { await Task.Run(() => store.RebuildSummary()); rebuiltOnStart = true; }
                    catch { rebuildProblem = " 已加载历史记录，总表暂未更新；关闭Excel后可点更新总表。"; }
                }
                var read = await Task.Run(() => HardwareCollector.Collect());
                var previous = loaded.LastOrDefault(r => !String.IsNullOrEmpty(read.DeviceKey) && r.DeviceKey == read.DeviceKey);
                bool preserveInputs = current != null; current = read; records = loaded;
                if (!preserveInputs)
                {
                    if (previous != null) FillManual(previous);
                    else { type.SelectedItem = read.DeviceType ?? "待核实"; state.SelectedItem = "待核实"; acquired.DateValue = ""; }
                }
                ShowHardware(read); UpdateCount(); RefreshHistory();
                SetStatus((previous == null ? "读取完成。核对现场信息，点击保存即可追加一条记录。" : "已带出这台电脑上次填写的信息。本次保存仍会新增一条，旧记录不变。") + rebuildProblem, rebuildProblem.Length > 0 ? "warning" : "success");
            }
            catch (Exception e) { SetStatus("读取未完成：" + e.Message, "warning"); }
            finally { SetBusy(false, null); }
        }

        private void FillManual(InventoryRecord r)
        {
            site.Text = r.Site; location.Text = r.Location; person.Text = r.AssignedTo; number.Text = r.DeviceNumber; acquired.DateValue = r.AcquiredDate ?? ""; remarks.Text = r.Remarks;
            type.SelectedItem = type.Items.Contains(r.DeviceType ?? "") ? r.DeviceType : "待核实";
            state.SelectedItem = state.Items.Contains(r.DeviceStatus ?? "") ? r.DeviceStatus : "待核实";
        }

        private void ShowHardware(InventoryRecord r)
        {
            modelTitle.Text = r.ManufacturerModel ?? "未读到型号，待核实"; tips.SetToolTip(modelTitle, modelTitle.Text);
            modelHint.Text = (r.DeviceType ?? "待核实") + "  ·  数量 1";
            hardware.Rows.Clear();
            foreach (var pair in new[] { new[] { "CPU", r.Cpu }, new[] { "内存", r.Memory }, new[] { "系统", r.OperatingSystem }, new[] { "整机序列号", r.BiosSerial }, new[] { "物理硬盘", r.DiskSerials }, new[] { "网卡 / IP", r.AdapterPairs }, new[] { "核对提示", r.Verification } })
                hardware.Rows.Add(pair[0], String.IsNullOrWhiteSpace(pair[1]) ? "未读取到，待核实" : pair[1]);
            hardware.ClearSelection();
        }

        private void UpdateCount() { count.Text = records.Count + " 条"; }

        private async Task SaveComputer()
        {
            if (busy || current == null) return;
            if (String.IsNullOrWhiteSpace(site.Text) || String.IsNullOrWhiteSpace(location.Text) || String.IsNullOrWhiteSpace(person.Text))
            { SetStatus("请填站点、位置和实际使用人。不清楚填“待核实”，公用电脑填“公用”。", "warning"); return; }
            current.Site = site.Text.Trim(); current.Location = location.Text.Trim(); current.AssignedTo = person.Text.Trim(); current.DeviceNumber = number.Text.Trim();
            current.AcquiredDate = acquired.DateValue; current.Remarks = remarks.Text.Trim(); current.DeviceType = Convert.ToString(type.SelectedItem); current.DeviceStatus = Convert.ToString(state.SelectedItem);
            current.RecordId = Guid.NewGuid().ToString("N");
            SetBusy(true, save); SetStatus("正在追加记录、更新总表并生成备份。请不要拔出U盘……", "working");
            try
            {
                var result = await Task.Run(() => store.Save(current));
                count.Text = result.Count + " 条";
                backupState.Text = result.BackupUpdated ? "本次备份已完成" : "本次备份待完成";
                backupState.ForeColor = result.BackupUpdated ? AppTheme.Success : Color.FromArgb(163, 95, 32);
                string brief = "已追加第 " + result.Count + " 条，旧记录保留。" + (result.SummaryUpdated ? "总表已更新；" : "总表待更新；") + (result.BackupUpdated ? "完整备份已保存。" : "本次备份失败，请另存备份。" );
                SetStatus(brief, result.SummaryUpdated && result.BackupUpdated ? "success" : "warning"); fullStatus = result.Message; tips.SetToolTip(status, result.Message);
                try { records = await Task.Run(() => store.LoadRecords()); RefreshHistory(); }
                catch { SetStatus(result.Message + " 历史列表暂未刷新，可重新打开程序。", "warning"); }
            }
            catch (Exception e) { SetStatus("保存未完成，请保留本窗口并检查磁盘是否可写。" + e.Message, "warning"); }
            finally { SetBusy(false, null); }
        }

        private async Task OpenSummary()
        {
            if (busy) return; SetBusy(true, summary); SetStatus("正在用全部历史记录更新总表……", "working");
            try
            {
                string path = await Task.Run(() => store.RebuildSummary()); records = await Task.Run(() => store.LoadRecords()); UpdateCount(); RefreshHistory();
                SetStatus("总表已按全部记录更新。手工编辑请另存一份，以免下次生成时覆盖你的表格改动。", "success");
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception e) { SetStatus("总表暂时无法更新或打开，已保存记录仍保留。请关闭Excel后重试。" + e.Message, "warning"); }
            finally { SetBusy(false, null); }
        }

        private void RefreshHistory()
        {
            history.Rows.Clear(); string filter = historyFilter.Text.Trim(); int shown = 0;
            for (int i = 0; i < records.Count; i++)
            {
                InventoryRecord r = records[i]; string text = (r.Site + " " + r.Location + " " + r.AssignedTo + " " + r.ManufacturerModel + " " + r.DeviceNumber);
                if (filter.Length > 0 && text.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0) continue;
                int row = history.Rows.Add(i + 1, r.Site, r.Location, r.AssignedTo, r.ManufacturerModel, ReadableTime(r.CollectedAt)); history.Rows[row].Tag = r; shown++;
            }
            history.ClearSelection(); historyCount.Text = "显示 " + shown + " / " + records.Count + " 条记录  ·  双击查看详情  ·  保存顺序不会因电脑时钟变化而打乱";
        }

        private static string ReadableTime(string value) { DateTimeOffset time; return DateTimeOffset.TryParse(value, out time) ? time.ToString("yyyy-MM-dd HH:mm:ss") : value ?? "未记录"; }

        private void ShowRecordDetails(InventoryRecord r)
        {
            using (var dialog = new Form { Text = "采集记录详情", StartPosition = FormStartPosition.CenterParent, Size = new Size(720, 620), Font = Font, MinimizeBox = false, MaximizeBox = false })
            {
                var text = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.White, BorderStyle = BorderStyle.None, Font = Font };
                text.Text = "站点：" + r.Site + "\r\n位置：" + r.Location + "\r\n使用人：" + r.AssignedTo + "\r\n资产编号：" + r.DeviceNumber + "\r\n取得日期：" + r.AcquiredDate + "\r\n类型 / 状态：" + r.DeviceType + " / " + r.DeviceStatus +
                    "\r\n\r\n品牌型号：" + r.ManufacturerModel + "\r\nCPU：" + r.Cpu + "\r\n内存：" + r.Memory + "\r\n系统：" + r.OperatingSystem + "\r\n整机序列号：" + r.BiosSerial + "\r\n硬盘：" + r.DiskSerials + "\r\n网卡：" + r.AdapterPairs +
                    "\r\n\r\n采集时间：" + r.CollectedAt + "\r\n备注：" + r.Remarks + "\r\n核对提示：" + r.Verification;
                dialog.Padding = new Padding(22); dialog.Controls.Add(text); dialog.ShowDialog(this);
            }
        }

        private async Task BackupToFolder()
        {
            if (busy) return;
            using (var picker = new FolderBrowserDialog { Description = "选择单位允许的本地目录，建议选另一块盘或单位电脑里的文件夹。", ShowNewFolderButton = true })
            {
                if (picker.ShowDialog(this) != DialogResult.OK) return;
                SetBusy(true, backup); SetStatus("正在创建完整备份，请稍候……", "working");
                try
                {
                    string path = await Task.Run(() => store.BackupNow(picker.SelectedPath));
                    backupState.Text = "手动备份已完成"; backupState.ForeColor = AppTheme.Success;
                    bool sameDrive = String.Equals(Path.GetPathRoot(path), Path.GetPathRoot(baseDirectory), StringComparison.OrdinalIgnoreCase);
                    SetStatus("备份成功：" + path + (sameDrive ? "。与原数据在同一盘，建议再备份到另一块盘。" : "。已保留到所选位置。"), "success");
                }
                catch (Exception e) { SetStatus("备份未完成，原记录没有改动。" + e.Message, "warning"); }
                finally { SetBusy(false, null); }
            }
        }

        private async Task RestoreFromBackup()
        {
            if (busy) return;
            using (var file = new OpenFileDialog { Title = "选择本工具生成的备份", Filter = "台账备份 (*.zip)|*.zip", CheckFileExists = true })
            using (var target = new FolderBrowserDialog { Description = "选择恢复位置。工具会在其中新建文件夹，不覆盖任何现有记录。", ShowNewFolderButton = true })
            {
                if (file.ShowDialog(this) != DialogResult.OK || target.ShowDialog(this) != DialogResult.OK) return;
                SetBusy(true, restore); SetStatus("正在检查备份并恢复到新文件夹……", "working");
                try
                {
                    string restored = await Task.Run(() => store.RestoreBackupToNewFolder(file.FileName, target.SelectedPath));
                    string executable = Path.Combine(restored, "电脑台账采集.exe");
                    try { File.Copy(Application.ExecutablePath, executable, false); }
                    catch { SetStatus("数据已恢复至 " + restored + "；程序复制未完成，请将本exe复制进去后使用。现有目录未改动。", "warning"); OpenFolder(restored); return; }
                    SetStatus("恢复成功，原目录未改动。请从新文件夹打开程序继续使用：" + restored, "success"); OpenFolder(restored);
                }
                catch (Exception e) { SetStatus("恢复未完成，现有数据未改动。" + e.Message, "warning"); }
                finally { SetBusy(false, null); }
            }
        }

        private void OpenFolder(string path)
        {
            try { Directory.CreateDirectory(path); Process.Start("explorer.exe", "\"" + path + "\""); }
            catch (Exception e) { SetStatus("暂时无法打开文件夹。" + e.Message, "warning"); }
        }

        public void SetPreview()
        {
            current = new InventoryRecord { DeviceKey = "DEMO", Site = "示例管理站", Location = "一楼办公室", AssignedTo = "公用", DeviceType = "台式机", DeviceStatus = "在用", AcquiredDate = "2023-08-15",
                ManufacturerModel = "示例品牌 / 办公电脑", Cpu = "示例处理器 · 6 核", Memory = "16 GiB", OperatingSystem = "Windows 11 · 64 位", BiosSerial = "DEMO-001", DiskSerials = "硬盘0：示例固态硬盘 · 512 GB", AdapterPairs = "有线网卡\r\nMAC 02:00:00:00:00:01\r\nIP 192.0.2.1", Verification = "仅为界面示例，非真实资产信息。" };
            FillManual(current); ShowHardware(current); save.Enabled = true; count.Text = "12 条"; backupState.Text = "最近一次备份已完成"; backupState.ForeColor = AppTheme.Success;
            SetStatus("采集已完成。核对现场信息，点击“保存并追加一条”；旧记录不会被合并。", "success");
        }
    }

    internal sealed class BufferedPanel : Panel
    { public BufferedPanel() { DoubleBuffered = true; ResizeRedraw = true; } }

    internal sealed class GlyphTile : Control
    {
        public IconKind Icon { get; set; } public Color IconColor { get; set; } public Color FillColor { get; set; }
        public GlyphTile() { DoubleBuffered = true; ResizeRedraw = true; }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias; using (var brush = new SolidBrush(FillColor)) e.Graphics.FillRectangle(brush, ClientRectangle);
            float size = Math.Min(Width, Height) * .64F; AppTheme.DrawIcon(e.Graphics, Icon, new RectangleF((Width - size) / 2, (Height - size) / 2, size, size), IconColor, 1.8F);
        }
    }

    internal sealed class BackupDiagram : Control
    {
        public BackupDiagram() { DoubleBuffered = true; ResizeRedraw = true; }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            string[] labels = { "逐条原始记录", "全部记录总表", "独立版本备份" }; IconKind[] icons = { IconKind.Database, IconKind.Clipboard, IconKind.Shield };
            float unit = Width / 3F;
            for (int i = 0; i < 3; i++)
            {
                RectangleF box = new RectangleF(i * unit, 0, unit - 28, Height - 1);
                using (var brush = new SolidBrush(Color.FromArgb(239, 245, 254))) e.Graphics.FillRectangle(brush, box);
                AppTheme.DrawIcon(e.Graphics, icons[i], new RectangleF(box.X + 15, 16, 28, 28), AppTheme.Accent, 2);
                using (var font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold))
                    TextRenderer.DrawText(e.Graphics, labels[i], font, new Rectangle((int)box.X + 54, 0, (int)box.Width - 59, Height), AppTheme.Ink, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
                if (i < 2) AppTheme.DrawIcon(e.Graphics, IconKind.Arrow, new RectangleF(box.Right + 4, Height / 2 - 10, 20, 20), AppTheme.Muted, 1.5F);
            }
        }
    }
}
