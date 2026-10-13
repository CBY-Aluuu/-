using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PcInventory
{
    public enum IconKind { Computer, Clipboard, Shield, Folder, Calendar, Refresh, Save, Check, Arrow, Database, Clock, Chevron, Usb, Backup, Restore, Info, Warning }

    public static class AppTheme
    {
        public static readonly Color Background = Color.FromArgb(241, 245, 250);
        public static readonly Color Surface = Color.White;
        public static readonly Color Ink = Color.FromArgb(24, 40, 61);
        public static readonly Color Muted = Color.FromArgb(109, 125, 144);
        public static readonly Color Accent = Color.FromArgb(32, 101, 224);
        public static readonly Color Border = Color.FromArgb(222, 230, 239);
        public static readonly Color Success = Color.FromArgb(28, 146, 107);

        internal static GraphicsPath Round(RectangleF r, float radius)
        {
            GraphicsPath path = new GraphicsPath();
            float d = Math.Max(0, Math.Min(radius * 2, Math.Min(r.Width, r.Height)));
            if (d < 1) { path.AddRectangle(r); return path; }
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        internal static Color Mix(Color a, Color b, float amount)
        {
            amount = Math.Max(0, Math.Min(1, amount));
            return Color.FromArgb((int)(a.R + (b.R - a.R) * amount), (int)(a.G + (b.G - a.G) * amount), (int)(a.B + (b.B - a.B) * amount));
        }

        private static bool motionAllowed = true;
        private static int motionCheckedAt = unchecked(Environment.TickCount - 2000);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SystemParametersInfo(uint action, uint param, [MarshalAs(UnmanagedType.Bool)] out bool value, uint flags);
        internal static bool Animate
        {
            get
            {
                // WinForms on .NET 4.x does not expose ClientAreaAnimation. Read the
                // same Windows preference directly and refresh when it changes.
                int now = Environment.TickCount;
                if (unchecked(now - motionCheckedAt) >= 1000)
                {
                    bool enabled;
                    motionAllowed = SystemParametersInfo(0x1042, 0, out enabled, 0) ? enabled : false;
                    motionCheckedAt = now;
                }
                return motionAllowed && !SystemInformation.TerminalServerSession;
            }
        }

        public static void DrawIcon(Graphics g, IconKind kind, RectangleF bounds, Color color, float stroke = 2)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0) return;
            GraphicsState saved = g.Save();
            try
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TranslateTransform(bounds.X, bounds.Y);
                g.ScaleTransform(bounds.Width / 24F, bounds.Height / 24F);
                using (Pen p = new Pen(color, stroke * 24F / Math.Min(bounds.Width, bounds.Height)))
                {
                    p.StartCap = LineCap.Round; p.EndCap = LineCap.Round; p.LineJoin = LineJoin.Round;
                    switch (kind)
                    {
                        case IconKind.Computer:
                            using (GraphicsPath box = Round(new RectangleF(3, 3, 18, 13), 2)) g.DrawPath(p, box);
                            g.DrawLine(p, 12, 16, 12, 21); g.DrawLine(p, 7, 21, 17, 21); break;
                        case IconKind.Clipboard:
                            using (GraphicsPath box = Round(new RectangleF(5, 4, 14, 18), 2)) g.DrawPath(p, box);
                            using (GraphicsPath clip = Round(new RectangleF(9, 2, 6, 4), 1)) { using (Brush b = new SolidBrush(Surface)) g.FillPath(b, clip); g.DrawPath(p, clip); }
                            g.DrawLine(p, 9, 11, 15, 11); g.DrawLine(p, 9, 15, 15, 15); g.DrawLine(p, 9, 18, 13, 18); break;
                        case IconKind.Shield:
                            using (GraphicsPath shield = new GraphicsPath())
                            {
                                shield.AddLines(new PointF[] { new PointF(12, 2), new PointF(21, 6), new PointF(20, 13) });
                                shield.AddBezier(20, 13, 19, 18, 15, 21, 12, 22);
                                shield.AddBezier(12, 22, 9, 21, 5, 18, 4, 13);
                                shield.AddLine(4, 13, 3, 6); shield.CloseFigure(); g.DrawPath(p, shield);
                            }
                            g.DrawLines(p, new PointF[] { new PointF(8, 12), new PointF(11, 15), new PointF(16, 9) }); break;
                        case IconKind.Folder:
                            using (GraphicsPath folder = new GraphicsPath())
                            {
                                folder.AddLines(new PointF[] { new PointF(3, 7), new PointF(3, 4), new PointF(9, 4), new PointF(12, 7), new PointF(21, 7), new PointF(21, 20), new PointF(3, 20) }); folder.CloseFigure(); g.DrawPath(p, folder);
                            }
                            g.DrawLine(p, 3, 10, 21, 10); break;
                        case IconKind.Calendar:
                            using (GraphicsPath box = Round(new RectangleF(3, 5, 18, 16), 2)) g.DrawPath(p, box);
                            g.DrawLine(p, 3, 10, 21, 10); g.DrawLine(p, 8, 2, 8, 7); g.DrawLine(p, 16, 2, 16, 7);
                            g.DrawLine(p, 8, 14, 9, 14); g.DrawLine(p, 15, 14, 16, 14); g.DrawLine(p, 8, 18, 9, 18); break;
                        case IconKind.Refresh:
                            g.DrawArc(p, 4, 4, 16, 16, 35, 285); g.DrawLines(p, new PointF[] { new PointF(14, 4), new PointF(20, 4), new PointF(20, 10) }); break;
                        case IconKind.Save:
                            using (GraphicsPath save = new GraphicsPath())
                            {
                                save.AddLines(new PointF[] { new PointF(4, 3), new PointF(17, 3), new PointF(21, 7), new PointF(21, 21), new PointF(3, 21), new PointF(3, 4) }); save.CloseFigure(); g.DrawPath(p, save);
                            }
                            g.DrawRectangle(p, 7, 3, 9, 6); g.DrawRectangle(p, 7, 14, 10, 7); break;
                        case IconKind.Check:
                            g.DrawLines(p, new PointF[] { new PointF(4, 12), new PointF(9, 17), new PointF(20, 6) }); break;
                        case IconKind.Arrow:
                            g.DrawLine(p, 4, 12, 20, 12); g.DrawLines(p, new PointF[] { new PointF(14, 6), new PointF(20, 12), new PointF(14, 18) }); break;
                        case IconKind.Database:
                            g.DrawEllipse(p, 4, 3, 16, 6); g.DrawLine(p, 4, 6, 4, 18); g.DrawLine(p, 20, 6, 20, 18);
                            g.DrawArc(p, 4, 9, 16, 6, 0, 180); g.DrawArc(p, 4, 15, 16, 6, 0, 180); break;
                        case IconKind.Clock:
                            g.DrawEllipse(p, 3, 3, 18, 18); g.DrawLines(p, new PointF[] { new PointF(12, 7), new PointF(12, 12), new PointF(16, 14) }); break;
                        case IconKind.Chevron:
                            g.DrawLines(p, new PointF[] { new PointF(6, 9), new PointF(12, 15), new PointF(18, 9) }); break;
                        case IconKind.Usb:
                            using (GraphicsPath body = Round(new RectangleF(6, 8, 12, 14), 2)) g.DrawPath(p, body);
                            g.DrawRectangle(p, 8, 2, 8, 6); g.DrawLine(p, 11, 4, 11, 6); g.DrawLine(p, 14, 4, 14, 6); g.DrawLine(p, 12, 13, 12, 17); break;
                        case IconKind.Backup:
                            g.DrawArc(p, 3, 3, 18, 18, 210, 310); g.DrawLines(p, new PointF[] { new PointF(2, 3), new PointF(3, 9), new PointF(9, 8) });
                            g.DrawLine(p, 12, 17, 12, 8); g.DrawLines(p, new PointF[] { new PointF(8, 12), new PointF(12, 8), new PointF(16, 12) }); break;
                        case IconKind.Restore:
                            g.DrawArc(p, 3, 3, 18, 18, 210, 310); g.DrawLines(p, new PointF[] { new PointF(2, 3), new PointF(3, 9), new PointF(9, 8) });
                            g.DrawLines(p, new PointF[] { new PointF(12, 7), new PointF(12, 12), new PointF(16, 14) }); break;
                        case IconKind.Info:
                            g.DrawEllipse(p, 3, 3, 18, 18); g.DrawLine(p, 12, 11, 12, 17); g.DrawLine(p, 12, 7, 12, 7.2F); break;
                        case IconKind.Warning:
                            g.DrawPolygon(p, new PointF[] { new PointF(12, 3), new PointF(22, 21), new PointF(2, 21) }); g.DrawLine(p, 12, 9, 12, 14); g.DrawLine(p, 12, 17, 12, 17.2F); break;
                    }
                }
            }
            finally { g.Restore(saved); }
        }
    }

    public class SurfacePanel : Panel
    {
        public int Radius = 18;
        public Color BorderColor = AppTheme.Border;
        public bool DrawShadow;
        public SurfacePanel()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = AppTheme.Surface;
        }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(Parent == null ? AppTheme.Background : Parent.BackColor);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            RectangleF r = new RectangleF(0.5F, 0.5F, Math.Max(0, Width - 1F), Math.Max(0, Height - (DrawShadow ? 5F : 1F)));
            if (r.Width < 1 || r.Height < 1) return;
            if (DrawShadow)
            {
                for (int i = 3; i >= 1; i--)
                    using (GraphicsPath p = AppTheme.Round(new RectangleF(r.X + 1, r.Y + i, r.Width - 2, r.Height), Radius))
                    using (Brush b = new SolidBrush(Color.FromArgb(5, 32, 49, 72))) e.Graphics.FillPath(b, p);
            }
            using (GraphicsPath path = AppTheme.Round(r, Radius))
            {
                using (Brush b = new SolidBrush(BackColor)) e.Graphics.FillPath(b, path);
                using (Pen p = new Pen(BorderColor)) e.Graphics.DrawPath(p, path);
            }
        }
    }

    public class ActionButton : Button
    {
        private readonly Timer animation;
        private readonly Stopwatch clock = new Stopwatch();
        private bool hovering, pressed, isBusy, primary;
        private float hover, hoverStart, hoverTarget;
        private long hoverTime;
        private IconKind icon = IconKind.Arrow;
        private Color accentColor = AppTheme.Accent;
        public IconKind Icon { get { return icon; } set { icon = value; Invalidate(); } }
        public bool Primary { get { return primary; } set { primary = value; Invalidate(); } }
        public Color AccentColor { get { return accentColor; } set { accentColor = value; Invalidate(); } }
        public bool Busy
        {
            get { return isBusy; }
            set { if (isBusy == value) return; isBusy = value; pressed = false; UpdateAnimation(); Invalidate(); }
        }
        public ActionButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0; UseVisualStyleBackColor = false;
            BackColor = Color.Transparent; ForeColor = AppTheme.Ink;
            Cursor = Cursors.Hand; Size = new Size(156, 44); MinimumSize = new Size(40, 32);
            Font = new Font("Microsoft YaHei UI", 10F);
            animation = new Timer(); animation.Interval = 16;
            animation.Tick += delegate
            {
                float t = Math.Min(1F, (clock.ElapsedMilliseconds - hoverTime) / 145F);
                float eased = 1F - (1F - t) * (1F - t) * (1F - t);
                hover = hoverStart + (hoverTarget - hoverStart) * eased;
                if ((!isBusy || !Enabled) && t >= 1F) animation.Stop();
                if (!AppTheme.Animate) { hover = hoverTarget; animation.Stop(); }
                Invalidate();
            };
            clock.Start();
        }
        private void SetHover(bool value)
        {
            hovering = value; hoverStart = hover; hoverTarget = value ? 1 : 0; hoverTime = clock.ElapsedMilliseconds;
            UpdateAnimation(); Invalidate();
        }
        private void UpdateAnimation()
        {
            if (IsDisposed) return;
            if (AppTheme.Animate && Enabled && (isBusy || Math.Abs(hover - hoverTarget) > 0.01F)) animation.Start();
            else { hover = hoverTarget; animation.Stop(); }
        }
        protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); SetHover(true); }
        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); pressed = false; SetHover(false); }
        protected override void OnMouseDown(MouseEventArgs e) { if (!isBusy) { base.OnMouseDown(e); pressed = e.Button == MouseButtons.Left; Invalidate(); } }
        protected override void OnMouseUp(MouseEventArgs e) { pressed = false; if (!isBusy) base.OnMouseUp(e); Invalidate(); }
        protected override void OnMouseCaptureChanged(EventArgs e) { base.OnMouseCaptureChanged(e); pressed = false; Invalidate(); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (isBusy && (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)) { e.Handled = true; e.SuppressKeyPress = true; return; }
            if (e.KeyCode == Keys.Space) { pressed = true; Invalidate(); }
            base.OnKeyDown(e);
        }
        protected override void OnKeyUp(KeyEventArgs e) { pressed = false; Invalidate(); base.OnKeyUp(e); }
        protected override void OnClick(EventArgs e) { if (!isBusy && Enabled) base.OnClick(e); }
        protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); if (!Enabled) { pressed = false; SetHover(false); } UpdateAnimation(); Invalidate(); }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); pressed = false; Invalidate(); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            Color fill = primary ? AppTheme.Mix(accentColor, Color.FromArgb(13, 48, 112), hover * 0.20F) : AppTheme.Mix(Color.White, Color.FromArgb(232, 240, 253), hover);
            Color ink = primary ? Color.White : AppTheme.Ink;
            if (!Enabled) { fill = primary ? Color.FromArgb(170, 188, 215) : Color.FromArgb(244, 246, 249); ink = primary ? Color.White : Color.FromArgb(157, 168, 181); }
            RectangleF bounds = new RectangleF(1, 1 + (pressed ? 1 : 0), Math.Max(1, Width - 3), Math.Max(1, Height - 3 - (pressed ? 1 : 0)));
            using (GraphicsPath path = AppTheme.Round(bounds, 10))
            {
                using (Brush b = new SolidBrush(fill)) g.FillPath(b, path);
                if (!primary) using (Pen p = new Pen(hover > 0.1F ? AppTheme.Mix(AppTheme.Border, accentColor, hover * 0.4F) : AppTheme.Border)) g.DrawPath(p, path);
            }
            int iconSize = Math.Max(16, Math.Min(22, Height / 2));
            Size textSize = TextRenderer.MeasureText(Text, Font, new Size(Math.Max(1, Width - iconSize - 28), Height), TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);
            int fullWidth = iconSize + (String.IsNullOrEmpty(Text) ? 0 : 9 + Math.Min(textSize.Width, Math.Max(0, Width - iconSize - 28)));
            int left = Math.Max(10, (Width - fullWidth) / 2), top = (Height - iconSize) / 2 + (pressed ? 1 : 0);
            RectangleF iconBounds = new RectangleF(left, top, iconSize, iconSize);
            if (isBusy)
            {
                using (Pen faded = new Pen(Color.FromArgb(60, ink), 2)) g.DrawEllipse(faded, iconBounds);
                using (Pen p = new Pen(ink, 2)) { p.StartCap = LineCap.Round; p.EndCap = LineCap.Round; g.DrawArc(p, iconBounds, AppTheme.Animate ? (float)(clock.ElapsedMilliseconds % 1000) * 0.36F : -90F, 105F); }
            }
            else AppTheme.DrawIcon(g, icon, iconBounds, ink, 1.8F);
            if (!String.IsNullOrEmpty(Text)) TextRenderer.DrawText(g, Text, Font, new Rectangle(left + iconSize + 9, pressed ? 1 : 0, Math.Max(1, Width - left - iconSize - 18), Height), ink, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
            if (Focused && ShowFocusCues)
                using (GraphicsPath path = AppTheme.Round(new RectangleF(4, 4, Math.Max(1, Width - 9), Math.Max(1, Height - 9)), 7))
                using (Pen p = new Pen(primary ? Color.White : accentColor, 1.2F)) { p.DashStyle = DashStyle.Dot; g.DrawPath(p, path); }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { animation.Stop(); animation.Dispose(); clock.Stop(); }
            base.Dispose(disposing);
        }
    }

    public class StatusGlyph : Control
    {
        private readonly Timer animation;
        private readonly Stopwatch clock = new Stopwatch();
        private string state = "idle";
        private long changedAt;
        public string State
        {
            get { return state; }
            set
            {
                string next = String.IsNullOrEmpty(value) ? "idle" : value.ToLowerInvariant();
                if (state == next) return;
                state = next; changedAt = clock.ElapsedMilliseconds;
                AccessibleDescription = state == "working" ? "正在处理" : state == "success" ? "已完成" : state == "warning" ? "需要核实" : "准备就绪";
                if (AppTheme.Animate && Visible) animation.Start(); else animation.Stop();
                Invalidate();
            }
        }
        public StatusGlyph()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor | ControlStyles.ResizeRedraw, true);
            BackColor = Color.Transparent; Size = new Size(68, 68); TabStop = false;
            AccessibleRole = AccessibleRole.Graphic; AccessibleName = "采集状态";
            animation = new Timer(); animation.Interval = 30;
            animation.Tick += delegate { if (!AppTheme.Animate || (state != "working" && clock.ElapsedMilliseconds - changedAt > 340)) animation.Stop(); Invalidate(); };
            clock.Start();
        }
        protected override void OnVisibleChanged(EventArgs e) { base.OnVisibleChanged(e); if (Visible && state == "working" && AppTheme.Animate) animation.Start(); else animation.Stop(); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            Color color = state == "success" ? AppTheme.Success : state == "warning" ? Color.FromArgb(191, 124, 27) : AppTheme.Accent;
            float d = Math.Max(1, Math.Min(Width, Height) - 4), x = (Width - d) / 2F, y = (Height - d) / 2F;
            using (Brush b = new SolidBrush(AppTheme.Mix(Color.White, color, 0.08F))) g.FillEllipse(b, x, y, d, d);
            float t = AppTheme.Animate ? Math.Min(1, (clock.ElapsedMilliseconds - changedAt) / 320F) : 1;
            float ease = 1 - (float)Math.Pow(1 - t, 3);
            float iconSize = d * (0.42F + 0.07F * ease);
            RectangleF iconBounds = new RectangleF((Width - iconSize) / 2, (Height - iconSize) / 2, iconSize, iconSize);
            if (state == "working")
            {
                using (Pen p = new Pen(AppTheme.Mix(Color.White, color, 0.2F), 2.5F)) g.DrawEllipse(p, x + 2, y + 2, d - 4, d - 4);
                using (Pen p = new Pen(color, 2.5F)) { p.StartCap = LineCap.Round; p.EndCap = LineCap.Round; g.DrawArc(p, x + 2, y + 2, d - 4, d - 4, AppTheme.Animate ? (float)(clock.ElapsedMilliseconds % 1500) * 0.24F : -90, 110); }
                AppTheme.DrawIcon(g, IconKind.Computer, iconBounds, color, 2F);
            }
            else AppTheme.DrawIcon(g, state == "success" ? IconKind.Check : state == "warning" ? IconKind.Warning : IconKind.Computer, iconBounds, color, 2.1F);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { animation.Stop(); animation.Dispose(); clock.Stop(); }
            base.Dispose(disposing);
        }
    }

    public class CalendarField : UserControl
    {
        private string dateValue = "";
        private bool hover, clearHover;
        private ToolStripDropDown popup;
        private MonthCalendar calendar;
        private readonly ToolTip hints;
        public event EventHandler ValueChanged;
        public string DateValue
        {
            get { return dateValue; }
            set
            {
                string next = (value ?? "").Trim();
                if (next == dateValue) return;
                dateValue = next;
                AccessibleDescription = String.IsNullOrEmpty(next) ? "日期未填写。按回车选择日期，按Delete清空。" : "日期：" + next + "。按回车修改，按Delete清空。";
                Invalidate();
                EventHandler handler = ValueChanged; if (handler != null) handler(this, EventArgs.Empty);
            }
        }
        public CalendarField()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            BackColor = Color.White; ForeColor = AppTheme.Ink; Font = new Font("Microsoft YaHei UI", 10F);
            MinimumSize = new Size(150, 34); Size = new Size(250, 38); TabStop = true; Cursor = Cursors.Hand;
            AccessibleRole = AccessibleRole.DropList; AccessibleName = "取得日期";
            AccessibleDescription = "日期未填写。按回车选择日期，按Delete清空。";
            hints = new ToolTip(); hints.SetToolTip(this, "点击打开日历；按 Delete 清空。不清楚可以留空。");
        }
        private Rectangle ClearArea { get { return new Rectangle(Width - 55, 1, 27, Height - 2); } }
        protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); hover = true; Invalidate(); }
        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); hover = clearHover = false; Invalidate(); }
        protected override void OnMouseMove(MouseEventArgs e) { base.OnMouseMove(e); bool now = !String.IsNullOrEmpty(dateValue) && ClearArea.Contains(e.Location); if (now != clearHover) { clearHover = now; Invalidate(); } }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e); if (e.Button != MouseButtons.Left || !Enabled) return;
            Focus();
            if (!String.IsNullOrEmpty(dateValue) && ClearArea.Contains(e.Location)) { DateValue = ""; return; }
            OpenCalendar();
        }
        protected override bool IsInputKey(Keys keyData)
        {
            Keys key = keyData & Keys.KeyCode;
            return key == Keys.Enter || key == Keys.Space || key == Keys.Down || key == Keys.Delete || base.IsInputKey(keyData);
        }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space || e.KeyCode == Keys.Down)
            { OpenCalendar(); e.Handled = true; e.SuppressKeyPress = true; }
            else if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
            { DateValue = ""; e.Handled = true; e.SuppressKeyPress = true; }
            else if (e.KeyCode == Keys.Escape && popup != null && popup.Visible)
            { popup.Close(); e.Handled = true; e.SuppressKeyPress = true; }
            base.OnKeyDown(e);
        }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        private void OpenCalendar()
        {
            if (!Enabled) return;
            if (popup != null && popup.Visible) { popup.Close(); return; }
            if (popup == null)
            {
                calendar = new MonthCalendar();
                calendar.MaxSelectionCount = 1; calendar.ShowToday = true; calendar.ShowTodayCircle = true;
                calendar.Font = Font; calendar.Margin = Padding.Empty;
                calendar.DateSelected += delegate(object sender, DateRangeEventArgs e)
                { DateValue = e.Start.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); popup.Close(ToolStripDropDownCloseReason.ItemClicked); Focus(); };
                calendar.KeyDown += delegate(object sender, KeyEventArgs e)
                {
                    if (e.KeyCode == Keys.Escape) { popup.Close(); Focus(); e.Handled = true; e.SuppressKeyPress = true; }
                    else if (e.KeyCode == Keys.Enter)
                    { DateValue = calendar.SelectionStart.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); popup.Close(); Focus(); e.Handled = true; e.SuppressKeyPress = true; }
                };
                ToolStripControlHost host = new ToolStripControlHost(calendar);
                host.Margin = Padding.Empty; host.Padding = Padding.Empty;
                popup = new ToolStripDropDown(); popup.AutoClose = true; popup.Padding = new Padding(6); popup.BackColor = Color.White;
                popup.Items.Add(host); popup.Closed += delegate { Invalidate(); };
            }
            DateTime initial;
            string[] formats = new string[] { "yyyy-MM-dd", "yyyy/M/d", "yyyy-M-d", "yyyy-MM", "yyyy-M", "yyyy/MM", "yyyy/M", "yyyy" };
            if (!DateTime.TryParseExact(dateValue, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out initial)) initial = DateTime.Today;
            initial = initial < calendar.MinDate ? calendar.MinDate : initial > calendar.MaxDate ? calendar.MaxDate : initial;
            calendar.SetDate(initial);
            popup.Show(this, new Point(0, Height + 3));
            calendar.Focus(); Invalidate();
        }
        protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); if (!Enabled && popup != null) popup.Close(); Invalidate(); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            Color edge = Focused || (popup != null && popup.Visible) ? AppTheme.Accent : hover ? Color.FromArgb(168, 186, 209) : AppTheme.Border;
            using (GraphicsPath path = AppTheme.Round(new RectangleF(0.7F, 0.7F, Width - 1.5F, Height - 1.5F), 7))
            {
                using (Brush b = new SolidBrush(Enabled ? Color.White : Color.FromArgb(245, 247, 250))) g.FillPath(b, path);
                using (Pen p = new Pen(edge, Focused ? 1.5F : 1F)) g.DrawPath(p, path);
            }
            AppTheme.DrawIcon(g, IconKind.Calendar, new RectangleF(11, (Height - 18) / 2F, 18, 18), Enabled ? AppTheme.Muted : Color.Gray, 1.6F);
            bool empty = String.IsNullOrEmpty(dateValue);
            int right = empty ? 29 : 56;
            TextRenderer.DrawText(g, empty ? "未填写，点日历选择" : dateValue, Font, new Rectangle(38, 0, Math.Max(1, Width - 38 - right), Height), empty || !Enabled ? AppTheme.Muted : AppTheme.Ink, TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            if (!empty)
            {
                Rectangle r = ClearArea;
                if (clearHover && Enabled) using (Brush b = new SolidBrush(Color.FromArgb(233, 238, 245))) g.FillEllipse(b, r.X + 2, (Height - 23) / 2F, 23, 23);
                using (Pen p = new Pen(AppTheme.Muted, 1.5F)) { float cx = r.X + r.Width / 2F, cy = Height / 2F; g.DrawLine(p, cx - 3, cy - 3, cx + 3, cy + 3); g.DrawLine(p, cx + 3, cy - 3, cx - 3, cy + 3); }
            }
            AppTheme.DrawIcon(g, IconKind.Chevron, new RectangleF(Width - 23, (Height - 14) / 2F, 14, 14), AppTheme.Muted, 1.4F);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { if (popup != null) popup.Dispose(); hints.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
