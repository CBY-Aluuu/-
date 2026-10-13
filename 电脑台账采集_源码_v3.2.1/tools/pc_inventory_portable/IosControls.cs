using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using Calendar = System.Windows.Controls.Calendar;

namespace PcInventory.Ios
{
    /// <summary>A blank-by-default date field. Partial dates stay unchanged until a day is chosen.</summary>
    public sealed class IosDateField : UserControl
    {
        private readonly Button openButton;
        private readonly Button clearButton;
        private readonly TextBlock label;
        private readonly Popup popup;
        private readonly Calendar calendar;
        private string dateValue = "";

        public event EventHandler ValueChanged;

        public string DateValue
        {
            get { return dateValue; }
            set
            {
                string next = value ?? "";
                if (next == dateValue) return;
                dateValue = next;
                RefreshLabel();
                var handler = ValueChanged;
                if (handler != null) handler(this, EventArgs.Empty);
            }
        }

        public IosDateField()
        {
            Height = 46;
            MinWidth = 0;
            Focusable = false;
            KeyboardNavigation.SetTabNavigation(this, KeyboardNavigationMode.Local);
            var root = new Grid { Background = Brushes.Transparent };
            openButton = new Button
            {
                Height = 46,
                Background = Brush("#F5F5F7"),
                BorderBrush = Brush("#E5E5EA"),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(14, 0, 14, 0),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                ToolTip = "点击日历选择日期；不清楚就留空"
            };
            System.Windows.Automation.AutomationProperties.SetName(openButton, "取得日期，打开日历");
            var buttonContent = new Grid();
            buttonContent.ColumnDefinitions.Add(new ColumnDefinition());
            buttonContent.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(19) });
            label = new TextBlock { VerticalAlignment = VerticalAlignment.Center, FontSize = 13, TextTrimming = TextTrimming.CharacterEllipsis };
            buttonContent.Children.Add(label);
            var icon = new Path
            {
                Data = Geometry.Parse("M 3,4 L 17,4 Q 18,4 18,5 L 18,17 Q 18,18 17,18 L 3,18 Q 2,18 2,17 L 2,5 Q 2,4 3,4 M 2,8 L 18,8 M 6,2 L 6,6 M 14,2 L 14,6 M 6,12 L 8,12 M 12,12 L 14,12"),
                Width = 18,
                Height = 18,
                Stretch = Stretch.Uniform,
                Stroke = Brush("#8E8E93"),
                StrokeThickness = 1.4,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(icon, 1);
            buttonContent.Children.Add(icon);
            openButton.Content = buttonContent;
            openButton.Click += delegate { ToggleCalendar(); };
            openButton.PreviewKeyDown += OnOpenKeyDown;
            root.Children.Add(openButton);

            clearButton = new Button
            {
                Content = "×",
                Width = 28,
                Height = 28,
                Padding = new Thickness(0),
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
                FontSize = 20,
                Background = Brush("#E9E9EF"),
                Foreground = Brush("#86868B"),
                BorderThickness = new Thickness(0),
                ToolTip = "清空日期"
            };
            System.Windows.Automation.AutomationProperties.SetName(clearButton, "清空取得日期");
            clearButton.Click += delegate { CloseCalendar(false); DateValue = ""; openButton.Focus(); };
            root.Children.Add(clearButton);

            calendar = new Calendar
            {
                SelectionMode = CalendarSelectionMode.SingleDate,
                Background = Brushes.White,
                BorderThickness = new Thickness(0),
                FontFamily = new FontFamily("Microsoft YaHei UI, Segoe UI"),
                FontSize = 13,
                Language = System.Windows.Markup.XmlLanguage.GetLanguage("zh-CN")
            };
            calendar.SetResourceReference(Calendar.CalendarDayButtonStyleProperty, "IosCalendarDay");
            calendar.PreviewMouseLeftButtonUp += OnCalendarMouseUp;
            calendar.PreviewKeyDown += OnCalendarKeyDown;
            var popupContent = new Border
            {
                Background = Brushes.White,
                BorderBrush = Brush("#E5E5EA"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(16),
                Padding = new Thickness(10),
                Margin = new Thickness(10),
                Child = calendar,
                Effect = new DropShadowEffect { BlurRadius = 18, ShadowDepth = 3, Color = Colors.Black, Opacity = 0.10 }
            };
            popup = new Popup
            {
                PlacementTarget = openButton,
                Placement = PlacementMode.Bottom,
                HorizontalOffset = -10,
                VerticalOffset = -4,
                AllowsTransparency = true,
                StaysOpen = false,
                PopupAnimation = SystemParameters.ClientAreaAnimation ? PopupAnimation.Fade : PopupAnimation.None,
                Child = popupContent
            };
            root.Children.Add(popup);
            popup.Opened += delegate { calendar.Focus(); };
            popup.Closed += delegate { ReleaseCalendarCapture(); };
            Unloaded += delegate { popup.IsOpen = false; };
            IsEnabledChanged += delegate { if (!IsEnabled) CloseCalendar(false); };
            Content = root;
            RefreshLabel();
        }

        private void RefreshLabel()
        {
            bool hasDate = dateValue.Length > 0;
            label.Text = hasDate ? dateValue : "选择日期 · 可留空";
            label.Foreground = Brush(hasDate ? "#1D1D1F" : "#8E8E93");
            clearButton.Visibility = hasDate ? Visibility.Visible : Visibility.Collapsed;
            openButton.Padding = new Thickness(14, 0, hasDate ? 46 : 14, 0);
            System.Windows.Automation.AutomationProperties.SetName(openButton, hasDate ? "取得日期 " + dateValue + "，打开日历" : "取得日期未填写，打开日历");
        }

        private void ToggleCalendar()
        {
            if (!IsEnabled) return;
            if (popup.IsOpen) { CloseCalendar(true); return; }
            DateTime parsed;
            bool exact = DateTime.TryParseExact(dateValue, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed);
            if (!exact && !DateTime.TryParseExact(dateValue, new[] { "yyyy-MM", "yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed)) parsed = DateTime.Today;
            calendar.DisplayMode = CalendarMode.Month;
            calendar.DisplayDate = parsed;
            calendar.SelectedDate = exact ? (DateTime?)parsed : null;
            popup.PopupAnimation = SystemParameters.ClientAreaAnimation ? PopupAnimation.Fade : PopupAnimation.None;
            popup.IsOpen = true;
        }

        private void OnOpenKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down || e.Key == Key.Enter || e.Key == Key.Space)
            {
                ToggleCalendar();
                e.Handled = true;
            }
            else if (e.Key == Key.Delete || e.Key == Key.Back)
            {
                DateValue = "";
                e.Handled = true;
            }
            else if (e.Key == Key.Escape && popup.IsOpen)
            {
                CloseCalendar(true);
                e.Handled = true;
            }
        }

        private void OnCalendarMouseUp(object sender, MouseButtonEventArgs e)
        {
            DependencyObject node = e.OriginalSource as DependencyObject;
            while (node != null && !(node is CalendarDayButton))
            {
                if (node is Visual || node is System.Windows.Media.Media3D.Visual3D) node = VisualTreeHelper.GetParent(node);
                else node = LogicalTreeHelper.GetParent(node);
            }
            var day = node as CalendarDayButton;
            if (day != null && !day.IsBlackedOut && day.IsEnabled && day.DataContext is DateTime)
            {
                Commit((DateTime)day.DataContext);
                e.Handled = true;
            }
        }

        private void OnCalendarKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) { CloseCalendar(true); e.Handled = true; }
            else if ((e.Key == Key.Enter || e.Key == Key.Space) && calendar.DisplayMode == CalendarMode.Month)
            {
                Commit(calendar.SelectedDate ?? calendar.DisplayDate);
                e.Handled = true;
            }
            else if (e.Key == Key.Tab)
            {
                CloseCalendar(true);
                openButton.MoveFocus(new TraversalRequest((Keyboard.Modifiers & ModifierKeys.Shift) != 0 ? FocusNavigationDirection.Previous : FocusNavigationDirection.Next));
                e.Handled = true;
            }
        }

        private void Commit(DateTime date)
        {
            DateValue = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            CloseCalendar(true);
        }

        private void CloseCalendar(bool returnFocus)
        {
            ReleaseCalendarCapture();
            popup.IsOpen = false;
            if (returnFocus) openButton.Focus();
        }

        private void ReleaseCalendarCapture()
        {
            // WPF CalendarItem can own capture after selecting a day. Release only this
            // calendar's capture so the very next click can reach another input field.
            DependencyObject node = Mouse.Captured as DependencyObject;
            while (node != null)
            {
                if (node == calendar) { Mouse.Capture(null); return; }
                if (node is Visual || node is System.Windows.Media.Media3D.Visual3D) node = VisualTreeHelper.GetParent(node);
                else node = LogicalTreeHelper.GetParent(node);
            }
        }

        private static SolidColorBrush Brush(string value)
        {
            var result = (SolidColorBrush)new BrushConverter().ConvertFromInvariantString(value);
            result.Freeze();
            return result;
        }
    }
}
