using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Animation;
using KioskPresenter.Host.Interop;

namespace KioskPresenter.Host;

/// <summary>
/// The native Home / Previous / counter / Next pill that floats over PowerPoint's slideshow.
/// It never takes focus (WS_EX_NOACTIVATE), so arrow keys and swipes keep reaching PowerPoint.
/// Hidden means fully transparent, which also makes it click-through.
/// </summary>
public partial class ControlStrip : Window
{
    private const double BottomMargin = 32;
    private static readonly Duration Fade = new(TimeSpan.FromMilliseconds(200));

    public ControlStrip()
    {
        InitializeComponent();
        Opacity = 0;
        IsHitTestVisible = false;
        SourceInitialized += (_, _) => WindowHelper.MakeNoActivate(new WindowInteropHelper(this).Handle);
        SizeChanged += (_, _) => Reposition();
    }

    public event Action? HomeRequested;
    public event Action? PreviousRequested;
    public event Action? NextRequested;

    public bool IsShowing { get; private set; }

    public void SetCounter(int position, int total)
    {
        Counter.Text = $"{position} / {total}";
        PrevButton.IsEnabled = position > 1;
    }

    public void ShowStrip()
    {
        if (!IsVisible) Show();
        Reposition();
        WindowHelper.SetTopmost(new WindowInteropHelper(this).Handle, true);
        if (IsShowing) return;
        IsShowing = true;
        IsHitTestVisible = true;
        BeginAnimation(OpacityProperty, new DoubleAnimation(1, Fade));
    }

    public void HideStrip()
    {
        if (!IsShowing) return;
        IsShowing = false;
        IsHitTestVisible = false;
        BeginAnimation(OpacityProperty, new DoubleAnimation(0, Fade));
    }

    /// <summary>Removes the strip entirely (between shows).</summary>
    public void Retire()
    {
        IsShowing = false;
        IsHitTestVisible = false;
        BeginAnimation(OpacityProperty, null);
        Opacity = 0;
        if (IsVisible) Hide();
    }

    private void Reposition()
    {
        var area = SystemParameters.WorkArea;
        var screenHeight = SystemParameters.PrimaryScreenHeight;
        Left = (SystemParameters.PrimaryScreenWidth - ActualWidth) / 2;
        Top = Math.Max(area.Top, screenHeight - ActualHeight - BottomMargin);
    }

    private void Home_Click(object sender, RoutedEventArgs e) => HomeRequested?.Invoke();

    private void Prev_Click(object sender, RoutedEventArgs e) => PreviousRequested?.Invoke();

    private void Next_Click(object sender, RoutedEventArgs e) => NextRequested?.Invoke();
}
