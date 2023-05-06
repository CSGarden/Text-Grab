using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Text_Grab.Properties;
using Text_Grab.Utilities;
using Windows.Globalization;
using Windows.Media.Ocr;

namespace Text_Grab.Views;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class FullscreenGrab : Window {
    #region Fields

    private System.Windows.Point clickedPoint = new System.Windows.Point();
    private TextBox? destinationTextBox;
    private DpiScale? dpiScale;
    private bool isComboBoxReady = false;
    private bool isSelecting = false;
    private bool isShiftDown = false;
    private Border selectBorder = new Border();
    double selectLeft;
    double selectTop;
    private System.Windows.Point shiftPoint = new System.Windows.Point();
    double xShiftDelta;
    double yShiftDelta;

    #endregion Fields

    #region Constructors

    public FullscreenGrab() {
        InitializeComponent();
        App.SetTheme();
    }

    #endregion Constructors

    #region Properties

    public TextBox? DestinationTextBox {
        get { return destinationTextBox; }
        set {
            destinationTextBox = value;
            if (destinationTextBox != null)
                SendToEditTextToggleButton.IsChecked = true;
            else
                SendToEditTextToggleButton.IsChecked = false;
        }
    }

    public bool IsFreeze { get; set; } = false;
    public string? textFromOCR { get; set; }
    private System.Windows.Forms.Screen? currentScreen { get; set; }

    #endregion Properties

    #region Methods
    /// <summary>
    /// 将 BackgroundImage 对象的源设置为 null，以清除之前可能存在的图片。
    ///调用 ImageMethods.GetWindowBoundsImage 方法获取 FullscreenGrab 窗口覆盖范围内的屏幕截图，并将其设置为 BackgroundImage 的源。
    ///将 BackgroundBrush 对象的 Opacity 属性设置为 0.2，使其透明度变为 20%，用于在悬浮识别区域之上形成半透明遮罩。
    /// </summary>
    public void SetImageToBackground() {
        BackgroundImage.Source = null;
        BackgroundImage.Source = ImageMethods.GetWindowBoundsImage(this);
        BackgroundBrush.Opacity = 0.2;
    }

    internal void KeyPressed(Key key, bool? isActive = null) {
        switch (key) {
            // This case is handled in the WindowUtilities.FullscreenKeyDown
            // case Key.Escape:
            //     WindowUtilities.CloseAllFullscreenGrabs();
            //     break;
            case Key.G:
                if (isActive is null)
                    NewGrabFrameMenuItem.IsChecked = !NewGrabFrameMenuItem.IsChecked;
                else
                    NewGrabFrameMenuItem.IsChecked = isActive.Value;
                break;
            case Key.S:
                if (isActive is null)
                    SingleLineMenuItem.IsChecked = !SingleLineMenuItem.IsChecked;
                else
                    SingleLineMenuItem.IsChecked = isActive.Value;

                Settings.Default.FSGMakeSingleLineToggle = SingleLineMenuItem.IsChecked;
                Settings.Default.Save();
                break;
            case Key.E:
                if (isActive is null)
                    SendToEtwMenuItem.IsChecked = !SendToEtwMenuItem.IsChecked;
                else
                    SendToEtwMenuItem.IsChecked = isActive.Value;
                break;
            case Key.F:
                if (isActive is null)
                    FreezeMenuItem.IsChecked = !FreezeMenuItem.IsChecked;
                else
                    FreezeMenuItem.IsChecked = isActive.Value;

                FreezeUnfreeze(FreezeMenuItem.IsChecked);
                break;
            case Key.T:
                if (isActive is null)
                    TableToggleButton.IsChecked = !TableToggleButton.IsChecked;
                else
                    TableToggleButton.IsChecked = isActive.Value;
                break;
            case Key.D1:
            case Key.D2:
            case Key.D3:
            case Key.D4:
            case Key.D5:
            case Key.D6:
            case Key.D7:
            case Key.D8:
            case Key.D9:
                int numberPressed = (int)key - 34; // D1 casts to 35, D2 to 36, etc.
                int numberOfLanguages = LanguagesComboBox.Items.Count;

                if (numberPressed <= numberOfLanguages
                    && numberPressed - 1 >= 0
                    && numberPressed - 1 != LanguagesComboBox.SelectedIndex
                    && isComboBoxReady)
                    LanguagesComboBox.SelectedIndex = numberPressed - 1;
                break;
            default:
                break;
        }
    }
    /// <summary>
    /// 这是一个C#方法，名为`CheckIfCheckingOrUnchecking`，它接受一个类型为`object`的参数并返回一个`bool`值。、
    ///这个方法的目的是确定一个切换按钮或菜单项是否被选中或取消选中。方法首先将一个名为`isActive`的布尔变量初始化为`false`。、
    ///然后它检查 `sender` 参数是否是一个 `ToggleButton` 并且它的 `IsChecked` 属性不为 null。如果两个条件都为真，则该方法将 `isActive` 设置为 `IsChecked` 属性的值。
    ///如果 `sender` 参数不是 `ToggleButton` 或它的 `IsChecked` 属性为 null，则该方法检查 `sender` 参数是否为 `MenuItem`，并将 `isActive` 设置为其 `IsChecked` 属性的值。
    ///最后，该方法返回 `isActive` 的值，如果切换按钮或菜单项被选中，则为 `true`，否则为 `false`。
    /// </summary>
    /// <param name="sender"></param>
    /// <returns></returns>
    private static bool CheckIfCheckingOrUnchecking(object? sender) {
        bool isActive = false;
        if (sender is ToggleButton tb && tb.IsChecked is not null)
            isActive = tb.IsChecked.Value;
        else if (sender is MenuItem mi)
            isActive = mi.IsChecked;
        return isActive;
    }

    private void CancelMenuItem_Click(object sender, RoutedEventArgs e) {
        WindowUtilities.CloseAllFullscreenGrabs();
    }
    /// <summary>
    /// 当触发这个事件的时候通过检查 CheckIfCheckingOrUnchecking 方法返回的结果来确定 isActive 变量的值，可能是 true 或 false。
    /// 然后，该方法使用 WindowUtilities.FullscreenKeyDown 方法将按键事件发送到全屏抓取窗口对象中以执行对应的操作，即模拟在键盘上按下 F 键。
    /// 需要注意的是，sender 和 e 参数都是可选的，如果不传入则为 null。
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void FreezeMenuItem_Click(object? sender = null, RoutedEventArgs? e = null) {
        bool isActive = CheckIfCheckingOrUnchecking(sender);

        WindowUtilities.FullscreenKeyDown(Key.F, isActive);
    }

    private async void FreezeUnfreeze(bool Activate) {
        if (FreezeMenuItem.IsChecked is true) {
            TopButtonsStackPanel.Visibility = Visibility.Collapsed;
            BackgroundBrush.Opacity = 0;
            RegionClickCanvas.ContextMenu.IsOpen = false;
            await Task.Delay(150);
            SetImageToBackground();
            TopButtonsStackPanel.Visibility = Visibility.Visible;
        } else {
            BackgroundImage.Source = null;
        }
    }

    private void FullscreenGrab_KeyDown(object sender, KeyEventArgs e) {
        WindowUtilities.FullscreenKeyDown(e.Key);
    }

    private void FullscreenGrab_KeyUp(object sender, KeyEventArgs e) {
        switch (e.Key) {
            case Key.LeftShift:
            case Key.RightShift:
                isShiftDown = false;
                clickedPoint = new System.Windows.Point(clickedPoint.X + xShiftDelta, clickedPoint.Y + yShiftDelta);
                break;
            default:
                break;
        }
    }
    /// <summary>
    /// 用于对 OCR 引擎返回的识别结果进行处理<br/>
    /// </summary>
    /// <param name="grabbedText"></param>
    private void HandleTextFromOcr(string grabbedText) {
        //根据界面设置中“单行”和“表格模式”选择状态的不同，选定是否需要将返回的多行文本合并为一行。
        if (SingleLineMenuItem.IsChecked is true && TableToggleButton.IsChecked is false)
            grabbedText = grabbedText.MakeStringSingleLine();

        textFromOCR = grabbedText;
        //果程序设置中禁用了自动将内容复制到系统剪贴板的选项（即 NeverAutoUseClipboard 为 false），
        //并且 destinationTextBox 不处于活动状态，则调用 Clipboard.SetDataObject 方法将文本复制到系统剪贴板中。
        if (!Settings.Default.NeverAutoUseClipboard && destinationTextBox is null)
            try { Clipboard.SetDataObject(grabbedText, true); } catch { }
        //如果启用了显示通知的功能（ShowToast 为 true），并且 destinationTextBox 不处于活动状态，
        //则调用 NotificationUtilities.ShowToast 显示一个通知提示框，将识别出的结果作为参数。
        if (Settings.Default.ShowToast && destinationTextBox is null)
            NotificationUtilities.ShowToast(grabbedText);
        //如果 destinationTextBox 存在，则在其中插入被识别出的文本。具体地，在指定的位置插入文本，然后将焦点设置为 destinationTextBox。
        if (destinationTextBox is not null) {
            // Do it this way instead of append text because it inserts the text at the cursor
            // Then puts the cursor at the end of the newly added text
            // AppendText() just adds the text to the end no matter what.
            destinationTextBox.SelectedText = grabbedText;
            destinationTextBox.Select(destinationTextBox.SelectionStart + grabbedText.Length, 0);
            destinationTextBox.Focus();
        }
    }
    /// <summary>
    /// <see cref="PreviewMouseDown"/> PreviewMouseDown 事件主要用于（比如点击、鼠标右键点击、鼠标滚轮等），系统会判断事件参数 e 中的鼠标中键是否按下。
    /// 如果鼠标中键确实按下，系统将会清空保存在应用程序设置中的 LastUsedLang 属性，这个属性可能用于记录应用程序中最近使用的语言。
    /// 由于 Settings.Default 是可序列化的，系统会将更改保存到设置文件中。
    /// 需要注意的是，该代码中语句的正确执行条件是事件参数 e 中需要包含鼠标中键的状态信息，如果该信息丢失可能会导致代码无法正确执行。
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void LanguagesComboBox_PreviewMouseDown(object sender, MouseButtonEventArgs e) {
        if (e.MiddleButton == MouseButtonState.Pressed) {
            Settings.Default.LastUsedLang = String.Empty;
            Settings.Default.Save();
        }
    }

    private void LanguagesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
        if (sender is not ComboBox languageCmbBox || !isComboBoxReady)
            return;

        Language? pickedLang = languageCmbBox.SelectedItem as Language;

        if (pickedLang != null) {
            Settings.Default.LastUsedLang = pickedLang.LanguageTag;
            Settings.Default.Save();
        }

        int selection = languageCmbBox.SelectedIndex;

        switch (selection) {
            case 0:
                WindowUtilities.FullscreenKeyDown(Key.D1);
                break;
            case 1:
                WindowUtilities.FullscreenKeyDown(Key.D2);
                break;
            case 2:
                WindowUtilities.FullscreenKeyDown(Key.D3);
                break;
            case 3:
                WindowUtilities.FullscreenKeyDown(Key.D4);
                break;
            case 4:
                WindowUtilities.FullscreenKeyDown(Key.D5);
                break;
            case 5:
                WindowUtilities.FullscreenKeyDown(Key.D6);
                break;
            case 6:
                WindowUtilities.FullscreenKeyDown(Key.D7);
                break;
            case 7:
                WindowUtilities.FullscreenKeyDown(Key.D8);
                break;
            case 8:
                WindowUtilities.FullscreenKeyDown(Key.D9);
                break;
            default:
                break;
        }
    }
    /// <summary>
    /// ，LoadOcrLanguages 方法是用于加载和创建 FullscreenGrab 窗口中 OCR 识别语言的下拉框列表的工具方法，并基于用户之前的选择对该下拉框进行一些设置。
    /// </summary>
    private void LoadOcrLanguages() {
        if (LanguagesComboBox.Items.Count > 0)
            return;

        IReadOnlyList<Language> possibleOCRLangs = OcrEngine.AvailableRecognizerLanguages;
        Language firstLang = LanguageUtilities.GetOCRLanguage();

        int count = 0;

        foreach (Language language in possibleOCRLangs) {
            LanguagesComboBox.Items.Add(language);

            if (language.LanguageTag == firstLang?.LanguageTag)
                LanguagesComboBox.SelectedIndex = count;

            count++;
        }

        isComboBoxReady = true;
    }

    private void NewEditTextMenuItem_Click(object sender, RoutedEventArgs e) {
        WindowUtilities.OpenOrActivateWindow<EditTextWindow>();
        WindowUtilities.CloseAllFullscreenGrabs();
    }

    private void NewGrabFrameMenuItem_Click(object sender, RoutedEventArgs e) {
        bool isActive = CheckIfCheckingOrUnchecking(sender);
        WindowUtilities.FullscreenKeyDown(Key.G, isActive);
    }

    private void PanSelection(System.Windows.Point movingPoint) {
        if (!isShiftDown) {
            shiftPoint = movingPoint;
            selectLeft = Canvas.GetLeft(selectBorder);
            selectTop = Canvas.GetTop(selectBorder);
        }

        isShiftDown = true;
        xShiftDelta = (movingPoint.X - shiftPoint.X);
        yShiftDelta = (movingPoint.Y - shiftPoint.Y);

        double leftValue = selectLeft + xShiftDelta;
        double topValue = selectTop + yShiftDelta;

        if (currentScreen is not null && dpiScale is not null) {
            double currentScreenLeft = currentScreen.Bounds.Left; // Should always be 0
            double currentScreenRight = currentScreen.Bounds.Right / dpiScale.Value.DpiScaleX;
            double currentScreenTop = currentScreen.Bounds.Top; // Should always be 0
            double currentScreenBottom = currentScreen.Bounds.Bottom / dpiScale.Value.DpiScaleY;

            leftValue = Math.Clamp(leftValue, currentScreenLeft, (currentScreenRight - selectBorder.Width));
            topValue = Math.Clamp(topValue, currentScreenTop, (currentScreenBottom - selectBorder.Height));
        }

        clippingGeometry.Rect = new Rect(
            new System.Windows.Point(leftValue, topValue),
            new System.Windows.Size(selectBorder.Width - 2, selectBorder.Height - 2));
        Canvas.SetLeft(selectBorder, leftValue - 1);
        Canvas.SetTop(selectBorder, topValue - 1);
    }

    private void PlaceGrabFrameInSelectionRect() {
        // Make a new GrabFrame and show it on screen
        // Then place it where the user just drew the region
        // Add space around the window to account for titlebar
        // bottom bar and width of GrabFrame
        System.Windows.Point absPosPoint = this.GetAbsolutePosition();
        DpiScale dpi = VisualTreeHelper.GetDpi(this);
        int firstScreenBPP = System.Windows.Forms.Screen.AllScreens[0].BitsPerPixel;
        GrabFrame grabFrame = new();
        if (destinationTextBox is not null)
            grabFrame.DestinationTextBox = destinationTextBox;

        grabFrame.TableToggleButton.IsChecked = TableToggleButton.IsChecked;
        grabFrame.Show();

        double posLeft = Canvas.GetLeft(selectBorder); // * dpi.DpiScaleX;
        double posTop = Canvas.GetTop(selectBorder); // * dpi.DpiScaleY;
        grabFrame.Left = posLeft + (absPosPoint.X / dpi.PixelsPerDip);
        grabFrame.Top = posTop + (absPosPoint.Y / dpi.PixelsPerDip);

        grabFrame.Left -= (2 / dpi.PixelsPerDip);
        grabFrame.Top -= (34 / dpi.PixelsPerDip);
        // if (grabFrame.Top < 0)
        //     grabFrame.Top = 0;

        if (selectBorder.Width > 20 && selectBorder.Height > 20) {
            grabFrame.Width = selectBorder.Width + 4;
            grabFrame.Height = selectBorder.Height + 72;
        }
        grabFrame.Activate();
        WindowUtilities.CloseAllFullscreenGrabs();
    }
    /// <summary>
    /// 用户按下鼠标左键选择屏幕区域时处理相应的事件<br/>
    /// 检查是否按下了右键，如果按下了则直接返回。<br/>
    /// 将 isSelecting 标志设置为 true，表示当前正在进行区域选择操作。<br/>
    ///隐藏顶部按钮栏 TopButtonsStackPanel，以便更好地进行选区。<br/>
    ///调用 RegionClickCanvas.CaptureMouse() 方法将鼠标指针捕获到画布上，确保画布接收所有输入事件。<br/>
    ///通过 CursorClipper.ClipCursor(this) 将光标限制在选取区域内。<br/>
    ///获取鼠标点击的点 clickedPoint，并将 selectBorder 初始化为一个空白的矩形。<br/>
    ///获取当前屏幕的 DPI 缩放比例 dpiScale。<br/>
    ///尝试从 RegionClickCanvas 中移除之前的 selectBorder，即清除之前可能存在的选区边框。<br/>
    ///设置 selectBorder 的属性，包括边框大小、颜色和位置等。<br/>
    ///将新的 selectBorder 添加到 RegionClickCanvas 中。<br/>
    ///通过 Canvas.SetLeft 和 Canvas.SetTop 方法将 selectBorder 移到鼠标点击的点。<br/>
    ///根据鼠标所在的点，在所有显示器中确定当前显示器的范围 currentScreen<br/>
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void RegionClickCanvas_MouseDown(object sender, MouseButtonEventArgs e) {
        if (e.RightButton == MouseButtonState.Pressed)
            return;

        isSelecting = true;
        TopButtonsStackPanel.Visibility = Visibility.Collapsed;
        RegionClickCanvas.CaptureMouse();
        CursorClipper.ClipCursor(this);
        clickedPoint = e.GetPosition(this);
        selectBorder.Height = 1;
        selectBorder.Width = 1;

        dpiScale = VisualTreeHelper.GetDpi(this);

        try { RegionClickCanvas.Children.Remove(selectBorder); } catch (Exception) { }

        selectBorder.BorderThickness = new Thickness(2);
        System.Windows.Media.Color borderColor = System.Windows.Media.Color.FromArgb(255, 40, 118, 126);
        selectBorder.BorderBrush = new SolidColorBrush(borderColor);
        _ = RegionClickCanvas.Children.Add(selectBorder);
        Canvas.SetLeft(selectBorder, clickedPoint.X);
        Canvas.SetTop(selectBorder, clickedPoint.Y);

        var screens = System.Windows.Forms.Screen.AllScreens;
        System.Drawing.Point formsPoint = new((int)clickedPoint.X, (int)clickedPoint.Y);
        foreach (var scr in screens)
            if (scr.Bounds.Contains(formsPoint))
                currentScreen = scr;
    }

    /// <summary>
    /// 用于实现鼠标移动时更新选择框的大小和位置<br/>
    /// 判断是否正在进行选区操作，如果否则返回。<br/>
    /// 获取当前鼠标所在的坐标点 movingPoint。<br/>
    ///如果按下了 Shift 键，则调用 PanSelection 方法进行平移并返回。<br/>
    ///否则将 isShiftDown 标志设置为 false。<br/>
    ///计算选中区域的左上角点(left, top)，同时计算出新选择框的宽度和高度。<br/>
    ///更新选择框 selectBorder 的宽度和高度，使其与鼠标移动后的选区大小相等，并在高度和宽度上各增加 2 个像素以确保选择框不重叠。<br/>
    ///使用剪裁几何 clippingGeometry 限制遮罩层的矩形范围，即保证截屏操作只在选定的区域内部进行。<br/>
    ///调整 selectBorder 框架的位置，将其移动到左上角点位置。<br/>
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void RegionClickCanvas_MouseMove(object sender, MouseEventArgs e) {
        if (!isSelecting)
            return;

        System.Windows.Point movingPoint = e.GetPosition(this);

        if (Keyboard.Modifiers == ModifierKeys.Shift) {
            PanSelection(movingPoint);
            return;
        }

        isShiftDown = false;

        var left = Math.Min(clickedPoint.X, movingPoint.X);
        var top = Math.Min(clickedPoint.Y, movingPoint.Y);

        selectBorder.Height = Math.Max(clickedPoint.Y, movingPoint.Y) - top;
        selectBorder.Width = Math.Max(clickedPoint.X, movingPoint.X) - left;
        selectBorder.Height = selectBorder.Height + 2;
        selectBorder.Width = selectBorder.Width + 2;

        clippingGeometry.Rect = new Rect(
            new System.Windows.Point(left, top),
            new System.Windows.Size(selectBorder.Width - 2, selectBorder.Height - 2));
        Canvas.SetLeft(selectBorder, left - 1);
        Canvas.SetTop(selectBorder, top - 1);
    }

    /// <summary>
    /// 用于在用户进行完选区操作后进行一些处理<br/>
    /// 如果当前没有进行选区操作，则直接返回。<br/>
    ///将 isSelecting 标志设置为 false，表示选区操作已结束。<br/>
    ///清空 currentScreen 变量，重置一些用于截屏的相关参数和属性。<br/>
    ///显示顶部按钮栏 TopButtonsStackPanel。<br/>
    ///通过 CursorClipper.UnClipCursor() 方法解除对鼠标光标的限制。<br/>
    ///释放对画布鼠标事件的捕获。<br/>
    ///将剪裁几何 clippingGeometry 的矩形范围设置为（0,0）。<br/>
    ///获取鼠标抬起的点 movingPoint，并根据当前屏幕的 DPI 缩放比例转换其坐标。<br/>
    ///对移动坐标进行四舍五入处理。<br/>
    ///计算窗口最左侧、最上方的可见位置 correctedLeft 和 correctedTop，并分别赋给变量。<br/>
    ///计算经过缩放后选择框左侧和上侧的坐标 xDimScaled 和 yDimScaled。<br/>
    ///使用 regionScaled 实例存储经过缩放后的选择框区域信息，包括其左侧坐标、上侧坐标、宽度和高度。<br/>
    ///如果 "NewGrabFrameMenuItem" 被选中，则调用 PlaceGrabFrameInSelectionRect 方法并返回。<br/>
    /// 尝试从 RegionClickCanvas 中移除之前创建的 selectBorder 选择框。<br/>
    /// 获取用户选择的语言类型，如果为空，则使用系统默认的OCR语言。<br/>
    /// 根据选择的操作类型，调用相应的方法。当选择框宽度和高度均小于 3 像素时，截取单个词语。否则根据 "TableToggleButton" 是否被选中决定是要作为文本表格截屏还是普通截屏。<br/>
    ///处理 OCR 返回的文本，将其传送到 EditTextWindow 窗口或其他窗口，并在需要时进行全屏截屏处理。<br/>
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private async void RegionClickCanvas_MouseUp(object sender, MouseButtonEventArgs e) {
        if (!isSelecting)
            return;

        isSelecting = false;
        currentScreen = null;
        TopButtonsStackPanel.Visibility = Visibility.Visible;
        CursorClipper.UnClipCursor();
        RegionClickCanvas.ReleaseMouseCapture();
        clippingGeometry.Rect = new Rect(
            new System.Windows.Point(0, 0),
            new System.Windows.Size(0, 0));

        System.Windows.Point movingPoint = e.GetPosition(this);
        Matrix m = PresentationSource.FromVisual(this).CompositionTarget.TransformToDevice;
        movingPoint.X *= m.M11;
        movingPoint.Y *= m.M22;

        movingPoint.X = Math.Round(movingPoint.X);
        movingPoint.Y = Math.Round(movingPoint.Y);

        double correctedLeft = Left;
        double correctedTop = Top;

        if (correctedLeft < 0)
            correctedLeft = 0;

        if (correctedTop < 0)
            correctedTop = 0;

        double xDimScaled = Canvas.GetLeft(selectBorder) * m.M11;
        double yDimScaled = Canvas.GetTop(selectBorder) * m.M22;

        Rectangle regionScaled = new Rectangle(
            (int)xDimScaled,
            (int)yDimScaled,
            (int)(selectBorder.Width * m.M11),
            (int)(selectBorder.Height * m.M22));

        string grabbedText = string.Empty;

        if (NewGrabFrameMenuItem.IsChecked is true) {
            PlaceGrabFrameInSelectionRect();
            return;
        }

        try { RegionClickCanvas.Children.Remove(selectBorder); } catch { }

        Language? selectedOcrLang = LanguagesComboBox.SelectedItem as Language;

        if (selectedOcrLang is null)
            selectedOcrLang = LanguageUtilities.GetOCRLanguage();

        if (regionScaled.Width < 3 || regionScaled.Height < 3) {
            BackgroundBrush.Opacity = 0;
            grabbedText = await OcrExtensions.GetClickedWordAsync(this, new System.Windows.Point(xDimScaled, yDimScaled), selectedOcrLang);
        } else if (TableToggleButton.IsChecked is true)
            grabbedText = await OcrExtensions.GetRegionsTextAsTableAsync(this, regionScaled, selectedOcrLang);
        else
            grabbedText = await OcrExtensions.GetRegionsTextAsync(this, regionScaled, selectedOcrLang);

        if (!string.IsNullOrWhiteSpace(grabbedText)) {
            //如果选中打开文件按钮并且输入文本不为null，直接跳编辑页面
            if (SendToEditTextToggleButton.IsChecked is true && destinationTextBox is null) {
                EditTextWindow etw = WindowUtilities.OpenOrActivateWindow<EditTextWindow>();
                destinationTextBox = etw.PassedTextControl;
            }

            HandleTextFromOcr(grabbedText);
            //WindowUtilities.CloseAllFullscreenGrabs();
        } else
            BackgroundBrush.Opacity = 0.2;
    }

    private void SendToEditTextToggleButton_Click(object sender, RoutedEventArgs e) {
        bool isActive = CheckIfCheckingOrUnchecking(sender);
        WindowUtilities.FullscreenKeyDown(Key.E, isActive);
    }

    private void SettingsMenuItem_Click(object sender, RoutedEventArgs e) {
        WindowUtilities.OpenOrActivateWindow<SettingsWindow>();
        WindowUtilities.CloseAllFullscreenGrabs();
    }

    private void SingleLineMenuItem_Click(object? sender = null, RoutedEventArgs? e = null) {
        bool isActive = CheckIfCheckingOrUnchecking(sender);
        WindowUtilities.FullscreenKeyDown(Key.S, isActive);
    }
    /// <summary>
    ///  FullscreenGrab 窗口加载时被调用
    ///将 FullscreenGrab 窗口的状态设置为最大化；
    ///设置 FullWindow 对象代表的矩形的左上角坐标为(0, 0)，右下角坐标为窗口的 Width 和 Height；
    ///给 FullscreenGrab 窗口绑定 KeyDown 和 KeyUp 事件；
    ///调用 SetImageToBackground() 方法将当前屏幕截图设置为窗口的背景图片；
    ///根据用户设置中的 FSGMakeSingleLineToggle 配置项，设置 SingleLineMenuItem 是否选中；
    ///示 TopButtonsStackPanel 工具栏；
    ///调用 LoadOcrLanguages() 方法加载 OCR 支持的语言并生成对应的菜单项
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Window_Loaded(object sender, RoutedEventArgs e) {
        WindowState = WindowState.Maximized;
        FullWindow.Rect = new System.Windows.Rect(0, 0, Width, Height);
        this.KeyDown += FullscreenGrab_KeyDown;
        this.KeyUp += FullscreenGrab_KeyUp;

        SetImageToBackground();

        if (Settings.Default.FSGMakeSingleLineToggle)
            SingleLineMenuItem.IsChecked = true;

        TopButtonsStackPanel.Visibility = Visibility.Visible;

        LoadOcrLanguages();
    }

    private void Window_Unloaded(object sender, RoutedEventArgs e) {
        BackgroundImage.Source = null;
        BackgroundImage.UpdateLayout();
        currentScreen = null;
        dpiScale = null;
        textFromOCR = null;

        this.Loaded -= Window_Loaded;
        this.Unloaded -= Window_Unloaded;

        RegionClickCanvas.MouseDown -= RegionClickCanvas_MouseDown;
        RegionClickCanvas.MouseMove -= RegionClickCanvas_MouseMove;
        RegionClickCanvas.MouseUp -= RegionClickCanvas_MouseUp;

        SingleLineMenuItem.Click -= SingleLineMenuItem_Click;
        FreezeMenuItem.Click -= FreezeMenuItem_Click;
        NewGrabFrameMenuItem.Click -= NewGrabFrameMenuItem_Click;
        SendToEtwMenuItem.Click -= NewEditTextMenuItem_Click;
        SettingsMenuItem.Click -= SettingsMenuItem_Click;
        CancelMenuItem.Click -= CancelMenuItem_Click;

        LanguagesComboBox.SelectionChanged -= LanguagesComboBox_SelectionChanged;

        SingleLineToggleButton.Click -= SingleLineMenuItem_Click;
        FreezeToggleButton.Click -= FreezeMenuItem_Click;
        NewGrabFrameToggleButton.Click -= NewGrabFrameMenuItem_Click;
        SendToEditTextToggleButton.Click -= NewEditTextMenuItem_Click;
        SettingsButton.Click -= SettingsMenuItem_Click;
        CancelButton.Click -= CancelMenuItem_Click;

        this.KeyDown -= FullscreenGrab_KeyDown;
        this.KeyUp -= FullscreenGrab_KeyUp;
    }

    #endregion Methods
}
