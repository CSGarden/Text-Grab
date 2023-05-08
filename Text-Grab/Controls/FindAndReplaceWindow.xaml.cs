using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Text_Grab.Models;
using Text_Grab.Utilities;
using Wpf.Ui.Controls.Window;

namespace Text_Grab.Controls;

/// <summary>
/// Interaction logic for FindAndReplaceWindow.xaml
/// </summary>
public partial class FindAndReplaceWindow : FluentWindow {
    #region Fields

    public static RoutedCommand CopyMatchesCmd = new();
    public static RoutedCommand DeleteAllCmd = new();
    public static RoutedCommand ExtractPatternCmd = new();
    public static RoutedCommand ReplaceAllCmd = new();
    public static RoutedCommand ReplaceOneCmd = new();
    public static RoutedCommand TextSearchCmd = new();
    DispatcherTimer ChangeFindTextTimer = new();
    private MatchCollection? Matches;
    private string stringFromWindow = "";
    private EditTextWindow? textEditWindow;

    #endregion Fields

    #region Constructors

    public FindAndReplaceWindow() {
        InitializeComponent();

        ChangeFindTextTimer.Interval = TimeSpan.FromMilliseconds(400);
        ChangeFindTextTimer.Tick -= ChangeFindText_Tick;
        ChangeFindTextTimer.Tick += ChangeFindText_Tick;
    }

    #endregion Constructors

    #region Properties

    public List<FindResult> FindResults { get; set; } = new();

    public string StringFromWindow {
        get { return stringFromWindow; }
        set { stringFromWindow = value; }
    }
    public EditTextWindow? TextEditWindow {
        get {
            return textEditWindow;
        }
        set {
            textEditWindow = value;

            if (textEditWindow is not null)
                textEditWindow.PassedTextControl.TextChanged += EditTextBoxChanged;
        }
    }
    private string? Pattern { get; set; }

    #endregion Properties

    #region Methods
    /// <summary>
    /// 将查找结果显示在界面上，并尝试通过一些选项设置和匹配模式来改善查找行为。<br/>
    /// 在这个程序中对接收的字符串中搜索匹配给正则表达式模式或者匹配文本<br/>
    /// </summary>
    public void SearchForText() {
        //清除之前的搜索结果并更新界面
        FindResults.Clear();
        ResultsListView.ItemsSource = null;
        //获取要搜索的模式或文本，并将其保存到类字段 “Pattern” 中。
        Pattern = FindTextBox.Text;
        // 如果未选择 "UsePaternCheckBox" 并且已经选中了 "ExactMatchCheckBox"，则对 Pattern 进行转义处理。
        if (UsePaternCheckBox.IsChecked is false && ExactMatchCheckBox.IsChecked is bool matchExactly)
            Pattern = Pattern.EscapeSpecialRegexChars(matchExactly);
        // 使用正则表达式和搜索选项 ("ExactMatchCheckBox" 和 按钮等) 在 StringFromWindow 中查找所有匹配项。如果遇到错误，则捕获异常并返回错误信息。
        try {
            if (ExactMatchCheckBox.IsChecked is true)
                Matches = Regex.Matches(StringFromWindow, Pattern, RegexOptions.Multiline);
            else
                Matches = Regex.Matches(StringFromWindow, Pattern, RegexOptions.Multiline | RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
        } catch (Exception ex) {
            MatchesText.Text = "Error searching: " + ex.GetType().ToString();
            return;
        }
        //如果找不到匹配项，则显示 "0 Matches" (零匹配项) 的结果

        if (Matches.Count < 1 || string.IsNullOrWhiteSpace(FindTextBox.Text)) {
            MatchesText.Text = "0 Matches";
            return;
        }
        //在 UI 界面中显示所有匹配项的数量（Matches.Count）和列表 (ResultsListView) 中的所有项目。
        if (Matches.Count == 1)
            MatchesText.Text = $"{Matches.Count} Match";
        else
            MatchesText.Text = $"{Matches.Count} Matches";

        ResultsListView.IsEnabled = true;
        int count = 1;
        //为每个匹配字符串创建一个“FindResult”对象, 对象包括：索引、第 N 次出现、该字符串及其前后预览字符串片段，并将对象添加到 List<FindResult> (FindResults) 类型集合中。
        foreach (Match m in Matches) {
            FindResult fr = new() {
                Index = m.Index,
                Text = m.Value.MakeStringSingleLine(),
                PreviewLeft = GetCharactersToLeftOfNewLine(ref stringFromWindow, m.Index, 12).MakeStringSingleLine(),
                PreviewRight = GetCharactersToRightOfNewLine(ref stringFromWindow, m.Index + m.Length, 12).MakeStringSingleLine(),
                Count = count
            };
            FindResults.Add(fr);

            count++;
        }
        //将 ItemsSource 属性设置为 FindResults 集合，以便在 ResultsListView 列表中显示所有查找结果。
        ResultsListView.ItemsSource = FindResults;

        //如果 textEditWindow 及其控件已定义并具有焦点，则高亮显示第一次匹配项，并将光标设置为该文本控件。但如果没有定义，就不会高亮第一个匹配项。
        Match? firstMatch = Matches[0];
        if (textEditWindow is not null
            && firstMatch is not null
            && this.IsFocused) {
            textEditWindow.PassedTextControl.Select(firstMatch.Index, firstMatch.Value.Length);
            textEditWindow.PassedTextControl.Focus();
            this.Focus();
        }
    }

    public void ShouldCloseWithThisETW(EditTextWindow etw) {
        if (textEditWindow is not null && etw == textEditWindow)
            Close();
    }

    // a method which uses GetNewLineIndexToLeft and returns the string from the given index to the newLine character to the left of the given index
    // if the string is longer the x number of characters, it will return the last x number of characters
    // and if the string is at the beginning don't add "..." to the beginning
    private static string GetCharactersToLeftOfNewLine(ref string mainString, int index, int numberOfCharacters) {
        int newLineIndex = GetNewLineIndexToLeft(ref mainString, index);

        if (newLineIndex < 1)
            return mainString.Substring(0, index);

        newLineIndex++;

        if (newLineIndex > mainString.Length)
            return mainString.Substring(mainString.Length, 0);

        if (index - newLineIndex < 0)
            return mainString.Substring(newLineIndex, 0);

        if (index - newLineIndex < numberOfCharacters)
            return "..." + mainString.Substring(newLineIndex, index - newLineIndex);

        return "..." + mainString.Substring(index - numberOfCharacters, numberOfCharacters);
    }

    // same as GetCharactersToLeftOfNewLine but to the right
    private static string GetCharactersToRightOfNewLine(ref string mainString, int index, int numberOfCharacters) {
        int newLineIndex = GetNewLineIndexToRight(ref mainString, index);
        if (newLineIndex < 1)
            return mainString.Substring(index);

        if (newLineIndex - index > numberOfCharacters)
            return mainString.Substring(index, numberOfCharacters) + "...";

        if (newLineIndex == mainString.Length)
            return mainString.Substring(index);

        return mainString.Substring(index, newLineIndex - index) + "...";
    }

    // a method which returns the nearst newLine character index to the left of the given index
    private static int GetNewLineIndexToLeft(ref string mainString, int index) {
        char newLineChar = Environment.NewLine.ToArray().Last();

        int newLineIndex = index;
        while (newLineIndex > 0 && newLineIndex < mainString.Length && mainString[newLineIndex] != newLineChar)
            newLineIndex--;

        return newLineIndex;
    }

    // a method which returns the nearst newLine character index to the right of the given index
    private static int GetNewLineIndexToRight(ref string mainString, int index) {
        char newLineChar = Environment.NewLine.ToArray().First();

        int newLineIndex = index;
        while (newLineIndex < mainString.Length && mainString[newLineIndex] != newLineChar)
            newLineIndex++;

        return newLineIndex;
    }

    private void ChangeFindText_Tick(object? sender, EventArgs? e) {
        ChangeFindTextTimer.Stop();
        SearchForText();
    }

    private void CopyMatchesCmd_CanExecute(object sender, CanExecuteRoutedEventArgs e) {
        if (Matches is null || Matches.Count < 1 || string.IsNullOrEmpty(FindTextBox.Text))
            e.CanExecute = false;
        else
            e.CanExecute = true;
    }

    private void CopyMatchesCmd_Executed(object sender, ExecutedRoutedEventArgs e) {
        if (Matches is null
            || textEditWindow is null
            || Matches.Count < 1)
            return;

        StringBuilder stringBuilder = new();

        var selection = ResultsListView.SelectedItems;
        if (selection.Count < 2)
            selection = ResultsListView.Items;

        foreach (var item in selection)
            if (item is FindResult findResult)
                stringBuilder.AppendLine(findResult.Text);

        EditTextWindow etw = new();
        etw.AddThisText(stringBuilder.ToString());
        etw.Show();
    }

    private void DeleteAll_CanExecute(object sender, CanExecuteRoutedEventArgs e) {
        if (Matches is not null && Matches.Count > 1 && !string.IsNullOrEmpty(FindTextBox.Text))
            e.CanExecute = true;
        else
            e.CanExecute = false;
    }

    private void DeleteAll_Executed(object sender, ExecutedRoutedEventArgs e) {
        if (Matches is null
            || Matches.Count < 1
            || textEditWindow is null)
            return;

        var selection = ResultsListView.SelectedItems;
        if (selection.Count < 2)
            selection = ResultsListView.Items;

        for (int j = selection.Count - 1; j >= 0; j--) {
            if (selection[j] is not FindResult selectedResult)
                continue;

            textEditWindow.PassedTextControl.Select(selectedResult.Index, selectedResult.Length);
            textEditWindow.PassedTextControl.SelectedText = string.Empty;
        }

        textEditWindow.PassedTextControl.Select(0, 0);
        SearchForText();
    }

    private void EditTextBoxChanged(object sender, TextChangedEventArgs e) {
        ChangeFindTextTimer.Stop();
        if (textEditWindow is not null)
            StringFromWindow = textEditWindow.PassedTextControl.Text;

        ChangeFindTextTimer.Start();
    }

    private void ExtractPattern_CanExecute(object sender, CanExecuteRoutedEventArgs e) {
        if (textEditWindow is not null
            && textEditWindow.PassedTextControl.SelectedText.Length > 0)
            e.CanExecute = true;
        else
            e.CanExecute = false;
    }

    private void ExtractPattern_Executed(object sender, ExecutedRoutedEventArgs e) {
        if (textEditWindow is null)
            return;

        string? selection = textEditWindow.PassedTextControl.SelectedText;

        string simplePattern = selection.ExtractSimplePattern();

        UsePaternCheckBox.IsChecked = true;
        FindTextBox.Text = simplePattern;

        SearchForText();
    }

    private void FindAndReplacedLoaded(object sender, RoutedEventArgs e) {
        if (!string.IsNullOrWhiteSpace(FindTextBox.Text))
            SearchForText();

        FindTextBox.Focus();
    }
    /// <summary>
    /// 在用输入文本时，实时搜索文本框中的内容，并且在用户按下回车键时，执行搜索操作。
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void FindTextBox_KeyUp(object sender, KeyEventArgs e) {
        ChangeFindTextTimer.Stop();

        if (e.Key == Key.Enter) {
            ChangeFindTextTimer.Stop();
            SearchForText();
            e.Handled = true;
        } else {
            ChangeFindTextTimer.Start();
        }
    }

    private void MoreOptionsToggleButton_Click(object sender, RoutedEventArgs e) {
        Visibility optionsVisibility = Visibility.Collapsed;
        if (MoreOptionsToggleButton.IsChecked is true)
            optionsVisibility = Visibility.Visible;

        SetExtraOptionsVisibility(optionsVisibility);
    }

    private void OptionsChangedRefresh(object sender, RoutedEventArgs e) {
        SearchForText();
    }

    private void Replace_CanExecute(object sender, CanExecuteRoutedEventArgs e) {
        if (string.IsNullOrEmpty(ReplaceTextBox.Text)
            || Matches is null
            || Matches.Count < 1)
            e.CanExecute = false;
        else
            e.CanExecute = true;
    }
    /// <summary>
    /// 速替换文本，并显示与原始搜索条件相同的所有结果(在 <see cref="SearchForText()"/> 方法中)。
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Replace_Executed(object sender, ExecutedRoutedEventArgs e) {
        //检查 "Matches" 和 "textEditWindow" 是否为空以及搜索结果是否为空。
        if (Matches is null
            || textEditWindow is null
            || ResultsListView.Items.Count is 0)
            return;
        //如果当前没有选择项，则将选中索引设置为第一个匹配项。
        if (ResultsListView.SelectedIndex == -1)
            ResultsListView.SelectedIndex = 0;
        //检查已选项目是否是类型为 FindResult 的对象。如果不是，则返回。
        if (ResultsListView.SelectedItem is not FindResult selectedResult)
            return;
        //如果选中的项目是可用的 FindResult 对象，则通过选中文本控件并设置其所选文本来替换选定的文本。
        //即使用 ReplaceTextBox.Text 替换之前找到的 selectedResult。
        textEditWindow.PassedTextControl.Select(selectedResult.Index, selectedResult.Length);
        textEditWindow.PassedTextControl.SelectedText = ReplaceTextBox.Text;
        //重新开始搜索操作 SearchForText()，以便可以更新显示的结果列表和 textEditWindow 中所选内容的高亮显示，同时可能会对其他操作（如撤销或重做）进行更改以保持同步。
        SearchForText();
    }
    /// <summary>
    /// 允许更快地替换在整个字符串中找到的所有实例，并且与原始搜索条件相同，例如区分大小写等功能。
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ReplaceAll_Executed(object sender, ExecutedRoutedEventArgs e) {
        //检查边界条件，即 Matches 是否为空或不足一个，以及有没有定义 textEditWindow 窗口。
        if (Matches is null
            || Matches.Count < 1
            || textEditWindow is null)
            return;
        //获取所有选择项，如果没有选定任何项目，则将其设置为 ResultsListView.Items 列表中的所有项目。
        var selection = ResultsListView.SelectedItems;
        if (selection.Count < 2)
            selection = ResultsListView.Items;
        //遍历所有选择项，通过选中文本控件并替换选定的文本来逐个替换。即使用 ReplaceTextBox.Text 替换之前找到的 selectedResult。
        for (int j = selection.Count - 1; j >= 0; j--) {
            if (selection[j] is not FindResult selectedResult)
                continue;

            textEditWindow.PassedTextControl.Select(selectedResult.Index, selectedResult.Length);
            textEditWindow.PassedTextControl.SelectedText = ReplaceTextBox.Text;
        }

        SearchForText();
    }

    private void ResultsListView_SelectionChanged(object sender, SelectionChangedEventArgs e) {
        if (ResultsListView.SelectedItem is not FindResult selectedResult)
            return;

        if (textEditWindow is not null) {
            textEditWindow.PassedTextControl.Focus();
            textEditWindow.PassedTextControl.Select(selectedResult.Index, selectedResult.Length);
            this.Focus();
        }
    }

    private void SetExtraOptionsVisibility(Visibility optionsVisibility) {
        ReplaceTextBox.Visibility = optionsVisibility;
        ReplaceButton.Visibility = optionsVisibility;
        ReplaceAllButton.Visibility = optionsVisibility;
        MoreOptionsHozStack.Visibility = optionsVisibility;
        EvenMoreOptionsHozStack.Visibility = optionsVisibility;
    }

    private void TextSearch_CanExecute(object sender, CanExecuteRoutedEventArgs e) {
        if (string.IsNullOrWhiteSpace(FindTextBox.Text))
            e.CanExecute = false;
        else
            e.CanExecute = true;
    }

    private void TextSearch_Executed(object sender, ExecutedRoutedEventArgs e) {
        SearchForText();
    }

    private void Window_Closed(object? sender, EventArgs e) {
        ChangeFindTextTimer.Tick -= ChangeFindText_Tick;
        if (textEditWindow is not null)
            textEditWindow.PassedTextControl.TextChanged -= EditTextBoxChanged;
    }
    private void Window_KeyUp(object sender, KeyEventArgs e) {
        if (e.Key == Key.Escape) {
            if (!string.IsNullOrWhiteSpace(FindTextBox.Text))
                FindTextBox.Clear();
            else
                this.Close();
        }
    }

    #endregion Methods
}
