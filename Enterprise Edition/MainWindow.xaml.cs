using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using TodoExport.Models;
using TodoExport.Services;
using TodoExport.Utils;

namespace TodoExport;

public partial class MainWindow : Window
{
    private readonly AuthService _authService;
    private readonly GraphApiService _graphApiService;
    private readonly ExportService _exportService;
    private List<TodoList> _todoLists = new();
    private readonly Dictionary<string, CheckBox> _listCheckBoxes = new();
    private CancellationTokenSource? _loginCts;
    private string _verificationUrl = "https://microsoft.com/devicelogin";

    public MainWindow()
    {
        InitializeComponent();
        
        _authService = new AuthService();
        _graphApiService = new GraphApiService(_authService);
        _exportService = new ExportService();

        // Events registrieren
        _authService.DeviceCodeReceived += OnDeviceCodeReceived;
        _authService.AuthenticationCompleted += OnAuthenticationCompleted;
        _authService.AuthenticationFailed += OnAuthenticationFailed;
        
        // Default: Englisch
        Lang.CurrentLanguage = "en";
        UpdateLanguageButtons();
        UpdateAllTexts();
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        StatusText.Text = Lang.Get("CheckingLogin");
        
        try
        {
            if (await _authService.TrySilentSignInAsync())
            {
                await ShowLoggedInStateAsync();
                return;
            }
        }
        catch { }
        
        StatusText.Text = Lang.Get("Ready");
    }

    #region Authentication Events

    private void OnDeviceCodeReceived(string code, string url)
    {
        Dispatcher.Invoke(() =>
        {
            _verificationUrl = url;
            DeviceCodeText.Text = code;
            
            LoginSection.Visibility = Visibility.Collapsed;
            DeviceCodeSection.Visibility = Visibility.Visible;
            
            StatusText.Text = Lang.Get("WaitingForBrowser");
        });
    }

    private void OnAuthenticationCompleted()
    {
        Dispatcher.Invoke(async () =>
        {
            DeviceCodeSection.Visibility = Visibility.Collapsed;
            await ShowLoggedInStateAsync();
        });
    }

    private void OnAuthenticationFailed(string error)
    {
        Dispatcher.Invoke(() =>
        {
            DeviceCodeSection.Visibility = Visibility.Collapsed;
            LoginSection.Visibility = Visibility.Visible;
            
            StatusText.Text = Lang.Get("LoginFailed");
            MessageBox.Show($"{Lang.Get("LoginFailed")}:\n\n{error}", Lang.Get("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        });
    }

    #endregion

    #region Language Switch

    private void LanguageDEButton_Click(object sender, RoutedEventArgs e)
    {
        Lang.CurrentLanguage = "de";
        UpdateLanguageButtons();
        UpdateAllTexts();
    }

    private void LanguageENButton_Click(object sender, RoutedEventArgs e)
    {
        Lang.CurrentLanguage = "en";
        UpdateLanguageButtons();
        UpdateAllTexts();
    }

    private void UpdateLanguageButtons()
    {
        if (Lang.CurrentLanguage == "de")
        {
            LanguageDEButton.Style = (Style)FindResource("PrimaryButton");
            LanguageENButton.Style = (Style)FindResource("SecondaryButton");
        }
        else
        {
            LanguageDEButton.Style = (Style)FindResource("SecondaryButton");
            LanguageENButton.Style = (Style)FindResource("PrimaryButton");
        }
    }

    private void UpdateAllTexts()
    {
        // Header
        AppTitleText.Text = Lang.Get("AppTitle");
        AppSubtitleText.Text = Lang.Get("AppSubtitle");
        LoggedInText.Text = Lang.Get("LoggedIn");
        SignOutButton.Content = Lang.Get("SignOut");
        
        // Login Section
        SignInTitleText.Text = Lang.Get("SignInTitle");
        SignInSubtitleText.Text = Lang.Get("SignInSubtitle");
        SignInButton.Content = Lang.Get("SignInButton");
        
        // Device Code Section
        BrowserLoginText.Text = Lang.Get("BrowserLogin");
        BrowserLoginHintText.Text = Lang.Get("BrowserLoginHint");
        OpenBrowserButton.Content = Lang.Get("OpenBrowser");
        CancelLoginButton.Content = Lang.Get("Cancel");
        CopyCodeButton.ToolTip = Lang.Get("CopyCode");
        
        // Export Section
        SelectListsText.Text = Lang.Get("SelectLists");
        SelectAllButton.Content = Lang.Get("All");
        SelectNoneButton.Content = Lang.Get("None");
        RefreshButton.ToolTip = Lang.Get("RefreshLists");
        ExportOptionsText.Text = Lang.Get("ExportOptions");
        SelectedText.Text = Lang.Get("Selected");
        FormatText.Text = Lang.Get("Format");
        IncludeCompletedCheckBox.Content = Lang.Get("IncludeCompleted");
        ExportButton.Content = Lang.Get("Export");
        ExportHint.Text = Lang.Get("SelectListHint");
        
        // Format ComboBox Items
        FormatExcelItem.Content = "Excel (.xlsx)";
        FormatTodoistCsvItem.Content = "Todoist CSV (.csv)";
        FormatJsonItem.Content = "JSON (.json)";
        FormatTextItem.Content = Lang.CurrentLanguage == "de" ? "Text (.txt)" : "Text (.txt)";
        FormatDetailedTextItem.Content = Lang.CurrentLanguage == "de" ? "Detaillierter Text (.txt)" : "Detailed Text (.txt)";
        FormatDetailedCsvItem.Content = Lang.CurrentLanguage == "de" ? "Detaillierte CSV (.csv)" : "Detailed CSV (.csv)";
        
        // Update format description
        if (FormatDescription != null)
        {
            FormatDescription.Text = GetSelectedFormat() switch
            {
                ExportFormat.Excel => Lang.Get("FormatExcel"),
                ExportFormat.TodoistCsv => Lang.Get("FormatTodoistCsv"),
                ExportFormat.Json => Lang.Get("FormatJson"),
                ExportFormat.Text => Lang.Get("FormatText"),
                ExportFormat.DetailedText => Lang.Get("FormatDetailedText"),
                ExportFormat.DetailedCsv => Lang.Get("FormatDetailedCsv"),
                _ => ""
            };
        }
        
        // Update selected count label
        UpdateSelectedCount();
        
        // Update status if needed
        if (StatusText.Text == Lang.Get("Ready") || StatusText.Text == "Ready" || StatusText.Text == "Bereit")
        {
            StatusText.Text = Lang.Get("Ready");
        }
    }

    #endregion

    #region Button Handlers

    private async void SignInButton_Click(object sender, RoutedEventArgs e)
    {
        SignInButton.IsEnabled = false;
        _loginCts = new CancellationTokenSource();

        try
        {
            StatusText.Text = Lang.Get("StartingLogin");
            await _authService.SignInAsync(_loginCts.Token);
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = Lang.Get("LoginCancelled");
            DeviceCodeSection.Visibility = Visibility.Collapsed;
            LoginSection.Visibility = Visibility.Visible;
        }
        catch (Exception ex)
        {
            StatusText.Text = Lang.Get("Error");
            DeviceCodeSection.Visibility = Visibility.Collapsed;
            LoginSection.Visibility = Visibility.Visible;
            MessageBox.Show($"{Lang.Get("Error")}: {ex.Message}", Lang.Get("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SignInButton.IsEnabled = true;
            _loginCts?.Dispose();
            _loginCts = null;
        }
    }

    private void CancelLoginButton_Click(object sender, RoutedEventArgs e)
    {
        _loginCts?.Cancel();
        DeviceCodeSection.Visibility = Visibility.Collapsed;
        LoginSection.Visibility = Visibility.Visible;
        StatusText.Text = Lang.Get("LoginCancelled");
    }

    private void CopyCodeButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(DeviceCodeText.Text);
            StatusText.Text = Lang.Get("CopyCode") + "!";
        }
        catch { }
    }

    private void OpenBrowserButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo { FileName = _verificationUrl, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"{Lang.Get("Error")}:\n{ex.Message}", Lang.Get("Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void SignOutButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await _authService.SignOutAsync();
            ShowLoggedOutState();
            StatusText.Text = Lang.Get("SignedOut");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"{Lang.Get("Error")}: {ex.Message}", Lang.Get("Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadListsAsync();
    }

    private void SelectAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var cb in _listCheckBoxes.Values)
            cb.IsChecked = true;
        UpdateSelectedCount();
    }

    private void SelectNoneButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var cb in _listCheckBoxes.Values)
            cb.IsChecked = false;
        UpdateSelectedCount();
    }

    private void FormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (FormatDescription == null) return;
        
        FormatDescription.Text = GetSelectedFormat() switch
        {
            ExportFormat.Excel => Lang.Get("FormatExcel"),
            ExportFormat.TodoistCsv => Lang.Get("FormatTodoistCsv"),
            ExportFormat.Json => Lang.Get("FormatJson"),
            ExportFormat.Text => Lang.Get("FormatText"),
            ExportFormat.DetailedText => Lang.Get("FormatDetailedText"),
            ExportFormat.DetailedCsv => Lang.Get("FormatDetailedCsv"),
            _ => ""
        };
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedLists = GetSelectedLists();
        
        if (selectedLists.Count == 0)
        {
            MessageBox.Show(Lang.Get("PleaseSelectList"), Lang.Get("Hint"), MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var format = GetSelectedFormat();
        var (filter, ext) = format switch
        {
            ExportFormat.Excel => (Lang.CurrentLanguage == "de" ? "Excel Dateien (*.xlsx)|*.xlsx" : "Excel Files (*.xlsx)|*.xlsx", ".xlsx"),
            ExportFormat.TodoistCsv or ExportFormat.DetailedCsv => (Lang.CurrentLanguage == "de" ? "CSV Dateien (*.csv)|*.csv" : "CSV Files (*.csv)|*.csv", ".csv"),
            ExportFormat.Json => (Lang.CurrentLanguage == "de" ? "JSON Dateien (*.json)|*.json" : "JSON Files (*.json)|*.json", ".json"),
            _ => (Lang.CurrentLanguage == "de" ? "Text Dateien (*.txt)|*.txt" : "Text Files (*.txt)|*.txt", ".txt")
        };

        var dialog = new SaveFileDialog
        {
            Filter = filter,
            FileName = $"todo-export-{DateTime.Now:yyyy-MM-dd}{ext}"
        };

        if (dialog.ShowDialog() != true) return;

        SetLoading(true, Lang.Get("Exporting"));

        try
        {
            var tasks = await _graphApiService.GetAllTasksAsync(selectedLists);
            var includeCompleted = IncludeCompletedCheckBox.IsChecked == true;

            var data = new Dictionary<TodoList, List<TodoTask>>();
            foreach (var list in selectedLists)
            {
                var listTasks = tasks.TryGetValue(list.Id, out var t) ? t : new List<TodoTask>();
                
                // Filter: Abgeschlossene Aufgaben ausschließen wenn Option deaktiviert
                if (!includeCompleted)
                {
                    listTasks = listTasks.Where(task => task.Status != "completed").ToList();
                }
                
                data[list] = listTasks;
            }

            var total = data.Values.Sum(t => t.Count);

            if (format == ExportFormat.Excel)
            {
                // Excel direkt als Datei speichern
                _exportService.ExportToExcel(data, dialog.FileName);
            }
            else
            {
                // Andere Formate als Text
                string content = format switch
                {
                    ExportFormat.TodoistCsv => _exportService.ExportToTodoistCsv(data),
                    ExportFormat.DetailedCsv => _exportService.ExportToDetailedCsv(data),
                    ExportFormat.Json => _exportService.ExportToJson(data),
                    ExportFormat.Text => _exportService.ExportToText(data),
                    ExportFormat.DetailedText => _exportService.ExportToDetailedText(data),
                    _ => throw new NotSupportedException()
                };

                await File.WriteAllTextAsync(dialog.FileName, content, Encoding.UTF8);
            }

            var listWord = data.Count == 1 ? Lang.Get("List") : Lang.Get("Lists");
            var taskWord = total == 1 ? (Lang.CurrentLanguage == "de" ? "Aufgabe" : "task") : (Lang.CurrentLanguage == "de" ? "Aufgaben" : "tasks");
            StatusText.Text = $"✓ {total} {taskWord} {Lang.Get("ExportSuccess")}";
            
            var result = MessageBox.Show(
                $"{Lang.Get("ExportSuccessMessage")}" +
                $"📁 {Path.GetFileName(dialog.FileName)}\n" +
                $"📋 {data.Count} {listWord}\n" +
                $"✓ {total} {taskWord}\n\n" +
                $"{Lang.Get("ExportCompleteDetails")}",
                Lang.Get("ExportComplete"), 
                MessageBoxButton.YesNo, 
                MessageBoxImage.Information);
                
            if (result == MessageBoxResult.Yes)
            {
                Process.Start("explorer.exe", $"/select,\"{dialog.FileName}\"");
            }
        }
        catch (Exception ex)
        {
            StatusText.Text = Lang.Get("ExportFailed");
            MessageBox.Show($"{Lang.Get("ExportFailed")}:\n\n{ex.Message}", Lang.Get("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetLoading(false);
        }
    }

    #endregion

    #region State Management

    private async Task ShowLoggedInStateAsync()
    {
        LoginSection.Visibility = Visibility.Collapsed;
        DeviceCodeSection.Visibility = Visibility.Collapsed;
        ExportSection.Visibility = Visibility.Visible;
        UserPanel.Visibility = Visibility.Visible;
        
        var userName = _authService.UserName ?? (Lang.CurrentLanguage == "de" ? "Benutzer" : "User");
        UserNameText.Text = userName;
        UserInitial.Text = userName.Length > 0 ? userName[0].ToString().ToUpper() : "?";

        await LoadListsAsync();
    }

    private void ShowLoggedOutState()
    {
        LoginSection.Visibility = Visibility.Visible;
        DeviceCodeSection.Visibility = Visibility.Collapsed;
        ExportSection.Visibility = Visibility.Collapsed;
        UserPanel.Visibility = Visibility.Collapsed;
        _todoLists.Clear();
        _listCheckBoxes.Clear();
        ListsPanel.Children.Clear();
    }

    private async Task LoadListsAsync()
    {
        SetLoading(true, Lang.Get("LoadingLists"));

        try
        {
            _todoLists = await _graphApiService.GetTodoListsAsync();
            _listCheckBoxes.Clear();
            ListsPanel.Children.Clear();

            foreach (var list in _todoLists)
            {
                var checkBox = new CheckBox
                {
                    Content = list.DisplayName,
                    Tag = list,
                    IsChecked = true,
                    Style = (Style)FindResource("ListCheckBox")
                };
                checkBox.Checked += (s, e) => UpdateSelectedCount();
                checkBox.Unchecked += (s, e) => UpdateSelectedCount();
                
                _listCheckBoxes[list.Id] = checkBox;
                ListsPanel.Children.Add(checkBox);
            }

            ListCount.Text = $"({_todoLists.Count})";
            UpdateSelectedCount();
            StatusText.Text = $"{_todoLists.Count} {Lang.Get("ListsLoaded")}";
        }
        catch (Exception ex)
        {
            StatusText.Text = Lang.Get("LoadingListsError");
            MessageBox.Show($"{Lang.Get("Error")}:\n\n{ex.Message}", Lang.Get("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetLoading(false);
        }
    }

    private void UpdateSelectedCount()
    {
        var count = _listCheckBoxes.Values.Count(cb => cb.IsChecked == true);
        SelectedCount.Text = count.ToString();
        SelectedLabel.Text = count == 1 ? $" {Lang.Get("List")}" : $" {Lang.Get("Lists")}";
        ExportButton.IsEnabled = count > 0;
        ExportHint.Visibility = count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private List<TodoList> GetSelectedLists()
    {
        return _listCheckBoxes.Values
            .Where(cb => cb.IsChecked == true && cb.Tag is TodoList)
            .Select(cb => (TodoList)cb.Tag)
            .ToList();
    }

    private ExportFormat GetSelectedFormat()
    {
        return FormatComboBox.SelectedItem is ComboBoxItem { Tag: string tag } ? tag switch
        {
            "Excel" => ExportFormat.Excel,
            "TodoistCsv" => ExportFormat.TodoistCsv,
            "Json" => ExportFormat.Json,
            "Text" => ExportFormat.Text,
            "DetailedText" => ExportFormat.DetailedText,
            "DetailedCsv" => ExportFormat.DetailedCsv,
            _ => ExportFormat.Excel
        } : ExportFormat.Excel;
    }

    private void SetLoading(bool loading, string? msg = null)
    {
        LoadingBar.Visibility = loading ? Visibility.Visible : Visibility.Collapsed;
        ExportButton.IsEnabled = !loading && _listCheckBoxes.Values.Any(cb => cb.IsChecked == true);
        RefreshButton.IsEnabled = !loading;
        if (msg != null) StatusText.Text = msg;
    }

    #endregion
}
