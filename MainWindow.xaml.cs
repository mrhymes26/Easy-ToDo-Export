using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using TodoExport.Models;
using TodoExport.Services;

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
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        StatusText.Text = "Prüfe Anmeldung...";
        
        try
        {
            if (await _authService.TrySilentSignInAsync())
            {
                await ShowLoggedInStateAsync();
                return;
            }
        }
        catch { }
        
        StatusText.Text = "Bereit";
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
            
            StatusText.Text = "Warte auf Browser-Anmeldung...";
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
            
            StatusText.Text = "Anmeldung fehlgeschlagen";
            MessageBox.Show($"Anmeldung fehlgeschlagen:\n\n{error}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        });
    }

    #endregion

    #region Button Handlers

    private async void SignInButton_Click(object sender, RoutedEventArgs e)
    {
        SignInButton.IsEnabled = false;
        _loginCts = new CancellationTokenSource();

        try
        {
            StatusText.Text = "Starte Anmeldung...";
            await _authService.SignInAsync(_loginCts.Token);
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = "Anmeldung abgebrochen";
            DeviceCodeSection.Visibility = Visibility.Collapsed;
            LoginSection.Visibility = Visibility.Visible;
        }
        catch (Exception ex)
        {
            StatusText.Text = "Fehler";
            DeviceCodeSection.Visibility = Visibility.Collapsed;
            LoginSection.Visibility = Visibility.Visible;
            MessageBox.Show($"Fehler: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
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
        StatusText.Text = "Abgebrochen";
    }

    private void CopyCodeButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(DeviceCodeText.Text);
            StatusText.Text = "Code kopiert!";
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
            MessageBox.Show($"Browser konnte nicht geöffnet werden:\n{ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void SignOutButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await _authService.SignOutAsync();
            ShowLoggedOutState();
            StatusText.Text = "Abgemeldet";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Warning);
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
            ExportFormat.Excel => "Formatierte Tabelle mit Übersicht und Farben",
            ExportFormat.TodoistCsv => "Kompatibel mit Todoist Import-Funktion",
            ExportFormat.Json => "Strukturiertes Format für Entwickler",
            ExportFormat.Text => "Einfache Textliste zum Lesen",
            ExportFormat.DetailedText => "Ausführliche Informationen pro Aufgabe",
            ExportFormat.DetailedCsv => "Alle Details als Tabelle",
            _ => ""
        };
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedLists = GetSelectedLists();
        
        if (selectedLists.Count == 0)
        {
            MessageBox.Show("Bitte wählen Sie mindestens eine Liste aus.", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var format = GetSelectedFormat();
        var (filter, ext) = format switch
        {
            ExportFormat.Excel => ("Excel Dateien (*.xlsx)|*.xlsx", ".xlsx"),
            ExportFormat.TodoistCsv or ExportFormat.DetailedCsv => ("CSV Dateien (*.csv)|*.csv", ".csv"),
            ExportFormat.Json => ("JSON Dateien (*.json)|*.json", ".json"),
            _ => ("Text Dateien (*.txt)|*.txt", ".txt")
        };

        var dialog = new SaveFileDialog
        {
            Filter = filter,
            FileName = $"todo-export-{DateTime.Now:yyyy-MM-dd}{ext}"
        };

        if (dialog.ShowDialog() != true) return;

        SetLoading(true, "Exportiere Aufgaben...");

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

            StatusText.Text = $"✓ {total} Aufgaben aus {data.Count} Listen exportiert";
            
            var result = MessageBox.Show(
                $"Export erfolgreich!\n\n" +
                $"📁 {Path.GetFileName(dialog.FileName)}\n" +
                $"📋 {data.Count} Listen\n" +
                $"✓ {total} Aufgaben\n\n" +
                $"Datei im Explorer anzeigen?",
                "Export abgeschlossen", 
                MessageBoxButton.YesNo, 
                MessageBoxImage.Information);
                
            if (result == MessageBoxResult.Yes)
            {
                Process.Start("explorer.exe", $"/select,\"{dialog.FileName}\"");
            }
        }
        catch (Exception ex)
        {
            StatusText.Text = "Export fehlgeschlagen";
            MessageBox.Show($"Fehler beim Export:\n\n{ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
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
        
        var userName = _authService.UserName ?? "Benutzer";
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
        SetLoading(true, "Lade Listen...");

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
            StatusText.Text = $"{_todoLists.Count} Listen geladen";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Fehler beim Laden";
            MessageBox.Show($"Fehler:\n\n{ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
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
        SelectedLabel.Text = count == 1 ? " Liste" : " Listen";
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
