using System.Text;
using ClosedXML.Excel;
using Newtonsoft.Json;
using TodoExport.Models;

namespace TodoExport.Services;

public class ExportService
{
    /// <summary>
    /// Export als formatierte Excel-Datei (.xlsx)
    /// </summary>
    public void ExportToExcel(Dictionary<TodoList, List<TodoTask>> data, string filePath)
    {
        using var workbook = new XLWorkbook();
        
        // Übersichts-Sheet
        var summarySheet = workbook.Worksheets.Add("Übersicht");
        CreateSummarySheet(summarySheet, data);
        
        // Alle Aufgaben in einem Sheet
        var allTasksSheet = workbook.Worksheets.Add("Alle Aufgaben");
        CreateAllTasksSheet(allTasksSheet, data);
        
        // Ein Sheet pro Liste
        foreach (var (list, tasks) in data)
        {
            var sheetName = SanitizeSheetName(list.DisplayName);
            var sheet = workbook.Worksheets.Add(sheetName);
            CreateListSheet(sheet, list, tasks);
        }
        
        workbook.SaveAs(filePath);
    }

    private void CreateSummarySheet(IXLWorksheet sheet, Dictionary<TodoList, List<TodoTask>> data)
    {
        // Titel
        sheet.Cell(1, 1).Value = "Microsoft To Do Export";
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        
        sheet.Cell(2, 1).Value = $"Exportiert am: {DateTime.Now:dd.MM.yyyy HH:mm}";
        sheet.Cell(2, 1).Style.Font.FontColor = XLColor.Gray;
        
        // Statistik
        sheet.Cell(4, 1).Value = "Zusammenfassung";
        sheet.Cell(4, 1).Style.Font.Bold = true;
        sheet.Cell(4, 1).Style.Font.FontSize = 14;
        
        var totalTasks = data.Values.Sum(t => t.Count);
        var completedTasks = data.Values.Sum(t => t.Count(task => task.Status == "completed"));
        var openTasks = totalTasks - completedTasks;
        
        sheet.Cell(6, 1).Value = "Anzahl Listen:";
        sheet.Cell(6, 2).Value = data.Count;
        sheet.Cell(7, 1).Value = "Aufgaben gesamt:";
        sheet.Cell(7, 2).Value = totalTasks;
        sheet.Cell(8, 1).Value = "Offen:";
        sheet.Cell(8, 2).Value = openTasks;
        sheet.Cell(9, 1).Value = "Erledigt:";
        sheet.Cell(9, 2).Value = completedTasks;
        
        // Listen-Übersicht
        sheet.Cell(11, 1).Value = "Listen";
        sheet.Cell(11, 1).Style.Font.Bold = true;
        sheet.Cell(11, 1).Style.Font.FontSize = 14;
        
        var headers = new[] { "Liste", "Aufgaben", "Offen", "Erledigt" };
        for (int i = 0; i < headers.Length; i++)
        {
            sheet.Cell(13, i + 1).Value = headers[i];
            sheet.Cell(13, i + 1).Style.Font.Bold = true;
            sheet.Cell(13, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#0078D4");
            sheet.Cell(13, i + 1).Style.Font.FontColor = XLColor.White;
        }
        
        int row = 14;
        foreach (var (list, tasks) in data)
        {
            var completed = tasks.Count(t => t.Status == "completed");
            sheet.Cell(row, 1).Value = list.DisplayName;
            sheet.Cell(row, 2).Value = tasks.Count;
            sheet.Cell(row, 3).Value = tasks.Count - completed;
            sheet.Cell(row, 4).Value = completed;
            
            if (row % 2 == 0)
            {
                sheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#F5F5F5");
            }
            row++;
        }
        
        sheet.Columns().AdjustToContents();
    }

    private void CreateAllTasksSheet(IXLWorksheet sheet, Dictionary<TodoList, List<TodoTask>> data)
    {
        var headers = new[] { "Liste", "Aufgabe", "Status", "Priorität", "Fällig am", "Erstellt am", "Kategorien", "Notizen" };
        
        for (int i = 0; i < headers.Length; i++)
        {
            sheet.Cell(1, i + 1).Value = headers[i];
            sheet.Cell(1, i + 1).Style.Font.Bold = true;
            sheet.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#0078D4");
            sheet.Cell(1, i + 1).Style.Font.FontColor = XLColor.White;
        }
        
        int row = 2;
        foreach (var (list, tasks) in data)
        {
            foreach (var task in tasks)
            {
                sheet.Cell(row, 1).Value = list.DisplayName;
                sheet.Cell(row, 2).Value = task.Title;
                sheet.Cell(row, 3).Value = task.Status == "completed" ? "✓ Erledigt" : "○ Offen";
                sheet.Cell(row, 4).Value = task.Importance switch
                {
                    "high" => "🔴 Hoch",
                    "low" => "🟢 Niedrig",
                    _ => "Normal"
                };
                
                if (task.DueDateTime != null && DateTime.TryParse(task.DueDateTime.DateTime, out var dueDate))
                {
                    sheet.Cell(row, 5).Value = dueDate.ToString("dd.MM.yyyy");
                }
                
                if (task.CreatedDateTime.HasValue)
                {
                    sheet.Cell(row, 6).Value = task.CreatedDateTime.Value.ToString("dd.MM.yyyy");
                }
                
                sheet.Cell(row, 7).Value = string.Join(", ", task.Categories);
                sheet.Cell(row, 8).Value = task.Body?.Content ?? "";
                
                // Status-Farbe
                if (task.Status == "completed")
                {
                    sheet.Cell(row, 3).Style.Font.FontColor = XLColor.FromHtml("#107C10");
                }
                
                // Priorität-Farbe
                if (task.Importance == "high")
                {
                    sheet.Cell(row, 4).Style.Font.FontColor = XLColor.FromHtml("#D13438");
                }
                
                // Zeilen-Hintergrund
                if (row % 2 == 0)
                {
                    sheet.Range(row, 1, row, 8).Style.Fill.BackgroundColor = XLColor.FromHtml("#F5F5F5");
                }
                
                row++;
            }
        }
        
        sheet.Columns().AdjustToContents();
        sheet.Column(2).Width = 50; // Aufgabe
        sheet.Column(8).Width = 40; // Notizen
    }

    private void CreateListSheet(IXLWorksheet sheet, TodoList list, List<TodoTask> tasks)
    {
        // Titel
        sheet.Cell(1, 1).Value = list.DisplayName;
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        
        var completed = tasks.Count(t => t.Status == "completed");
        sheet.Cell(2, 1).Value = $"{tasks.Count} Aufgaben ({completed} erledigt)";
        sheet.Cell(2, 1).Style.Font.FontColor = XLColor.Gray;
        
        if (tasks.Count == 0)
        {
            sheet.Cell(4, 1).Value = "Keine Aufgaben in dieser Liste.";
            sheet.Columns().AdjustToContents();
            return;
        }
        
        var headers = new[] { "Status", "Aufgabe", "Priorität", "Fällig am", "Kategorien" };
        for (int i = 0; i < headers.Length; i++)
        {
            sheet.Cell(4, i + 1).Value = headers[i];
            sheet.Cell(4, i + 1).Style.Font.Bold = true;
            sheet.Cell(4, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#0078D4");
            sheet.Cell(4, i + 1).Style.Font.FontColor = XLColor.White;
        }
        
        int row = 5;
        foreach (var task in tasks)
        {
            sheet.Cell(row, 1).Value = task.Status == "completed" ? "✓" : "○";
            sheet.Cell(row, 2).Value = task.Title;
            sheet.Cell(row, 3).Value = task.Importance switch
            {
                "high" => "Hoch",
                "low" => "Niedrig",
                _ => "Normal"
            };
            
            if (task.DueDateTime != null && DateTime.TryParse(task.DueDateTime.DateTime, out var dueDate))
            {
                sheet.Cell(row, 4).Value = dueDate.ToString("dd.MM.yyyy");
            }
            
            sheet.Cell(row, 5).Value = string.Join(", ", task.Categories);
            
            // Status-Styling
            if (task.Status == "completed")
            {
                sheet.Cell(row, 1).Style.Font.FontColor = XLColor.FromHtml("#107C10");
                sheet.Range(row, 1, row, 5).Style.Font.FontColor = XLColor.Gray;
            }
            
            // Priorität-Farbe
            if (task.Importance == "high")
            {
                sheet.Cell(row, 3).Style.Font.FontColor = XLColor.FromHtml("#D13438");
                sheet.Cell(row, 3).Style.Font.Bold = true;
            }
            
            if (row % 2 == 1)
            {
                sheet.Range(row, 1, row, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#F5F5F5");
            }
            
            row++;
        }
        
        sheet.Columns().AdjustToContents();
        sheet.Column(2).Width = 50;
    }

    private string SanitizeSheetName(string name)
    {
        // Excel Sheet-Namen: max 31 Zeichen, keine: \ / ? * [ ] :
        var invalid = new[] { '\\', '/', '?', '*', '[', ']', ':' };
        foreach (var c in invalid)
        {
            name = name.Replace(c, '-');
        }
        return name.Length > 31 ? name.Substring(0, 31) : name;
    }

    /// <summary>
    /// Export im Todoist CSV-Format (wie im Original-Projekt)
    /// </summary>
    public string ExportToTodoistCsv(Dictionary<TodoList, List<TodoTask>> data)
    {
        var csv = new StringBuilder();
        csv.AppendLine("TYPE,CONTENT,PRIORITY,INDENT,AUTHOR,RESPONSIBLE,DATE,DATE_LANG,TIMEZONE");

        foreach (var (list, tasks) in data)
        {
            csv.AppendLine($"section,{EscapeCsvField(list.DisplayName)},,,,,,,");
            
            foreach (var task in tasks)
            {
                var priority = task.Importance?.ToLower() switch
                {
                    "high" => "1",
                    "normal" => "4",
                    "low" => "4",
                    _ => "4"
                };

                var date = task.DueDateTime?.DateTime ?? "";
                var timezone = task.DueDateTime?.TimeZone ?? "";

                csv.AppendLine($"task,{EscapeCsvField(task.Title)},{priority},,,,{date},en,{timezone}");
            }
        }

        return csv.ToString();
    }

    /// <summary>
    /// Export im erweiterten CSV-Format mit mehr Details
    /// </summary>
    public string ExportToDetailedCsv(Dictionary<TodoList, List<TodoTask>> data)
    {
        var csv = new StringBuilder();
        csv.AppendLine("Liste,Aufgabe,Status,Priorität,Fällig,Erstellt,Geändert,Kategorien,Notizen");

        foreach (var (list, tasks) in data)
        {
            foreach (var task in tasks)
            {
                var status = task.Status == "completed" ? "Erledigt" : "Offen";
                var priority = task.Importance switch
                {
                    "high" => "Hoch",
                    "low" => "Niedrig",
                    _ => "Normal"
                };
                var due = task.DueDateTime?.DateTime ?? "";
                var created = task.CreatedDateTime?.ToString("yyyy-MM-dd") ?? "";
                var modified = task.LastModifiedDateTime?.ToString("yyyy-MM-dd") ?? "";
                var categories = string.Join("; ", task.Categories);
                var notes = task.Body?.Content ?? "";

                csv.AppendLine($"{EscapeCsvField(list.DisplayName)},{EscapeCsvField(task.Title)},{status},{priority},{due},{created},{modified},{EscapeCsvField(categories)},{EscapeCsvField(notes)}");
            }
        }

        return csv.ToString();
    }

    /// <summary>
    /// Export als strukturiertes JSON
    /// </summary>
    public string ExportToJson(Dictionary<TodoList, List<TodoTask>> data)
    {
        var exportData = new
        {
            exportedAt = DateTime.UtcNow.ToString("o"),
            totalLists = data.Count,
            totalTasks = data.Values.Sum(t => t.Count),
            lists = data.Select(kvp => new
            {
                list = new
                {
                    id = kvp.Key.Id,
                    displayName = kvp.Key.DisplayName,
                    isOwner = kvp.Key.IsOwner,
                    isShared = kvp.Key.IsShared
                },
                taskCount = kvp.Value.Count,
                tasks = kvp.Value
            })
        };

        return JsonConvert.SerializeObject(exportData, Formatting.Indented);
    }

    /// <summary>
    /// Export als menschenlesbarer Text
    /// </summary>
    public string ExportToText(Dictionary<TodoList, List<TodoTask>> data)
    {
        var text = new StringBuilder();
        text.AppendLine("Microsoft To Do Export");
        text.AppendLine($"Exportiert am: {DateTime.Now:dd.MM.yyyy HH:mm}");
        text.AppendLine();

        foreach (var (list, tasks) in data)
        {
            text.AppendLine($"═══ {list.DisplayName} ({tasks.Count} Aufgaben) ═══");
            text.AppendLine();

            if (tasks.Count == 0)
            {
                text.AppendLine("  (Keine Aufgaben)");
                text.AppendLine();
                continue;
            }

            foreach (var task in tasks)
            {
                var status = task.Status == "completed" ? "✓" : "○";
                text.AppendLine($"  {status} {task.Title}");
                
                if (task.DueDateTime != null && !string.IsNullOrEmpty(task.DueDateTime.DateTime))
                {
                    if (DateTime.TryParse(task.DueDateTime.DateTime, out var dueDate))
                    {
                        text.AppendLine($"      Fällig: {dueDate:dd.MM.yyyy}");
                    }
                }
            }
            text.AppendLine();
        }

        return text.ToString();
    }

    /// <summary>
    /// Export als erweiterter lesbarer Text mit mehr Details
    /// </summary>
    public string ExportToDetailedText(Dictionary<TodoList, List<TodoTask>> data)
    {
        var text = new StringBuilder();
        text.AppendLine("╔════════════════════════════════════════════════════════════╗");
        text.AppendLine("║           Microsoft To Do Export                           ║");
        text.AppendLine($"║           {DateTime.Now:dd.MM.yyyy HH:mm}                               ║");
        text.AppendLine("╚════════════════════════════════════════════════════════════╝");
        text.AppendLine();

        var totalTasks = data.Values.Sum(t => t.Count);
        var completedTasks = data.Values.Sum(t => t.Count(task => task.Status == "completed"));
        text.AppendLine($"📊 {data.Count} Listen | {totalTasks} Aufgaben | {completedTasks} erledigt");
        text.AppendLine();

        foreach (var (list, tasks) in data)
        {
            text.AppendLine($"┌─────────────────────────────────────────────────────────────");
            text.AppendLine($"│ 📋 {list.DisplayName}");
            text.AppendLine($"│    {tasks.Count} Aufgaben");
            text.AppendLine($"└─────────────────────────────────────────────────────────────");

            if (tasks.Count == 0)
            {
                text.AppendLine("    (Keine Aufgaben)");
                text.AppendLine();
                continue;
            }

            foreach (var task in tasks)
            {
                var statusIcon = task.Status == "completed" ? "✅" : "⬜";
                var importanceIcon = task.Importance switch
                {
                    "high" => " 🔴",
                    "low" => " 🟢",
                    _ => ""
                };

                text.AppendLine($"    {statusIcon} {task.Title}{importanceIcon}");

                if (!string.IsNullOrWhiteSpace(task.Body?.Content))
                {
                    text.AppendLine($"       📝 {task.Body.Content.Replace("\n", " ").Trim()}");
                }

                if (task.DueDateTime != null && !string.IsNullOrEmpty(task.DueDateTime.DateTime))
                {
                    if (DateTime.TryParse(task.DueDateTime.DateTime, out var dueDate))
                    {
                        text.AppendLine($"       📅 Fällig: {dueDate:dd.MM.yyyy}");
                    }
                }

                if (task.Categories.Count > 0)
                {
                    text.AppendLine($"       🏷️  {string.Join(", ", task.Categories)}");
                }

                if (task.ChecklistItems != null && task.ChecklistItems.Count > 0)
                {
                    foreach (var item in task.ChecklistItems)
                    {
                        var checkIcon = item.IsChecked == true ? "☑" : "☐";
                        text.AppendLine($"          {checkIcon} {item.DisplayName}");
                    }
                }
            }

            text.AppendLine();
        }

        return text.ToString();
    }

    private string EscapeCsvField(string field)
    {
        if (string.IsNullOrEmpty(field))
            return "";

        field = field.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");

        if (field.Contains(',') || field.Contains('"') || field.Contains('\n'))
        {
            field = field.Replace("\"", "\"\"");
            field = $"\"{field}\"";
        }

        return field;
    }
}
