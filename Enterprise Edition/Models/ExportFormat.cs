namespace TodoExport.Models;

public enum ExportFormat
{
    /// <summary>
    /// Todoist CSV-Format (Original-Format)
    /// </summary>
    TodoistCsv,
    
    /// <summary>
    /// Strukturiertes JSON mit Metadaten
    /// </summary>
    Json,
    
    /// <summary>
    /// Menschenlesbarer Text (Original-Format)
    /// </summary>
    Text,
    
    /// <summary>
    /// Erweiterter Text mit mehr Details
    /// </summary>
    DetailedText,
    
    /// <summary>
    /// Erweitertes CSV mit mehr Details
    /// </summary>
    DetailedCsv,
    
    /// <summary>
    /// Formatierte Excel-Datei (.xlsx)
    /// </summary>
    Excel
}
