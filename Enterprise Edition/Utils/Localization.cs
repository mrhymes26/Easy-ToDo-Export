namespace TodoExport.Utils;

/// <summary>
/// Lokalisierung für DE/EN
/// </summary>
public static class Lang
{
    public static string CurrentLanguage { get; set; } = "en"; // Default: Englisch

    public static string Get(string key)
    {
        return CurrentLanguage == "de" ? GetGerman(key) : GetEnglish(key);
    }

    private static string GetGerman(string key) => key switch
    {
        // Header
        "AppTitle" => "Microsoft To Do Export",
        "AppSubtitle" => "Sichern Sie Ihre Aufgaben in verschiedenen Formaten",
        "LoggedIn" => "Angemeldet",
        "SignOut" => "Abmelden",
        
        // Login Section
        "SignInTitle" => "Mit Microsoft anmelden",
        "SignInSubtitle" => "Melden Sie sich an, um Ihre To-Do-Listen\nzu sichern und zu exportieren.",
        "SignInButton" => "Anmelden",
        "SignInHint" => "Unterstützt Outlook.de, Hotmail, Live und Firmenkonten",
        
        // Device Code Section
        "BrowserLogin" => "Browser-Anmeldung",
        "BrowserLoginHint" => "Ein Browser wurde geöffnet. Geben Sie dort\ndiesen Code ein, um sich anzumelden:",
        "OpenBrowser" => "Browser erneut öffnen",
        "Cancel" => "Abbrechen",
        "CopyCode" => "Code kopieren",
        
        // Export Section
        "SelectLists" => "Listen auswählen",
        "All" => "Alle",
        "None" => "Keine",
        "ExportOptions" => "Export-Optionen",
        "Selected" => "Ausgewählt",
        "List" => "Liste",
        "Lists" => "Listen",
        "Format" => "Format",
        "IncludeCompleted" => "Abgeschlossene Aufgaben einschließen",
        "Export" => "Exportieren",
        "SelectListHint" => "Wählen Sie mindestens eine Liste aus",
        "RefreshLists" => "Listen neu laden",
        "HideCompleted" => "Abgeschlossene ausblenden",
        
        // Format Descriptions
        "FormatExcel" => "Formatierte Tabelle mit Übersicht und Farben",
        "FormatTodoistCsv" => "Kompatibel mit Todoist Import-Funktion",
        "FormatJson" => "Strukturiertes Format für Entwickler",
        "FormatText" => "Einfache Textliste zum Lesen",
        "FormatDetailedText" => "Ausführliche Informationen pro Aufgabe",
        "FormatDetailedCsv" => "Alle Details als Tabelle",
        
        // Status Messages
        "Ready" => "Bereit",
        "CheckingLogin" => "Prüfe Anmeldung...",
        "StartingLogin" => "Starte Anmeldung...",
        "WaitingForBrowser" => "Warte auf Browser-Anmeldung...",
        "LoginCancelled" => "Anmeldung abgebrochen",
        "LoginFailed" => "Anmeldung fehlgeschlagen",
        "LoadingLists" => "Lade Listen...",
        "ListsLoaded" => "Listen geladen",
        "LoadingListsError" => "Fehler beim Laden",
        "Exporting" => "Exportiere Aufgaben...",
        "ExportFailed" => "Export fehlgeschlagen",
        "ExportSuccess" => "Aufgaben exportiert",
        "ExportComplete" => "Export erfolgreich!",
        "ExportCompleteDetails" => "Datei im Explorer anzeigen?",
        "SignedOut" => "Abgemeldet",
        
        // Export Messages
        "ExportSuccessMessage" => "Export erfolgreich!\n\n",
        "ExportFile" => "Datei",
        "ExportLists" => "Listen",
        "ExportTasks" => "Aufgaben",
        
        // Errors
        "Error" => "Fehler",
        "PleaseSelectList" => "Bitte wählen Sie mindestens eine Liste aus.",
        "Hint" => "Hinweis",
        
        _ => key
    };

    private static string GetEnglish(string key) => key switch
    {
        // Header
        "AppTitle" => "Microsoft To Do Export",
        "AppSubtitle" => "Back up your tasks in various formats",
        "LoggedIn" => "Signed in",
        "SignOut" => "Sign out",
        
        // Login Section
        "SignInTitle" => "Sign in with Microsoft",
        "SignInSubtitle" => "Sign in to back up and export\nyour To-Do lists.",
        "SignInButton" => "Sign in",
        "SignInHint" => "Supports Outlook.com, Hotmail, Live, and work accounts",
        
        // Device Code Section
        "BrowserLogin" => "Browser sign-in",
        "BrowserLoginHint" => "A browser has been opened. Enter\nthis code there to sign in:",
        "OpenBrowser" => "Open browser again",
        "Cancel" => "Cancel",
        "CopyCode" => "Copy code",
        
        // Export Section
        "SelectLists" => "Select lists",
        "All" => "All",
        "None" => "None",
        "ExportOptions" => "Export options",
        "Selected" => "Selected",
        "List" => "list",
        "Lists" => "lists",
        "Format" => "Format",
        "IncludeCompleted" => "Include completed tasks",
        "Export" => "Export",
        "SelectListHint" => "Select at least one list",
        "RefreshLists" => "Refresh lists",
        "HideCompleted" => "Hide completed",
        
        // Format Descriptions
        "FormatExcel" => "Formatted table with overview and colors",
        "FormatTodoistCsv" => "Compatible with Todoist import function",
        "FormatJson" => "Structured format for developers",
        "FormatText" => "Simple text list for reading",
        "FormatDetailedText" => "Detailed information per task",
        "FormatDetailedCsv" => "All details as table",
        
        // Status Messages
        "Ready" => "Ready",
        "CheckingLogin" => "Checking sign-in...",
        "StartingLogin" => "Starting sign-in...",
        "WaitingForBrowser" => "Waiting for browser sign-in...",
        "LoginCancelled" => "Sign-in cancelled",
        "LoginFailed" => "Sign-in failed",
        "LoadingLists" => "Loading lists...",
        "ListsLoaded" => "lists loaded",
        "LoadingListsError" => "Error loading",
        "Exporting" => "Exporting tasks...",
        "ExportFailed" => "Export failed",
        "ExportSuccess" => "tasks exported",
        "ExportComplete" => "Export complete!",
        "ExportCompleteDetails" => "Show file in Explorer?",
        "SignedOut" => "Signed out",
        
        // Export Messages
        "ExportSuccessMessage" => "Export successful!\n\n",
        "ExportFile" => "File",
        "ExportLists" => "Lists",
        "ExportTasks" => "Tasks",
        
        // Errors
        "Error" => "Error",
        "PleaseSelectList" => "Please select at least one list.",
        "Hint" => "Hint",
        
        _ => key
    };
}

