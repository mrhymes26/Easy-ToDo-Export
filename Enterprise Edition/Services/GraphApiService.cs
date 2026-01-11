using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using TodoExport.Models;

namespace TodoExport.Services;

public class GraphApiService : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly AuthService _authService;
    private const string GraphApiBaseUrl = "https://graph.microsoft.com/v1.0";

    public GraphApiService(AuthService authService)
    {
        _httpClient = new HttpClient();
        _authService = authService;
    }

    private async Task EnsureAuthenticatedAsync()
    {
        var token = await _authService.GetAccessTokenAsync();
        if (string.IsNullOrEmpty(token))
        {
            throw new Exception("Nicht angemeldet. Bitte melden Sie sich zuerst an.");
        }
        
        // Entferne vorherigen Authorization-Header falls vorhanden
        _httpClient.DefaultRequestHeaders.Authorization = null;
        _httpClient.DefaultRequestHeaders.Authorization = 
            new AuthenticationHeaderValue("Bearer", token);
    }

    public async Task<List<TodoList>> GetTodoListsAsync()
    {
        await EnsureAuthenticatedAsync();
        
        var allLists = new List<TodoList>();
        // Verwende die Standard-API ohne zusätzliche Parameter
        // Microsoft Graph API paginiert automatisch mit @odata.nextLink
        string? nextLink = $"{GraphApiBaseUrl}/me/todo/lists";
        int pageCount = 0;

        while (!string.IsNullOrEmpty(nextLink))
        {
            pageCount++;
            System.Diagnostics.Debug.WriteLine($"=== Lade Seite {pageCount}: {nextLink} ===");
            
            // Stelle sicher, dass der Authorization-Header für jeden Request gesetzt ist
            await EnsureAuthenticatedAsync();
            
            // Verwende HttpRequestMessage für bessere Kontrolle über den Request
            var request = new HttpRequestMessage(HttpMethod.Get, nextLink);
            var response = await _httpClient.SendAsync(request);
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"API-Fehler: {response.StatusCode} - {error}");
                throw new Exception($"API-Fehler: {response.StatusCode} - {error}");
            }

            var content = await response.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine($"=== Vollständige API-Antwort ===");
            System.Diagnostics.Debug.WriteLine(content);
            System.Diagnostics.Debug.WriteLine($"=== Ende API-Antwort ===");
            
            var json = JObject.Parse(content);

            var listsArray = json["value"];
            if (listsArray != null)
            {
                var lists = listsArray.ToObject<List<TodoList>>() ?? new List<TodoList>();
                
                // Debug: Log alle Listen mit ihren wellknownListName-Werten
                System.Diagnostics.Debug.WriteLine($"=== Seite {pageCount}: {lists.Count} Listen gefunden ===");
                foreach (var list in lists)
                {
                    System.Diagnostics.Debug.WriteLine($"  - '{list.DisplayName}' (ID: {list.Id}, Wellknown: '{list.WellknownListName ?? "null"}', IsOwner: {list.IsOwner}, IsShared: {list.IsShared})");
                }
                
                allLists.AddRange(lists);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"WARNUNG: 'value' Array ist null in der Antwort!");
            }

            // Prüfe auf @odata.nextLink für Paginierung
            // Microsoft Graph API kann sowohl @odata.nextLink als auch @odata.deltaLink verwenden
            nextLink = json["@odata.nextLink"]?.ToString();
            
            // Falls @odata.nextLink nicht vorhanden, prüfe ob es mehr Ergebnisse gibt
            // Standard-Paginierung: Wenn weniger als die maximale Anzahl zurückgegeben wird, gibt es keine weitere Seite
            if (string.IsNullOrEmpty(nextLink))
            {
                System.Diagnostics.Debug.WriteLine($"Keine @odata.nextLink gefunden.");
                
                // Prüfe auch auf @odata.context und andere Metadaten
                System.Diagnostics.Debug.WriteLine($"Antwort-Metadaten:");
                foreach (var prop in json.Properties())
                {
                    if (prop.Name.StartsWith("@odata"))
                    {
                        System.Diagnostics.Debug.WriteLine($"  {prop.Name}: {prop.Value}");
                    }
                }
                
                // Prüfe ob es möglicherweise eine @odata.count gibt, die die Gesamtanzahl angibt
                var count = json["@odata.count"]?.ToString();
                if (!string.IsNullOrEmpty(count))
                {
                    System.Diagnostics.Debug.WriteLine($"Gesamtanzahl laut @odata.count: {count}, Aktuell geladen: {allLists.Count}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Nächste Seite vorhanden: {nextLink}");
            }
        }

        System.Diagnostics.Debug.WriteLine($"=== Gesamt: {allLists.Count} Listen vor Filterung ===");
        
        // Debug: Zeige alle wellknownListName-Werte
        var wellknownValues = allLists
            .Select(l => l.WellknownListName ?? "null")
            .Distinct()
            .ToList();
        System.Diagnostics.Debug.WriteLine($"Gefundene wellknownListName-Werte: {string.Join(", ", wellknownValues)}");
        
        // Filter: Alle benutzerdefinierten Listen zurückgeben
        // wellknownListName ist null/leer für benutzerdefinierte Listen
        // "none" bedeutet auch benutzerdefinierte Liste
        // "defaultList" ist die Standard-Liste
        // Andere Werte wie "flaggedEmails" sind spezielle System-Listen, die wir normalerweise nicht exportieren wollen
        var filteredLists = allLists
            .Where(l => string.IsNullOrEmpty(l.WellknownListName) || 
                       l.WellknownListName == "none" || 
                       l.WellknownListName == "defaultList")
            .ToList();
        
        System.Diagnostics.Debug.WriteLine($"=== Nach Filterung: {filteredLists.Count} Listen ===");
        
        // Wenn nach Filterung weniger Listen vorhanden sind als vorher,
        // zeige welche Listen herausgefiltert wurden
        if (filteredLists.Count < allLists.Count)
        {
            var filteredOut = allLists.Except(filteredLists).ToList();
            System.Diagnostics.Debug.WriteLine($"=== {filteredOut.Count} Listen wurden herausgefiltert ===");
            foreach (var list in filteredOut)
            {
                System.Diagnostics.Debug.WriteLine($"  - '{list.DisplayName}' (Wellknown: '{list.WellknownListName ?? "null"}')");
            }
        }
        
        // TEMPORÄR: Gebe ALLE Listen zurück, um zu sehen, was die API tatsächlich zurückgibt
        // TODO: Filter später wieder aktivieren, wenn wir wissen, welche Listen gefiltert werden sollen
        if (allLists.Count > 0)
        {
            System.Diagnostics.Debug.WriteLine($"=== TEMPORÄR: Gebe alle {allLists.Count} Listen zurück (Filter deaktiviert für Debugging) ===");
            return allLists;
        }
        
        return filteredLists;
    }

    public async Task<List<TodoTask>> GetTasksAsync(string listId)
    {
        await EnsureAuthenticatedAsync();
        
        var allTasks = new List<TodoTask>();
        string? nextLink = $"{GraphApiBaseUrl}/me/todo/lists/{listId}/tasks";

        while (!string.IsNullOrEmpty(nextLink))
        {
            var response = await _httpClient.GetAsync(nextLink);
            response.EnsureSuccessStatusCode();

            var content = await response.Content.ReadAsStringAsync();
            var json = JObject.Parse(content);

            var tasksArray = json["value"];
            if (tasksArray != null)
            {
                var tasks = tasksArray.ToObject<List<TodoTask>>() ?? new List<TodoTask>();
                allTasks.AddRange(tasks);
            }

            nextLink = json["@odata.nextLink"]?.ToString();
        }

        return allTasks;
    }

    public async Task<Dictionary<string, List<TodoTask>>> GetAllTasksAsync(List<TodoList> lists)
    {
        var allTasks = new Dictionary<string, List<TodoTask>>();

        foreach (var list in lists)
        {
            try
            {
                var tasks = await GetTasksAsync(list.Id);
                allTasks[list.Id] = tasks;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fehler bei '{list.DisplayName}': {ex.Message}");
                allTasks[list.Id] = new List<TodoTask>();
            }
        }

        return allTasks;
    }

    public void Dispose()
    {
        _httpClient?.Dispose();
    }
}
