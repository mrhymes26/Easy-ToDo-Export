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
        
        _httpClient.DefaultRequestHeaders.Authorization = 
            new AuthenticationHeaderValue("Bearer", token);
    }

    public async Task<List<TodoList>> GetTodoListsAsync()
    {
        await EnsureAuthenticatedAsync();
        
        var response = await _httpClient.GetAsync($"{GraphApiBaseUrl}/me/todo/lists");
        
        if (!response.IsSuccessStatusCode)
        {
            var error = await response.Content.ReadAsStringAsync();
            throw new Exception($"API-Fehler: {response.StatusCode}");
        }

        var content = await response.Content.ReadAsStringAsync();
        var listResponse = JsonConvert.DeserializeObject<TodoListResponse>(content);

        return listResponse?.Value
            .Where(l => string.IsNullOrEmpty(l.WellknownListName) || 
                       l.WellknownListName == "none" || 
                       l.WellknownListName == "defaultList")
            .ToList() ?? new List<TodoList>();
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
