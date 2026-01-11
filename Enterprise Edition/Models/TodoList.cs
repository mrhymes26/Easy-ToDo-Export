using Newtonsoft.Json;

namespace TodoExport.Models;

public class TodoList
{
    [JsonProperty("id")]
    public string Id { get; set; } = string.Empty;

    [JsonProperty("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonProperty("isOwner")]
    public bool? IsOwner { get; set; }

    [JsonProperty("isShared")]
    public bool? IsShared { get; set; }

    [JsonProperty("wellknownListName")]
    public string? WellknownListName { get; set; }
}

public class TodoListResponse
{
    [JsonProperty("value")]
    public List<TodoList> Value { get; set; } = new();
}























