using Newtonsoft.Json;

namespace TodoExport.Models;

public class TodoTask
{
    [JsonProperty("id")]
    public string Id { get; set; } = string.Empty;

    [JsonProperty("title")]
    public string Title { get; set; } = string.Empty;

    [JsonProperty("body")]
    public TaskBody? Body { get; set; }

    [JsonProperty("status")]
    public string Status { get; set; } = "notStarted";

    [JsonProperty("importance")]
    public string Importance { get; set; } = "normal";

    [JsonProperty("createdDateTime")]
    public DateTimeOffset? CreatedDateTime { get; set; }

    [JsonProperty("lastModifiedDateTime")]
    public DateTimeOffset? LastModifiedDateTime { get; set; }

    [JsonProperty("dueDateTime")]
    public DateTimeTimeZone? DueDateTime { get; set; }

    [JsonProperty("completedDateTime")]
    public DateTimeTimeZone? CompletedDateTime { get; set; }

    [JsonProperty("reminderDateTime")]
    public DateTimeTimeZone? ReminderDateTime { get; set; }

    [JsonProperty("categories")]
    public List<string> Categories { get; set; } = new();

    [JsonProperty("hasAttachments")]
    public bool? HasAttachments { get; set; }

    [JsonProperty("checklistItems")]
    public List<ChecklistItem>? ChecklistItems { get; set; }
}

public class TaskBody
{
    [JsonProperty("content")]
    public string Content { get; set; } = string.Empty;

    [JsonProperty("contentType")]
    public string ContentType { get; set; } = "text";
}

public class DateTimeTimeZone
{
    [JsonProperty("dateTime")]
    public string DateTime { get; set; } = string.Empty;

    [JsonProperty("timeZone")]
    public string TimeZone { get; set; } = "UTC";
}

public class ChecklistItem
{
    [JsonProperty("id")]
    public string Id { get; set; } = string.Empty;

    [JsonProperty("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonProperty("isChecked")]
    public bool? IsChecked { get; set; }
}

public class TodoTaskResponse
{
    [JsonProperty("value")]
    public List<TodoTask> Value { get; set; } = new();
}




