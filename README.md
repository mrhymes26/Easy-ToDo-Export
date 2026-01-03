# Microsoft To Do Export

A user-friendly Windows application to backup and export your Microsoft To Do lists to various formats including Excel, CSV, JSON, and plain text.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-lightgrey.svg)
![.NET](https://img.shields.io/badge/.NET-8.0-purple.svg)

## Background

This project was inspired by [Microsoft-To-Do-Export](https://github.com/daylamtayari/Microsoft-To-Do-Export) and [TodoVodo](https://www.todovodo.com). However, these solutions didn't meet my requirements for a simple, easy-to-use application that saves time on post-export work. I wanted something that:

- **Just works** – No manual token handling or complex setup
- **Remembers your login** – No need to authenticate every time
- **Exports to Excel** – Formatted, colored spreadsheets ready to use
- **Lets you select multiple lists** – Export exactly what you need
- **Works on Windows without using command-line scripts**

## Features

- ✅ **Easy Authentication** - Sign in with your Microsoft account using Device Code Flow (no manual token input required)
- 🔐 **Secure Token Caching** - Your login is remembered using Windows DPAPI encryption
- 🌐 **Multi-Language Support** - Switch between German (DE) and English (EN) with a simple toggle
- 📋 **Multi-Select Lists** - Select multiple lists at once with checkboxes
- 📊 **Multiple Export Formats**:
  - **Excel (.xlsx)** - Formatted spreadsheet with overview, colors, and separate sheets per list
  - **Todoist CSV (.csv)** - Compatible with Todoist import function
  - **JSON (.json)** - Structured format for developers
  - **Text (.txt)** - Simple text list for reading
  - **Detailed Text (.txt)** - Detailed information per task with icons
  - **Detailed CSV (.csv)** - All details as a table
- 🎨 **Modern UI** - User-friendly Windows 11-like interface
- ⚡ **Pagination Support** - Handles lists with many tasks automatically
- ✅ **Include/Exclude Completed Tasks** - Option to include or exclude completed tasks during export
- 🔄 **Auto-Refresh** - Refresh your lists with a single click

## Screenshots

### Login Screen
Simple one-click authentication – a browser window opens automatically.

### Export Screen
Select your lists, choose a format, and export with one click. Switch between German and English with the language toggle in the top-right corner.

## Requirements

- Windows 10/11
- Microsoft account with Microsoft To Do

**Note:** No additional installation required! The application includes everything needed to run.

## Installation

### Option 1: Download Release
1. Go to [Releases](../../releases) (or download from the latest GitHub release)
2. Download the latest Release
3. Extract all files to a folder
4. Run `TodoExport.exe` - **No installation required!**

The application is self-contained and includes everything needed to run on Windows 10/11.

### Option 2: Build from Source
```bash
git clone https://github.com/YOUR_USERNAME/TodoExport.git
cd 2025-app-todo-export
dotnet publish -c Release -r win-x64 --self-contained true
```

The compiled application will be in `builds/self-contained/` or `bin/Release/net8.0-windows/win-x64/`.

## Usage

### 1. Sign In

1. Click **"Sign in"** button
2. A browser window will open automatically
3. Enter the code shown in the application
4. Sign in with your Microsoft account
5. Grant the required permissions
6. You're signed in! Your login will be remembered for future sessions

### 2. Select Lists

1. All your To-Do lists will be loaded automatically
2. Use **"All"** to select all lists or **"None"** to deselect all
3. Check/uncheck individual lists as needed
4. The selected count is displayed in real-time

### 3. Configure Export

1. Choose your export format from the dropdown
2. **(Optional)** Check/uncheck **"Include completed tasks"** to include or exclude completed tasks
3. Click **"Export"**
4. Choose the save location for your export file
5. The export will be performed and you'll receive a confirmation

## Export Formats

### 📊 Excel (.xlsx)
Formatted spreadsheet with:
- Overview sheet with summary statistics
- "All Tasks" sheet with all tasks from all selected lists
- Separate sheet for each To-Do list
- Color-coded status indicators
- Formatted dates and times

### 📊 Todoist CSV (.csv)
CSV format compatible with Todoist:
```
TYPE,CONTENT,PRIORITY,INDENT,AUTHOR,RESPONSIBLE,DATE,DATE_LANG,TIMEZONE
section,List Name,,,,,,,
task,Task Title,4,,,,2024-12-25,en,UTC
```

### 📊 Detailed CSV (.csv)
Detailed CSV format with all available fields:
```
TYPE,CONTENT,NOTE,PROJECT,LABELS,PRIORITY,INDENT,AUTHOR,RESPONSIBLE,DATE,DATE_LANG,TIMEZONE,STATUS,CREATED,MODIFIED
```

### 📄 JSON (.json)
Complete JSON format with metadata:
```json
{
  "about": "File exported using Microsoft To Do Export",
  "exportedAt": "2024-12-18T12:00:00Z",
  "lists": [
    {
      "list": { "id": "...", "displayName": "..." },
      "tasks": [...]
    }
  ]
}
```

### 📝 Text (.txt)
Simple tab-separated format:
```
	Content	Due Date	Timezone
List Name
	Task Title	2024-12-25	UTC
```

### 📝 Detailed Text (.txt)
Human-readable format with icons:
```
Microsoft To Do Export
Exported on: 2024-12-18 12:00:00
============================================================

📋 List Name
------------------------------------------------------------
  ○ Task Title 🔴
     📅 Due: 2024-12-25 12:00
     🏷️  Categories: Work, Important
```

## Export Formats Explained

| Format | Best For | Description |
|--------|----------|-------------|
| **Excel (.xlsx)** | Most users | Formatted spreadsheet with colors, summary, and per-list sheets |
| **Todoist CSV** | Migrating to Todoist | Compatible with Todoist's import function |
| **JSON** | Developers | Structured data with all task properties |
| **Text** | Quick overview | Simple human-readable list |
| **Detailed Text** | Archiving | Extended info with notes, dates, categories |
| **Detailed CSV** | Data analysis | All properties in tabular format |

## Language Support

The application supports two languages:
- **English (EN)** - Default language
- **German (DE)** - Full German translation

Switch between languages using the **DE/EN** toggle in the top-right corner of the header. All UI elements, messages, and dialogs are translated.

## Technical Details

- **Framework:** .NET 8.0 WPF
- **Authentication:** MSAL (Microsoft Authentication Library) with Device Code Flow
- **Excel Generation:** ClosedXML
- **API:** Microsoft Graph API
- **JSON Serialization:** Newtonsoft.Json

### Architecture

The application follows a service-oriented architecture:

```
2025-app-todo-export/
├── Models/
│   ├── TodoList.cs      # List model
│   ├── TodoTask.cs      # Task model
│   └── ExportFormat.cs  # Export format enum
├── Services/
│   ├── AuthService.cs       # Microsoft authentication (MSAL)
│   ├── GraphApiService.cs    # Microsoft Graph API client
│   └── ExportService.cs      # Export functions
├── Utils/
│   └── Localization.cs      # Multi-language support
├── MainWindow.xaml          # UI
├── MainWindow.xaml.cs       # Code-behind
├── App.xaml                 # App resources
└── App.xaml.cs
```

### Microsoft Graph API Endpoints

- `GET /me/todo/lists` - Retrieve all lists
- `GET /me/todo/lists/{listId}/tasks` - Retrieve tasks for a list

The application supports pagination via `@odata.nextLink` for lists with many tasks.

### Authentication

The application uses **Device Code Flow** for authentication:
- No manual token input required
- Secure token caching with Windows DPAPI encryption
- Automatic token refresh
- Supports personal Microsoft accounts (Outlook.com, Hotmail, Live) and work accounts

## Troubleshooting

### "Sign-in failed"
- Make sure you have an active internet connection
- Check if your Microsoft account has access to Microsoft To Do
- Try signing out and signing in again

### "No lists found"
- Ensure you're using Microsoft To Do
- Check if you're signed in with the correct account
- System lists (Flagged Email, etc.) are automatically filtered out

### Export fails
- Ensure you have write permissions for the chosen location
- Check if sufficient disk space is available
- Try selecting fewer lists if you have many tasks

### Token cache issues
- If you experience persistent login issues, you can clear the token cache
- The cache is stored in: `%LocalAppData%\TodoExport\`
- Delete the `todo_export_cache.bin` file to reset authentication

## Privacy & Security

- This app uses **Microsoft's official authentication** (MSAL)
- Your credentials are **never stored** by this app
- Only an authentication token is cached locally in `%LocalAppData%\TodoExport\`
- The app only requests **read-only permissions** - no modifications are made:
  - `Tasks.Read` - Read access to your To Do lists (for export only)
  - `User.Read` - Read your basic profile (to display your name)
  - `offline_access` - Refresh token for automatic re-login (optional)
- No data is sent to any third-party servers
- **Token Storage**: Access tokens are encrypted using Windows DPAPI (Data Protection API)
- **Local Storage Only**: All data is stored locally on your machine
- **Token Cache**: Encrypted token cache file is user-specific and machine-specific

## Development

### Build

```bash
cd 2025-app-todo-export
dotnet build
```

### Release Build

```bash
dotnet build --configuration Release
```

### Publish (Self-contained - Recommended)

```bash
dotnet publish -c Release -r win-x64 --self-contained true
```

The self-contained build includes the .NET 8.0 runtime and works on any Windows 10/11 system without additional installation.

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

If you find this project useful and would like to support its development, consider buying me a coffee! ☕

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://buymeacoffee.com/mrhymes)

[Support on Buy Me a Coffee](https://buymeacoffee.com/mrhymes)

## Acknowledgments

- [Microsoft-To-Do-Export](https://github.com/daylamtayari/Microsoft-To-Do-Export) – Original inspiration
- [TodoVoodoo](https://github.com/nickvdyck/todo-voodoo) – Additional inspiration
- [Microsoft Graph API](https://docs.microsoft.com/graph/) – API documentation
- [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet) – Authentication library
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) – Excel generation

---

Made with ❤️ for everyone who wants a simple way to backup their tasks.

**Note**: This project is an unofficial solution and is not connected to Microsoft or Todoist.
