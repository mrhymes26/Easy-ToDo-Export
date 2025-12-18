# Microsoft To Do Export

A user-friendly Windows application to backup and export your Microsoft To Do lists to various formats including Excel, CSV, JSON, and plain text.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-lightgrey.svg)
![.NET](https://img.shields.io/badge/.NET-8.0-purple.svg)

## Background

This project was inspired by [Microsoft-To-Do-Export](https://github.com/daylamtayari/Microsoft-To-Do-Export) and [TodoVodo](https://www.todovodo.com/start-tutorial). However, these solutions didn't meet my requirements for a simple, easy-to-use application that saves time on post-export work. I wanted something that:

- **Just works** – No manual token handling or complex setup
- **Remembers your login** – No need to authenticate every time
- **Exports to Excel** – Formatted, colored spreadsheets ready to use
- **Lets you select multiple lists** – Export exactly what you need
- **Works on Windows without using command-line scripts**

## Features

✅ **One-Click Authentication** – Browser opens automatically, just sign in with your Microsoft account  
✅ **Persistent Login** – Your session is saved, no re-authentication needed on next launch  
✅ **Multiple Export Formats:**
- **Excel (.xlsx)** – Formatted tables with colors, summary sheet, and separate sheets per list
- **Todoist CSV** – Compatible with Todoist import
- **JSON** – Structured data for developers
- **Plain Text** – Human-readable task lists
- **Detailed Text** – Extended information with notes and categories
- **Detailed CSV** – All task properties as a table

✅ **Multi-Select Lists** – Choose which lists to export with checkboxes  
✅ **Modern UI** – Clean, intuitive Windows interface  
✅ **Windows 10 & 11** – Full compatibility

## Screenshots

### Login Screen
Simple one-click authentication – a browser window opens automatically.

### Export Screen
Select your lists, choose a format, and export with one click.

## Installation

### Option 1: Download Release
1. Go to [Releases](../../releases)
2. Download the latest `TodoExport-vX.X.X.zip`
3. Extract and run `TodoExport.exe`

### Option 2: Build from Source
```bash
git clone https://github.com/YOUR_USERNAME/TodoExport.git
cd TodoExport
dotnet build --configuration Release
```

The compiled application will be in `bin/Release/net8.0-windows/`.

## Requirements

- Windows 10 or Windows 11
- .NET 8.0 Runtime ([Download](https://dotnet.microsoft.com/download/dotnet/8.0))
- Microsoft account (Outlook.com, Hotmail, Live, or Work/School account)

## Usage

1. **Launch the application**
2. **Click "Anmelden" (Sign In)** – A browser window opens
3. **Enter the displayed code** in the browser and sign in with your Microsoft account
4. **Select your lists** using the checkboxes
5. **Choose an export format** (Excel is recommended)
6. **Click "Exportieren"** and choose where to save

Your login is saved automatically – next time, you'll be signed in instantly!

## Export Formats Explained

| Format | Best For | Description |
|--------|----------|-------------|
| **Excel (.xlsx)** | Most users | Formatted spreadsheet with colors, summary, and per-list sheets |
| **Todoist CSV** | Migrating to Todoist | Compatible with Todoist's import function |
| **JSON** | Developers | Structured data with all task properties |
| **Text** | Quick overview | Simple human-readable list |
| **Detailed Text** | Archiving | Extended info with notes, dates, categories |
| **Detailed CSV** | Data analysis | All properties in tabular format |

## Privacy & Security

- This app uses **Microsoft's official authentication** (MSAL)
- Your credentials are **never stored** by this app
- Only an authentication token is cached locally in `%LocalAppData%\TodoExport\`
- The app only requests **read access** to your To Do lists (`Tasks.ReadWrite`)
- No data is sent to any third-party servers

## Technical Details

- **Framework:** .NET 8.0 WPF
- **Authentication:** MSAL (Microsoft Authentication Library) with Device Code Flow
- **Excel Generation:** ClosedXML
- **API:** Microsoft Graph API

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Microsoft-To-Do-Export](https://github.com/daylamtayari/Microsoft-To-Do-Export) – Original inspiration
- [TodoVoodoo](https://github.com/nickvdyck/todo-voodoo) – Additional inspiration
- [Microsoft Graph API](https://docs.microsoft.com/graph/) – API documentation
- [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet) – Authentication library
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) – Excel generation

---

Made with ❤️ for everyone who wants a simple way to backup their tasks.
