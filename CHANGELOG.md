# Changelog

All notable changes to this project will be documented in this file.

## [1.0.1] - 2025-01-XX

### Changed
- Improved authentication error handling
- Enhanced token caching reliability
- Updated UI text localization

### Fixed
- Fixed token refresh issues in GetAccessTokenAsync
- Improved error messages for better user experience

## [1.0.0] - 2025-12-18

### Added
- Initial release
- One-click Microsoft authentication with Device Code Flow
- Persistent login (token caching with Windows DPAPI encryption)
- Multi-select list export with checkboxes
- Option to include/exclude completed tasks
- Language switch (English/German) with automatic system language detection
- Export formats:
  - Excel (.xlsx) with formatted tables, colors, and summary sheet
  - Todoist CSV (compatible with Todoist import)
  - JSON (structured data)
  - Plain Text (human-readable)
  - Detailed Text (with notes and categories)
  - Detailed CSV (all properties)
- "Select All" / "Select None" buttons for quick list selection
- Modern, clean Windows UI
- Support for Windows 10 and Windows 11
- Support for personal Microsoft accounts (Outlook.com, Hotmail, Live)
- Support for Work/School accounts

### Technical
- Built with .NET 8.0 WPF
- Uses MSAL for authentication
- Uses Microsoft Graph API for data access
- Uses ClosedXML for Excel generation

