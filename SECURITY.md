# Security

## Overview

This document describes the security architecture of TodoExport and how your data is protected.

## What Data is Stored?

| Data | Stored? | Location | Protection |
|------|---------|----------|------------|
| **Password** | ❌ Never | - | Your password is never stored or transmitted through this app |
| **Access Token** | ✅ Yes | Local cache | Windows DPAPI encryption |
| **Refresh Token** | ✅ Yes | Local cache | Windows DPAPI encryption |
| **Email Address** | ✅ Yes | Local cache | Windows DPAPI encryption |
| **Tasks/Lists** | ❌ No | Only during export | Not persisted |

## Token Cache Location

```
%LocalAppData%\TodoExport\todo_export_cache.bin
```

On Windows, this is typically:
```
C:\Users\<YourUsername>\AppData\Local\TodoExport\todo_export_cache.bin
```

## How Tokens are Protected

### Windows (Primary Platform)
- **Windows DPAPI (Data Protection API)** is used to encrypt the token cache
- The encryption key is tied to your Windows user account
- Only your Windows user can decrypt the tokens
- Even if someone copies the file, they cannot use it on another machine or account

### Fallback (If DPAPI unavailable)
- If Windows DPAPI is not available, an unencrypted fallback is used
- A warning is logged in debug mode
- Consider deleting the cache after each session if on a shared computer

## Authentication Flow

```
┌─────────────┐    ┌──────────────┐    ┌─────────────────┐
│  TodoExport │───▶│   Browser    │───▶│ Microsoft Login │
│    App      │    │              │    │   (microsoft.com)│
└─────────────┘    └──────────────┘    └─────────────────┘
       │                                        │
       │         Device Code Flow               │
       │◀───────────────────────────────────────│
       │           (Token returned)             │
       ▼
┌─────────────────┐
│ Windows DPAPI   │
│ Encrypted Cache │
└─────────────────┘
```

1. App requests authentication via Device Code Flow
2. User authenticates directly with Microsoft in their browser
3. Microsoft returns tokens directly to MSAL library
4. Tokens are encrypted with Windows DPAPI before storage
5. **Your password never touches this application**

## Permissions Requested

| Scope | Purpose | Risk Level |
|-------|---------|------------|
| `Tasks.Read` | Read your To-Do lists and tasks (export only, no modifications) | Low - only To-Do read access |
| `User.Read` | Display your email/name | Minimal |
| `offline_access` | Stay logged in between sessions | Low |

**Note:** This app does NOT request:
- ❌ Mail access
- ❌ Calendar access
- ❌ Contacts access
- ❌ OneDrive access
- ❌ Admin permissions

## Can Someone Access My Microsoft Account?

### If someone has the token cache file:
- ❌ **Cannot login as you** - tokens are encrypted with DPAPI
- ❌ **Cannot access other services** - tokens only work for To-Do
- ❌ **Cannot change your password** - no such permission granted

### If someone has physical access to your PC:
- ⚠️ If logged into your Windows account, they could potentially use the cached token
- ✅ **Mitigation:** Use `SignOut` in the app to clear credentials
- ✅ **Mitigation:** Lock your Windows session when away

### Token Expiration
- Access tokens expire after ~1 hour
- Refresh tokens are long-lived but can be revoked
- You can revoke all tokens at: https://account.microsoft.com/security

## Security Best Practices

1. **Use Windows Hello or strong password** for your Windows account
2. **Lock your PC** when stepping away
3. **Sign out of TodoExport** if on a shared computer
4. **Review connected apps** periodically at https://account.microsoft.com/privacy/app-access

## How to Clear All Stored Credentials

### Option 1: Use the App
Click "Abmelden" (Sign Out) in the app - this clears the token cache.

### Option 2: Manual Deletion
Delete the folder:
```
%LocalAppData%\TodoExport
```

### Option 3: Revoke at Microsoft
1. Go to https://account.microsoft.com/privacy/app-access
2. Find "Microsoft Graph Command Line Tools"
3. Click "Remove"

## Reporting Security Issues

If you discover a security vulnerability, please:
1. **Do not** open a public GitHub issue
2. Contact the maintainer directly
3. Allow reasonable time for a fix before disclosure

## Third-Party Dependencies

| Library | Purpose | Security Notes |
|---------|---------|----------------|
| MSAL.NET | Microsoft Authentication | Official Microsoft library |
| ClosedXML | Excel generation | No network access |
| Newtonsoft.Json | JSON parsing | No network access |

All dependencies are from official NuGet sources and are widely used in production applications.

---

*Last updated: December 2025*

