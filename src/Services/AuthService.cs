using System.Diagnostics;
using System.IO;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace TodoExport.Services;

public class AuthService
{
    // Öffentliche Microsoft Client-ID für Graph-Zugriff
    // Hinweis: Bei der Anmeldung wird "Microsoft Graph Command Line Tools" angezeigt,
    // da wir eine öffentliche Microsoft Client-ID verwenden. Dies ist normal und sicher.
    private const string ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
    
    // Minimale Berechtigungen - nur was benötigt wird
    private static readonly string[] Scopes = new[]
    {
        "Tasks.ReadWrite",  // To-Do Lesen/Schreiben
        "User.Read",        // Benutzername anzeigen
        "offline_access"    // Refresh Token für Auto-Login
    };

    // Token Cache Konfiguration mit Windows DPAPI-Verschlüsselung
    private const string CacheFileName = "todo_export_cache.bin";
    private static readonly string CacheDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TodoExport");

    private IPublicClientApplication? _msalClient;
    private AuthenticationResult? _authResult;
    private bool _isInitialized;

    // Events für UI-Updates
    public event Action<string, string>? DeviceCodeReceived;
    public event Action? AuthenticationCompleted;
    public event Action<string>? AuthenticationFailed;

    public bool IsAuthenticated => _authResult != null && _authResult.ExpiresOn > DateTimeOffset.UtcNow;
    public string? UserName => _authResult?.Account?.Username;
    public string? AccessToken => _authResult?.AccessToken;

    private async Task InitializeAsync()
    {
        if (_isInitialized) return;

        _msalClient = PublicClientApplicationBuilder
            .Create(ClientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, "consumers")
            .WithDefaultRedirectUri()
            .Build();

        // Sicheres Token Caching mit Windows DPAPI-Verschlüsselung
        try
        {
            Directory.CreateDirectory(CacheDir);
            
            // Windows DPAPI verschlüsselt die Datei automatisch mit dem Windows-Benutzerkonto
            // Nur der angemeldete Windows-Benutzer kann die Datei entschlüsseln
            var storageProperties = new StorageCreationPropertiesBuilder(CacheFileName, CacheDir)
                .WithLinuxKeyring(
                    schemaName: "com.todoexport.tokencache",
                    collection: "default",
                    secretLabel: "TodoExport Token Cache",
                    attribute1: new KeyValuePair<string, string>("Version", "1"),
                    attribute2: new KeyValuePair<string, string>("App", "TodoExport"))
                .WithMacKeyChain(
                    serviceName: "TodoExport",
                    accountName: "TokenCache")
                .Build();

            var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
            
            // Verifiziere, dass die Verschlüsselung funktioniert
            cacheHelper.VerifyPersistence();
            
            cacheHelper.RegisterCache(_msalClient.UserTokenCache);
        }
        catch (Exception ex)
        {
            // App funktioniert auch ohne persistentes Caching
            // Token müssen dann bei jedem Start neu angefordert werden
            Debug.WriteLine($"Token cache setup failed: {ex.Message}");
        }

        _isInitialized = true;
    }

    public async Task<bool> SignInAsync(CancellationToken cancellationToken = default)
    {
        await InitializeAsync();
        
        if (_msalClient == null)
            throw new InvalidOperationException("MSAL Client nicht initialisiert");

        try
        {
            // Erst versuchen: Silent Token (aus Cache)
            var accounts = await _msalClient.GetAccountsAsync();
            var account = accounts.FirstOrDefault();

            if (account != null)
            {
                try
                {
                    _authResult = await _msalClient.AcquireTokenSilent(Scopes, account)
                        .ExecuteAsync(cancellationToken);
                    AuthenticationCompleted?.Invoke();
                    return true;
                }
                catch (MsalUiRequiredException)
                {
                    // Interaktive Anmeldung nötig (Token abgelaufen oder widerrufen)
                }
            }

            // Device Code Flow - Browser öffnet sich automatisch
            _authResult = await _msalClient.AcquireTokenWithDeviceCode(Scopes, deviceCodeResult =>
            {
                // Browser automatisch öffnen
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = deviceCodeResult.VerificationUrl.ToString(),
                        UseShellExecute = true
                    });
                }
                catch { }

                DeviceCodeReceived?.Invoke(deviceCodeResult.UserCode, deviceCodeResult.VerificationUrl.ToString());
                return Task.CompletedTask;
            }).ExecuteAsync(cancellationToken);

            AuthenticationCompleted?.Invoke();
            return _authResult != null;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
        catch (MsalException ex)
        {
            AuthenticationFailed?.Invoke(ex.Message);
            return false;
        }
    }

    public async Task<string?> GetAccessTokenAsync()
    {
        await InitializeAsync();
        
        if (_msalClient == null)
            return null;

        // Wenn kein Token vorhanden, versuche Silent Login
        if (_authResult == null)
        {
            try
            {
                var accounts = await _msalClient.GetAccountsAsync();
                var account = accounts.FirstOrDefault();
                
                if (account != null)
                {
                    _authResult = await _msalClient.AcquireTokenSilent(Scopes, account)
                        .ExecuteAsync();
                }
                else
                {
                    return null; // Kein Account vorhanden - Anmeldung erforderlich
                }
            }
            catch (MsalUiRequiredException)
            {
                // Token abgelaufen oder widerrufen - Anmeldung erforderlich
                return null;
            }
        }

        // Token erneuern wenn es in 5 Minuten abläuft
        if (_authResult.ExpiresOn <= DateTimeOffset.UtcNow.AddMinutes(5))
        {
            try
            {
                var accounts = await _msalClient.GetAccountsAsync();
                var account = accounts.FirstOrDefault();
                
                if (account != null)
                {
                    _authResult = await _msalClient.AcquireTokenSilent(Scopes, account)
                        .ExecuteAsync();
                }
            }
            catch (MsalUiRequiredException)
            {
                // Token widerrufen - Anmeldung erforderlich
                _authResult = null;
                return null;
            }
        }

        return _authResult?.AccessToken;
    }

    public async Task SignOutAsync()
    {
        await InitializeAsync();
        
        if (_msalClient == null) return;

        // Alle Accounts aus Cache entfernen
        var accounts = await _msalClient.GetAccountsAsync();
        foreach (var account in accounts)
        {
            await _msalClient.RemoveAsync(account);
        }
        _authResult = null;
        
        // Optional: Cache-Datei löschen für vollständige Bereinigung
        try
        {
            var cacheFile = Path.Combine(CacheDir, CacheFileName);
            if (File.Exists(cacheFile))
            {
                File.Delete(cacheFile);
            }
        }
        catch { }
    }

    public async Task<bool> TrySilentSignInAsync()
    {
        try
        {
            await InitializeAsync();
            
            if (_msalClient == null) return false;

            var accounts = await _msalClient.GetAccountsAsync();
            var account = accounts.FirstOrDefault();

            if (account != null)
            {
                _authResult = await _msalClient.AcquireTokenSilent(Scopes, account)
                    .ExecuteAsync();
                return true;
            }
        }
        catch
        {
            // Silent login fehlgeschlagen - keine Aktion nötig
        }
        return false;
    }
    
    /// <summary>
    /// Löscht alle gespeicherten Anmeldedaten vollständig
    /// </summary>
    public static void ClearAllCredentials()
    {
        try
        {
            if (Directory.Exists(CacheDir))
            {
                Directory.Delete(CacheDir, recursive: true);
            }
        }
        catch { }
    }
}
