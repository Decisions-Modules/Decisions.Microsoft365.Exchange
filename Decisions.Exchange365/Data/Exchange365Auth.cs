using Azure.Identity;
using DecisionsFramework.ServiceLayer;
using Microsoft.Graph;

namespace Decisions.Exchange365.Data;

public class Exchange365Auth
{
    private static string clientId = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().ClientId;
    private static string tenantId = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().TenantId;
    private static string[] scopes = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().Scopes;

    private static string clientSecret = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().ClientSecret;
    private static string authorizationCode = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().AuthorizationCode;

    // using Azure.Identity;
    static AuthorizationCodeCredentialOptions options = new AuthorizationCodeCredentialOptions
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };

    // https://learn.microsoft.com/dotnet/api/azure.identity.authorizationcodecredential
    static AuthorizationCodeCredential authCodeCredential = new AuthorizationCodeCredential(
        tenantId, clientId, clientSecret, authorizationCode, options);

    internal static GraphServiceClient GraphClient = new GraphServiceClient(authCodeCredential, scopes);
}