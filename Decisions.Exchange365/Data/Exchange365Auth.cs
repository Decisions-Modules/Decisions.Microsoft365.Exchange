using Azure.Identity;
using DecisionsFramework.ServiceLayer;
using Microsoft.Graph;

namespace Decisions.Exchange365.Data;

public class Exchange365Auth
{
    private static string clientId = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().ClientId;
    private static string tenantId = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().TenantId;
    private static string[] scopes = {"https://graph.microsoft.com/.default"};//ModuleSettingsAccessor<Exchange365Settings>.GetSettings().Scopes;

    private static string clientSecret = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().ClientSecretValue;

    // using Azure.Identity;
    private static ClientSecretCredentialOptions options = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };

    // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
    private static ClientSecretCredential clientSecretCredential = new(
            tenantId, clientId, clientSecret, options);

    internal static GraphServiceClient GraphClient = new(clientSecretCredential, scopes);
}