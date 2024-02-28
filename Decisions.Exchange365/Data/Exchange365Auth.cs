using Azure.Identity;
using DecisionsFramework.ServiceLayer;
using Microsoft.Graph;

namespace Decisions.Exchange365.Data;

public class Exchange365Auth
{
    private static DeviceCodeCredentialOptions options = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
        ClientId = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().ClientId,
        TenantId = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().TenantId,
        // Callback function that receives the user prompt
        // Prompt contains the generated device code that user must
        // enter during the auth process in the browser
        DeviceCodeCallback = (code, cancellation) =>
        {
            Console.WriteLine(code.Message);
            return Task.FromResult(0);
        },
    };

    static DeviceCodeCredential deviceCodeCredential = new(options);

    internal static GraphServiceClient GraphClient = new(deviceCodeCredential, ModuleSettingsAccessor<Exchange365Settings>.GetSettings().Scopes);
}