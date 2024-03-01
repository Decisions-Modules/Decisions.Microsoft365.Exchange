using Azure.Core;
using Azure.Identity;
using DecisionsFramework.ServiceLayer;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Decisions.Exchange365.Data
{
    public class GraphHelper
    {
        // App-ony auth token credential
        private static ClientSecretCredential? _clientSecretCredential;
        // Client configured with app-only authentication
        private static GraphServiceClient? _appClient;
        
        private static string tenantId = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().TenantId;
        
        private static string clientId = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().ClientId;
        
        private static string clientSecret = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().ClientSecretValue;

        public static void InitializeGraphForAppOnlyAuth()
        {
            if (_clientSecretCredential == null)
            {
                _clientSecretCredential = new ClientSecretCredential(
                    tenantId, clientId, clientSecret);
            }

            if (_appClient == null)
            {
                _appClient = new GraphServiceClient(_clientSecretCredential,
                    // Use the default scope, which will request the scopes
                    // configured on the app registration
                    new[] {"https://graph.microsoft.com/.default"});
            }

            StoreAccessTokenAsync();
        }
        
        public static async Task<string> GetAppOnlyTokenAsync()
        {
            // Ensure credential isn't null
            _ = _clientSecretCredential ??
                throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

            // Request token with given scopes
            var context = new TokenRequestContext(new[] {"https://graph.microsoft.com/.default"});
            var response = await _clientSecretCredential.GetTokenAsync(context);
            return response.Token;
        }

        static async Task StoreAccessTokenAsync()
        {
            try
            {
                string appOnlyToken = await GraphHelper.GetAppOnlyTokenAsync();

                ModuleSettingsAccessor<Exchange365Settings>.GetSettings().Token = appOnlyToken;
                ModuleSettingsAccessor<Exchange365Settings>.SaveSettings();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting app-only access token: {ex.Message}");
            }
        }
    }
}