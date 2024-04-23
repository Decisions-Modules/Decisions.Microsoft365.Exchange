using Decisions.OAuth;
using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.ServiceLayer.Services.OAuth.OAuth2;
using DecisionsFramework.Utilities.Data;

namespace Decisions.Microsoft365.Exchange
{
    public class GraphRest
    {
        public static HttpResponseMessage HttpResponsePost(ExchangeSettings? settingsOverride, string urlExtension, HttpContent content)
        {
            return SendHttpRequest(settingsOverride, urlExtension, content, HttpMethod.Post);
        }
        
        public static string Post(ExchangeSettings? settingsOverride, string urlExtension, HttpContent content)
        {
            HttpResponseMessage response = HttpResponsePost(settingsOverride, urlExtension, content);
            Task<string> resultTask = response.Content.ReadAsStringAsync();
            resultTask.Wait();

            return resultTask.Result;
        }
        
        public static string Get(ExchangeSettings? settingsOverride, string urlExtension)
        {
            HttpResponseMessage response = SendHttpRequest(settingsOverride, urlExtension, null, HttpMethod.Get);

            try
            {
                Task<string> resultTask = response.Content.ReadAsStringAsync();
                resultTask.Wait();

                return resultTask.Result;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not read response content.", ex);
            }
        }
        
        public static HttpResponseMessage HttpResponsePatch(ExchangeSettings? settingsOverride, string urlExtension, HttpContent content)
        {
            return SendHttpRequest(settingsOverride, urlExtension, content, HttpMethod.Patch);
        }

        public static string Patch(ExchangeSettings? settingsOverride, string urlExtension, HttpContent content)
        {
            HttpResponseMessage response = HttpResponsePatch(settingsOverride, urlExtension, content);
            Task<string> resultTask = response.Content.ReadAsStringAsync();
            resultTask.Wait();

            return resultTask.Result;
        }
        
        public static HttpResponseMessage Delete(ExchangeSettings? settingsOverride, string urlExtension)
        {
            return SendHttpRequest(settingsOverride, urlExtension, null, HttpMethod.Delete);
        }
        
        private static HttpResponseMessage SendHttpRequest(ExchangeSettings? settingsOverride, string urlExtension,
            HttpContent? content, HttpMethod httpMethod)
        {
            ExchangeSettings settings = ModuleSettingsAccessor<ExchangeSettings>.GetSettings();
            OAuthToken token = new ORM<OAuthToken>().Fetch(!string.IsNullOrEmpty(settingsOverride?.TokenId)
                ? settingsOverride.TokenId : settings.TokenId);
            
            string url = (!string.IsNullOrEmpty(settingsOverride?.GraphUrl) ? settingsOverride.GraphUrl : settings.GraphUrl) + urlExtension;
            string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");

            HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);

            HttpRequestMessage request = new HttpRequestMessage(httpMethod, url);
            request.Headers.Add("Authorization", tokenHeader);

            if (content != null)
            {
                request.Content = content;
            }
            
            try
            {
                HttpResponseMessage response = client.Send(request);
                response.EnsureSuccessStatusCode();

                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
    }
}