using Decisions.OAuth;
using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.ServiceLayer.Services.OAuth.OAuth2;
using DecisionsFramework.Utilities.Data;

namespace Decisions.Microsoft365.Exchange;

public class GraphRest
{
    private static ExchangeSettings settings = ModuleSettingsAccessor<ExchangeSettings>.GetSettings();
    
    private static OAuthToken token = new ORM<OAuthToken>().Fetch(settings.TokenId);
    
    public static HttpResponseMessage HttpResponsePost(string urlExtension, HttpContent content)
    {
        string url = $"{settings.GraphUrl}{urlExtension}";
        string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");
            
        HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
            
        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
        request.Headers.Add("Authorization", tokenHeader);
            
        request.Content = content;
        
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
    
    public static string Post(string urlExtension, HttpContent content)
    {
        HttpResponseMessage response = HttpResponsePost(urlExtension, content);
        Task<string> resultTask = response.Content.ReadAsStringAsync();
        resultTask.Wait();
        
        return resultTask.Result;
    }

    public static string Get(string urlExtension)
    {
        string url = $"{settings.GraphUrl}{urlExtension}";
        string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");
        
        HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
        
        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Add("Authorization", tokenHeader);
        
        try
        {
            HttpResponseMessage response = client.Send(request);
            response.EnsureSuccessStatusCode();

            Task<string> resultTask = response.Content.ReadAsStringAsync();
            resultTask.Wait();

            return resultTask.Result;
        }
        catch (Exception ex)
        {
            throw new BusinessRuleException("The request was unsuccessful.", ex);
        }
    }
    
    public static HttpResponseMessage HttpResponsePatch(string urlExtension, HttpContent content)
    {
        string url = $"{settings.GraphUrl}{urlExtension}";
        string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");
        
        HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
        
        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Patch, url);
        request.Headers.Add("Authorization", tokenHeader);
            
        request.Content = content;
        
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
    
    public static string Patch(string urlExtension, HttpContent content)
    {
        HttpResponseMessage response = HttpResponsePatch(urlExtension, content);
        Task<string> resultTask = response.Content.ReadAsStringAsync();
        resultTask.Wait();
        
        return resultTask.Result;
    }
    
    public static HttpResponseMessage Delete(string urlExtension)
    {
        string url = $"{settings.GraphUrl}{urlExtension}";
        string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");
        
        HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
        
        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, url);
        request.Headers.Add("Authorization", tokenHeader);
        
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