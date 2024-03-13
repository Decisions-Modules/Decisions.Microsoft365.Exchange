using System.Net.Http.Json;
using Decisions.OAuth;
using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.ServiceLayer.Services.OAuth.OAuth2;
using DecisionsFramework.Utilities.Data;

namespace Decisions.Exchange365;

public class GraphRest
{
    public static HttpResponseMessage HttpResponsePost(string url, JsonContent content)
    {
        try
        {
            OAuthToken token = new ORM<OAuthToken>().Fetch(ModuleSettingsAccessor<Exchange365Settings>.GetSettings().TokenId);
            string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");
            
            HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
            
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Add("Authorization", tokenHeader);
            
            request.Content = content;
            
            HttpResponseMessage response = client.Send(request);
            response.EnsureSuccessStatusCode();
            
            return response;
        }
        catch (Exception ex)
        {
            throw new BusinessRuleException("The request was unsuccessful.", ex);
        }
    }

    public static string Post(string url, JsonContent? content)
    {
        HttpResponseMessage response = HttpResponsePost(url, content);
        Task<string> resultTask = response.Content.ReadAsStringAsync();
        resultTask.Wait();
        
        return resultTask.Result;
    }

    public static string Get(string url)
    {
        try
        {
            OAuthToken token = new ORM<OAuthToken>().Fetch(ModuleSettingsAccessor<Exchange365Settings>.GetSettings().TokenId);
            string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");
        
            HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
        
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Add("Authorization", tokenHeader);

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
    
    public static HttpResponseMessage Patch(string url, JsonContent? content)
    {
        try
        {
            OAuthToken token = new ORM<OAuthToken>().Fetch(ModuleSettingsAccessor<Exchange365Settings>.GetSettings().TokenId);
            string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");
        
            HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
        
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Patch, url);
            request.Headers.Add("Authorization", tokenHeader);
        
            request.Content = content;
        
            HttpResponseMessage response = client.Send(request);
            response.EnsureSuccessStatusCode();

            return response;
        }
        catch (Exception ex)
        {
            throw new BusinessRuleException("The request was unsuccessful.", ex);
        }
    }
    
    public static HttpResponseMessage Delete(string url)
    {
        try
        {
            OAuthToken token = new ORM<OAuthToken>().Fetch(ModuleSettingsAccessor<Exchange365Settings>.GetSettings().TokenId);
            string tokenHeader = OAuth2Utility.GetOAuth2HeaderValue(token.TokenData, "Bearer");
        
            HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
        
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, url);
            request.Headers.Add("Authorization", tokenHeader);

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