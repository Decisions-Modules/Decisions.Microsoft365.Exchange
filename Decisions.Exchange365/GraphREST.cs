using System.Net.Http.Json;
using Decisions.Exchange365.Data;
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.Utilities.Data;

namespace Decisions.Exchange365;

public class GraphREST
{
    public static async Task<string> Post(string url, JsonContent? content)
    {
        await GraphHelper.InitializeGraphForAppOnlyAuth();
        
        string token = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().Token;
        HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
        
        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
        request.Headers.Add("Authorization", $"Bearer {token}");
        
        request.Content = content;
        
        HttpResponseMessage response = client.Send(request);
        //response.EnsureSuccessStatusCode();

        Task<string> resultTask = response.Content.ReadAsStringAsync();
        resultTask.Wait();

        return resultTask.Result;
    }

    public static async Task<string> Get(string url)
    {
        await GraphHelper.InitializeGraphForAppOnlyAuth();
        
        string token = ModuleSettingsAccessor<Exchange365Settings>.GetSettings().Token;
        HttpClient client = HttpClients.GetHttpClient(HttpClientAuthType.Normal);
        
        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Add("Authorization", $"Bearer {token}");

        HttpResponseMessage response = client.Send(request);
        //response.EnsureSuccessStatusCode();

        Task<string> resultTask = response.Content.ReadAsStringAsync();
        resultTask.Wait();

        return resultTask.Result;
    }
}