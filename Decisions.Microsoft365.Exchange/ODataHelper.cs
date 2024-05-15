using Decisions.Microsoft365.Common;
using DecisionsFramework;

namespace Decisions.Microsoft365.Exchange
{
    public static class ODataHelper<T>
    {
        internal static T? GetNextPage(ExchangeSettings? settingsOverride, string oDataUrl)
        {
            try
            {
                HttpResponseMessage response = GraphRest.SendHttpRequest(settingsOverride, oDataUrl, null, HttpMethod.Get);
                
                try
                {
                    Task<string> resultTask = response.Content.ReadAsStringAsync();
                    resultTask.Wait();
                    
                    return JsonHelper<T?>.JsonDeserialize(resultTask.Result);
                }
                catch (Exception ex)
                {
                    throw new BusinessRuleException("Could not read response content.", ex);
                }
            }
            catch (Exception ex)
            {
                throw new LoggedException($"Could not deserialize result: {oDataUrl}", ex);
            }
        }
    }
}