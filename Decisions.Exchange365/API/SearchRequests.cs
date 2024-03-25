using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class SearchRequests
    {
        [WritableValue]
        [JsonProperty("requests")]
        public Request[] Requests { get; set; }
    }

    [Writable]
    public class Request
    {
        [WritableValue]
        [JsonProperty("entityTypes")]
        public string[] EntityTypes { get; set; }

        [WritableValue]
        [JsonProperty("query")]
        public Query Query { get; set; }

        [WritableValue]
        [JsonProperty("from")]
        public long From { get; set; }

        [WritableValue]
        [JsonProperty("size")]
        public long Size { get; set; }
    }

    [Writable]
    public class Query
    {
        [WritableValue]
        [JsonProperty("queryString")]
        public string QueryString { get; set; }
    }
}