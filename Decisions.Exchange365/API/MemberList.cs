using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class MemberList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public MemberValue[] Value { get; set; }
    }

    [Writable]
    public class MemberValue
    {
        [WritableValue]
        [JsonProperty("@odata.type")]
        public string OdataType { get; set; }

        [WritableValue]
        [JsonProperty("id")]
        public Guid Id { get; set; }

        [WritableValue]
        [JsonProperty("businessPhones")]
        public object[] BusinessPhones { get; set; }

        [WritableValue]
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("givenName")]
        public string GivenName { get; set; }

        [WritableValue]
        [JsonProperty("jobTitle")]
        public object JobTitle { get; set; }

        [WritableValue]
        [JsonProperty("mail")]
        public string Mail { get; set; }

        [WritableValue]
        [JsonProperty("mobilePhone")]
        public object MobilePhone { get; set; }

        [WritableValue]
        [JsonProperty("officeLocation")]
        public object OfficeLocation { get; set; }

        [WritableValue]
        [JsonProperty("preferredLanguage")]
        public string PreferredLanguage { get; set; }

        [WritableValue]
        [JsonProperty("surname")]
        public string Surname { get; set; }

        [WritableValue]
        [JsonProperty("userPrincipalName")]
        public string UserPrincipalName { get; set; }
    }
}