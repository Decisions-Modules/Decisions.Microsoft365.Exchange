using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class GroupList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public GroupValue[] Value { get; set; }
    }

    [Writable]
    public class GroupValue
    {
        [WritableValue]
        [JsonProperty("id")]
        public Guid Id { get; set; }

        [WritableValue]
        [JsonProperty("deletedDateTime")]
        public object DeletedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("classification")]
        public object Classification { get; set; }

        [WritableValue]
        [JsonProperty("createdDateTime")]
        public DateTimeOffset CreatedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("creationOptions")]
        public string[] CreationOptions { get; set; }

        [WritableValue]
        [JsonProperty("description")]
        public string Description { get; set; }

        [WritableValue]
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("expirationDateTime")]
        public object ExpirationDateTime { get; set; }

        [WritableValue]
        [JsonProperty("groupTypes")]
        public string[] GroupTypes { get; set; }

        [WritableValue]
        [JsonProperty("isAssignableToRole")]
        public bool? IsAssignableToRole { get; set; }

        [WritableValue]
        [JsonProperty("mail")]
        public string Mail { get; set; }

        [WritableValue]
        [JsonProperty("mailEnabled")]
        public bool MailEnabled { get; set; }

        [WritableValue]
        [JsonProperty("mailNickname")]
        public string MailNickname { get; set; }

        [WritableValue]
        [JsonProperty("membershipRule")]
        public object MembershipRule { get; set; }

        [WritableValue]
        [JsonProperty("membershipRuleProcessingState")]
        public object MembershipRuleProcessingState { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesDomainName")]
        public object OnPremisesDomainName { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesLastSyncDateTime")]
        public object OnPremisesLastSyncDateTime { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesNetBiosName")]
        public object OnPremisesNetBiosName { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesSamAccountName")]
        public object OnPremisesSamAccountName { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesSecurityIdentifier")]
        public object OnPremisesSecurityIdentifier { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesSyncEnabled")]
        public object OnPremisesSyncEnabled { get; set; }

        [WritableValue]
        [JsonProperty("preferredDataLocation")]
        public object PreferredDataLocation { get; set; }

        [WritableValue]
        [JsonProperty("preferredLanguage")]
        public object PreferredLanguage { get; set; }

        [WritableValue]
        [JsonProperty("proxyAddresses")]
        public string[] ProxyAddresses { get; set; }

        [WritableValue]
        [JsonProperty("renewedDateTime")]
        public DateTimeOffset RenewedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("resourceBehaviorOptions")]
        public string[] ResourceBehaviorOptions { get; set; }

        [WritableValue]
        [JsonProperty("resourceProvisioningOptions")]
        public string[] ResourceProvisioningOptions { get; set; }

        [WritableValue]
        [JsonProperty("securityEnabled")]
        public bool SecurityEnabled { get; set; }

        [WritableValue]
        [JsonProperty("securityIdentifier")]
        public string SecurityIdentifier { get; set; }

        [WritableValue]
        [JsonProperty("theme")]
        public object Theme { get; set; }

        [WritableValue]
        [JsonProperty("visibility")]
        public string? Visibility { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesProvisioningErrors")]
        public object[] OnPremisesProvisioningErrors { get; set; }

        [WritableValue]
        [JsonProperty("serviceProvisioningErrors")]
        public object[] ServiceProvisioningErrors { get; set; }
    }
}