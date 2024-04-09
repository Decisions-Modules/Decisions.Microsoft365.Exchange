using System;
using Decisions.Microsoft365.Exchange.API.People;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Group
{
    [Writable]
    public class Microsoft365Group
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public string? OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("id")]
        public string? Id { get; set; }

        [WritableValue]
        [JsonProperty("deletedDateTime")]
        public DateTimeOffset? DeletedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("classification")]
        public string? Classification { get; set; }
        [WritableValue]
        
        [JsonProperty("createdDateTime")]
        public DateTimeOffset? CreatedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("description")]
        public string? Description { get; set; }

        [WritableValue]
        [JsonProperty("displayName")]
        public string? DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("expirationDateTime")]
        public DateTimeOffset? ExpirationDateTime { get; set; }

        [WritableValue]
        [JsonProperty("groupTypes")]
        public string[]? GroupTypes { get; set; }

        [WritableValue]
        [JsonProperty("isAssignableToRole")]
        public bool? IsAssignableToRole { get; set; }

        [WritableValue]
        [JsonProperty("mail")]
        public string? Mail { get; set; }

        [WritableValue]
        [JsonProperty("mailEnabled")]
        public bool? MailEnabled { get; set; }

        [WritableValue]
        [JsonProperty("mailNickname")]
        public string? MailNickname { get; set; }

        [WritableValue]
        [JsonProperty("membershipRule")]
        public string? MembershipRule { get; set; }

        [WritableValue]
        [JsonProperty("membershipRuleProcessingState")]
        public string? MembershipRuleProcessingState { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesDomainName")]
        public string? OnPremisesDomainName { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesLastSyncDateTime")]
        public DateTimeOffset? OnPremisesLastSyncDateTime { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesNetBiosName")]
        public string? OnPremisesNetBiosName { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesSamAccountName")]
        public string? OnPremisesSamAccountName { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesSecurityIdentifier")]
        public string? OnPremisesSecurityIdentifier { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesSyncEnabled")]
        public bool? OnPremisesSyncEnabled { get; set; }

        [WritableValue]
        [JsonProperty("preferredDataLocation")]
        public string? PreferredDataLocation { get; set; }

        [WritableValue]
        [JsonProperty("preferredLanguage")]
        public string? PreferredLanguage { get; set; }

        [WritableValue]
        [JsonProperty("proxyAddresses")]
        public string[]? ProxyAddresses { get; set; }

        [WritableValue]
        [JsonProperty("renewedDateTime")]
        public DateTimeOffset? RenewedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("resourceBehaviorOptions")]
        public string[]? ResourceBehaviorOptions { get; set; }

        [WritableValue]
        [JsonProperty("resourceProvisioningOptions")]
        public string[]? ResourceProvisioningOptions { get; set; }

        [WritableValue]
        [JsonProperty("securityEnabled")]
        public bool? SecurityEnabled { get; set; }

        [WritableValue]
        [JsonProperty("securityIdentifier")]
        public string? SecurityIdentifier { get; set; }

        [WritableValue]
        [JsonProperty("serviceProvisioningErrors")]
        public Microsoft365ServiceProvisioningError[]? ServiceProvisioningErrors { get; set; }

        [WritableValue]
        [JsonProperty("theme")]
        public string? Theme { get; set; }

        [WritableValue]
        [JsonProperty("visibility")]
        public string? Visibility { get; set; }

        [WritableValue]
        [JsonProperty("onPremisesProvisioningErrors")]
        public Microsoft365OnPremisesProvisioningError[]? OnPremisesProvisioningErrors { get; set; }
        
        public static Microsoft365Group? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365Group>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
    
    [Writable]
    public class Microsoft365ServiceProvisioningError
    {
        [WritableValue]
        [JsonProperty("createdDateTime")]
        public DateTimeOffset? CreatedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("isResolved")]
        public bool? IsResolved { get; set; }

        [WritableValue]
        [JsonProperty("serviceInstance")]
        public string? ServiceInstance { get; set; }
    }
    
    [Writable]
    public class Microsoft365OnPremisesProvisioningError
    {
        [WritableValue]
        [JsonProperty("category")]
        public string? Category { get; set; }

        [WritableValue]
        [JsonProperty("occurredDateTime")]
        public string? OccurredDateTime { get; set; }

        [WritableValue]
        [JsonProperty("propertyCausingError")]
        public string? PropertyCausingError { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public string? Value { get; set; }
    }
}
