using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.People
{
    [Writable]
    public class Microsoft365Person
    {
        [WritableValue]
        [JsonProperty("birthday")]
        public string? Birthday { get; set; }

        [WritableValue]
        [JsonProperty("companyName")]
        public string? CompanyName { get; set; }

        [WritableValue]
        [JsonProperty("department")]
        public string? Department { get; set; }

        [WritableValue]
        [JsonProperty("displayName")]
        public string? DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("scoredEmailAddresses")]
        public Microsoft365ScoredEmailAddress[]? ScoredEmailAddresses { get; set; }

        [WritableValue]
        [JsonProperty("givenName")]
        public string? GivenName { get; set; }

        [WritableValue]
        [JsonProperty("id")]
        public string? Id { get; set; }

        [WritableValue]
        [JsonProperty("imAddress")]
        public string? ImAddress { get; set; }

        [WritableValue]
        [JsonProperty("isFavorite")]
        public bool? IsFavorite { get; set; }

        [WritableValue]
        [JsonProperty("jobTitle")]
        public string? JobTitle { get; set; }

        [WritableValue]
        [JsonProperty("officeLocation")]
        public string? OfficeLocation { get; set; }

        [WritableValue]
        [JsonProperty("personNotes")]
        public string? PersonNotes { get; set; }

        [WritableValue]
        [JsonProperty("personType")]
        public Microsoft365PersonType? PersonType { get; set; }

        [WritableValue]
        [JsonProperty("phones")]
        public Microsoft365Phone[]? Phones { get; set; }

        [WritableValue]
        [JsonProperty("postalAddresses")]
        public Microsoft365ExactLocation[]? PostalAddresses { get; set; }

        [WritableValue]
        [JsonProperty("profession")]
        public string? Profession { get; set; }

        [WritableValue]
        [JsonProperty("surname")]
        public string? Surname { get; set; }

        [WritableValue]
        [JsonProperty("userPrincipalName")]
        public string? UserPrincipalName { get; set; }

        [WritableValue]
        [JsonProperty("websites")]
        public Microsoft365Website[]? Websites { get; set; }

        [WritableValue]
        [JsonProperty("yomiCompany")]
        public string? YomiCompany { get; set; }
    }
    
    [Writable]
    public class Microsoft365ScoredEmailAddress
    {
        [WritableValue]
        [JsonProperty("address")]
        public string? Address { get; set; }

        [WritableValue]
        [JsonProperty("relevanceScore")]
        public double? RelevanceScore { get; set; }
    }
    
    [Writable]
    public class Microsoft365PersonType
    {
        [WritableValue]
        [JsonProperty("class")]
        public string? Class { get; set; }

        [WritableValue]
        [JsonProperty("subclass")]
        public string? Subclass { get; set; }
    }
    
    [Writable]
    public class Microsoft365Phone
    {
        [WritableValue]
        [JsonProperty("number")]
        public string? Number { get; set; }

        [WritableValue]
        [JsonProperty("type")]
        public string? Type { get; set; }
    }
    
    [Writable]
    public class Microsoft365ExactLocation
    {
        [WritableValue]
        [JsonProperty("address")]
        public Microsoft365PhysicalAddress Address { get; set; }

        [WritableValue]
        [JsonProperty("coordinates")]
        public Microsoft365GeoCoordinates Coordinates { get; set; }

        [WritableValue]
        [JsonProperty("displayName")]
        public string? DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("locationEmailAddress")]
        public string? LocationEmailAddress { get; set; }

        [WritableValue]
        [JsonProperty("locationUri")]
        public string? LocationUri { get; set; }

        [WritableValue]
        [JsonProperty("locationType")]
        public string? LocationType { get; set; }

        [WritableValue]
        [JsonProperty("uniqueId")]
        public string? UniqueId { get; set; }

        [WritableValue]
        [JsonProperty("uniqueIdType")]
        public string? UniqueIdType { get; set; }
    }
    
    [Writable]
    public class Microsoft365GeoCoordinates
    {
        [WritableValue]
        [JsonProperty("accuracy")]
        public double? Accuracy { get; set; }

        [WritableValue]
        [JsonProperty("altitude")]
        public double? Altitude { get; set; }

        [WritableValue]
        [JsonProperty("altitudeAccuracy")]
        public double? AltitudeAccuracy { get; set; }

        [WritableValue]
        [JsonProperty("latitude")]
        public double? Latitude { get; set; }

        [WritableValue]
        [JsonProperty("longitude")]
        public double? Longitude { get; set; }
    }
    
    [Writable]
    public class Microsoft365Website
    {
        [WritableValue]
        [JsonProperty("address")]
        public string? Address { get; set; }

        [WritableValue]
        [JsonProperty("displayName")]
        public string? DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("type")]
        public string? Type { get; set; }
    }
}
