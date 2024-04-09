using System;
using Decisions.Microsoft365.Exchange.API.Email;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.People
{
    [Writable]
    public class Microsoft365Contact
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public string? OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("@odata.etag")]
        public string? OdataEtag { get; set; }

        [WritableValue]
        [JsonProperty("id")]
        public string? Id { get; set; }

        [WritableValue]
        [JsonProperty("createdDateTime")]
        public DateTimeOffset? CreatedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("lastModifiedDateTime")]
        public DateTimeOffset? LastModifiedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("changeKey")]
        public string? ChangeKey { get; set; }

        [WritableValue]
        [JsonProperty("categories")]
        public string[]? Categories { get; set; }

        [WritableValue]
        [JsonProperty("parentFolderId")]
        public string? ParentFolderId { get; set; }

        [WritableValue]
        [JsonProperty("birthday")]
        public DateTimeOffset? Birthday { get; set; }

        [WritableValue]
        [JsonProperty("fileAs")]
        public string? FileAs { get; set; }

        [WritableValue]
        [JsonProperty("displayName")]
        public string? DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("givenName")]
        public string? GivenName { get; set; }

        [WritableValue]
        [JsonProperty("initials")]
        public string? Initials { get; set; }

        [WritableValue]
        [JsonProperty("middleName")]
        public string? MiddleName { get; set; }

        [WritableValue]
        [JsonProperty("nickName")]
        public string? NickName { get; set; }

        [WritableValue]
        [JsonProperty("surname")]
        public string? Surname { get; set; }

        [WritableValue]
        [JsonProperty("title")]
        public string? Title { get; set; }

        [WritableValue]
        [JsonProperty("yomiGivenName")]
        public string? YomiGivenName { get; set; }

        [WritableValue]
        [JsonProperty("yomiSurname")]
        public string? YomiSurname { get; set; }

        [WritableValue]
        [JsonProperty("yomiCompanyName")]
        public string? YomiCompanyName { get; set; }

        [WritableValue]
        [JsonProperty("generation")]
        public string? Generation { get; set; }

        [WritableValue]
        [JsonProperty("imAddresses")]
        public string[]? ImAddresses { get; set; }

        [WritableValue]
        [JsonProperty("jobTitle")]
        public string? JobTitle { get; set; }

        [WritableValue]
        [JsonProperty("companyName")]
        public string? CompanyName { get; set; }

        [WritableValue]
        [JsonProperty("department")]
        public string? Department { get; set; }

        [WritableValue]
        [JsonProperty("officeLocation")]
        public string? OfficeLocation { get; set; }

        [WritableValue]
        [JsonProperty("profession")]
        public string? Profession { get; set; }

        [WritableValue]
        [JsonProperty("businessHomePage")]
        public string? BusinessHomePage { get; set; }

        [WritableValue]
        [JsonProperty("assistantName")]
        public string? AssistantName { get; set; }

        [WritableValue]
        [JsonProperty("manager")]
        public string? Manager { get; set; }

        [WritableValue]
        [JsonProperty("homePhones")]
        public string[]? HomePhones { get; set; }

        [WritableValue]
        [JsonProperty("mobilePhone")]
        public string? MobilePhone { get; set; }

        [WritableValue]
        [JsonProperty("businessPhones")]
        public string[]? BusinessPhones { get; set; }

        [WritableValue]
        [JsonProperty("spouseName")]
        public string? SpouseName { get; set; }

        [WritableValue]
        [JsonProperty("personalNotes")]
        public string? PersonalNotes { get; set; }

        [WritableValue]
        [JsonProperty("children")]
        public string[]? Children { get; set; }

        [WritableValue]
        [JsonProperty("emailAddresses")]
        public Microsoft365EmailAddress[]? EmailAddresses { get; set; }

        [WritableValue]
        [JsonProperty("homeAddress")]
        public Microsoft365PhysicalAddress? HomeAddress { get; set; }

        [WritableValue]
        [JsonProperty("businessAddress")]
        public Microsoft365PhysicalAddress? BusinessAddress { get; set; }

        [WritableValue]
        [JsonProperty("otherAddress")]
        public Microsoft365PhysicalAddress? OtherAddress { get; set; }
        
        public static Microsoft365Contact? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365Contact>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }

    [Writable]
    public class Microsoft365PhysicalAddress
    {
        [WritableValue]
        [JsonProperty("city")]
        public string? City { get; set; }

        [WritableValue]
        [JsonProperty("countryOrRegion")]
        public string? CountryOrRegion { get; set; }

        [WritableValue]
        [JsonProperty("postalCode")]
        public string? PostalCode { get; set; }

        [WritableValue]
        [JsonProperty("state")]
        public string? State { get; set; }

        [WritableValue]
        [JsonProperty("street")]
        public string? Street { get; set; }
    }
}
