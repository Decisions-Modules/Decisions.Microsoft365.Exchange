using System.Runtime.Serialization;

namespace Decisions.Microsoft365.Exchange.API
{
    public enum AttendeeStatus
    {
        [EnumMember(Value = "free")]
        Free,
        
        [EnumMember(Value = "tentative")]
        Tentative,
        
        [EnumMember(Value = "busy")]
        Busy,
        
        [EnumMember(Value = "oof")]
        Oof,
        
        [EnumMember(Value = "workingElsewhere")]
        WorkingElsewhere,
        
        [EnumMember(Value = "unknown")]
        Unknown
    }
}