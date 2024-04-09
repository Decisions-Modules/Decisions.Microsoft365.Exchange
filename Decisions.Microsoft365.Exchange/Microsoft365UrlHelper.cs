namespace Decisions.Microsoft365.Exchange
{
    public class Microsoft365UrlHelper
    {
        internal static string GetUserUrl(string userIdentifier)
        {
            return $"/users/{userIdentifier}";
        }

        internal static string GetGroupUrl(string? groupId)
        {
            return (!string.IsNullOrEmpty(groupId)) ? $"/groups/{groupId}" : "/groups";
        }

        internal static string GetContactUrl(string userIdentifier, string? contactId, string? contactFolderId,
            string? childFolderId)
        {
            string urlExtension = GetUserUrl(userIdentifier);

            if (!string.IsNullOrEmpty(contactFolderId))
            {
                urlExtension = $"{urlExtension}/contactFolders/{contactFolderId}";

                if (!string.IsNullOrEmpty(childFolderId))
                {
                    urlExtension = $"{urlExtension}/childFolders/{childFolderId}";
                }
            }

            return (!string.IsNullOrEmpty(contactId))
                ? $"{urlExtension}/contacts/{contactId}"
                : $"{urlExtension}/contacts";
        }

        internal static string GetEmailUrl(string userIdentifier, string? messageId, string? mailFolderId)
        {
            return (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUserUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}"
                : $"{GetUserUrl(userIdentifier)}/messages/{messageId}";
        }

        internal static string GetCalendarEventUrl(string userIdentifier, string? eventId, string? calendarId,
            string? calendarGroupId)
        {
            string urlExtension = GetUserUrl(userIdentifier);

            if (!string.IsNullOrEmpty(calendarId))
            {
                if (!string.IsNullOrEmpty(calendarGroupId))
                {
                    urlExtension = $"{urlExtension}/calendarGroups/{calendarGroupId}/calendars/{calendarId}";
                }

                urlExtension = $"{urlExtension}/calendars/{calendarId}";
            }

            urlExtension = (!string.IsNullOrEmpty(eventId))
                ? $"{urlExtension}/events/{eventId}"
                : $"{urlExtension}/events";

            return urlExtension;
        }
    }
}