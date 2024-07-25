namespace Decisions.Microsoft365.Exchange;

public static class Microsoft365Utility
{
    internal static ExchangeSettings? GetExchangeSettings(InputExchangeSettings? settingsOverride)
    {
        var exchangeSettingsOverride = settingsOverride != null
            ? new ExchangeSettings() { TokenId = settingsOverride?.TokenId, GraphUrl = settingsOverride?.GraphUrl }
            : null;
        
        return exchangeSettingsOverride;
    }
}