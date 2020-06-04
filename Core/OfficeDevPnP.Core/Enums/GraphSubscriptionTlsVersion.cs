namespace OfficeDevPnP.Core.Enums
{
    /// <summary>
    /// Defines the TLS versions a Microsoft Graph subscription supports calling into when an event for which a subscription exists gets triggered
    /// </summary>
    public enum GraphSubscriptionTlsVersion : short
    {
        v1_0,
        v1_1,
        v1_2,
        v1_3
    }
}
