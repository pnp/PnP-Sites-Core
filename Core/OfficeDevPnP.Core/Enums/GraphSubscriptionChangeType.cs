using System;

namespace OfficeDevPnP.Core.Enums
{
    /// <summary>
    /// Supported Microsoft Graph Change Types on Subscriptions. Documentation at: https://docs.microsoft.com/graph/api/resources/subscription#properties
    /// </summary>
    [Flags]
    public enum GraphSubscriptionChangeType : short
    {
        /// <summary>
        /// Something got created
        /// </summary>
        Created = 1,

        /// <summary>
        /// Something existing got updated
        /// </summary>
        Updated = 2,

        /// <summary>
        /// Something got deleted
        /// </summary>
        Deleted = 4
    }
}
