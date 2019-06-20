using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;

namespace OfficeDevPnP.Core.Utilities.WebParts
{
    /// <summary>
    /// Interface for WebPart Post Processing
    /// </summary>
    public interface IWebPartPostProcessor
    {
        /// <summary>
        /// Method for processing webpart
        /// </summary>
        /// <param name="wpDefinition">WebPartDefinition object</param>
        /// <param name="webPartPage">File object</param>
        void Process(WebPartDefinition wpDefinition, File webPartPage);
    }
}
