using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;

namespace OfficeDevPnP.Core.Utilities.WebParts.Processors
{
    /// <summary>
    /// Default processor when others are not resolved
    /// </summary>
    public class PassThroughProcessor : IWebPartPostProcessor
    {
        /// <summary>
        /// Method to process webpart when it is not resolved
        /// </summary>
        /// <param name="wpDefinition">WebPartDefinition object</param>
        /// <param name="webPartPage">File object</param>
        public void Process(WebPartDefinition wpDefinition, File webPartPage)
        {
            
        }
    }
}
