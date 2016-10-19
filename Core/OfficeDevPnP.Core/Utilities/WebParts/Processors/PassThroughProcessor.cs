using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;

namespace OfficeDevPnP.Core.Utilities.WebParts.Processors
{
    /// <summary>
    /// Default processor when others are not resolved
    /// </summary>
    public class PassThroughProcessor : IWebPartPostProcessor
    {
        public void Process(WebPartDefinition wpDefinition, File webPartPage)
        {
            
        }
    }
}
