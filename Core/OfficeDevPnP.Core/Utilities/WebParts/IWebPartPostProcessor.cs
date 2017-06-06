using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;

namespace OfficeDevPnP.Core.Utilities.WebParts
{
    public interface IWebPartPostProcessor
    {
        void Process(WebPartDefinition wpDefinition, File webPartPage);
    }
}
