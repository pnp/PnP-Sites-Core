using OfficeDevPnP.Core.Pages;

namespace OfficeDevPnP.Core.Utilities.CanvasControl.Processors
{
#if !SP2013 && !SP2016
    /// <summary>
    ///     Default processor when others are not resolved
    /// </summary>
    public class CanvasControlPassThroughProcessor : ICanvasControlPostProcessor
    {
        /// <summary>
        ///  Method for processing canvas control
        /// </summary>
        /// <param name="canvasControl">Canvas control object</param>
        /// <param name="clientSidePage">ClientSidePage object</param>
        public void Process(Framework.Provisioning.Model.CanvasControl canvasControl, ClientSidePage clientSidePage)
        {
        }
    }
#endif    
}