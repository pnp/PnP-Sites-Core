using OfficeDevPnP.Core.Pages;

namespace OfficeDevPnP.Core.Utilities.CanvasControl
{
#if !SP2013 && !SP2016
    /// <summary>
    ///     Interface for WebPart Post Processing
    /// </summary>
    public interface ICanvasControlPostProcessor
    {
        /// <summary>
        ///     Method for processing canvas control
        /// </summary>
        /// <param name="canvasControl">Canvas control object</param>
        /// <param name="clientSidePage">ClientSidePage object</param>
        void Process(Framework.Provisioning.Model.CanvasControl canvasControl, ClientSidePage clientSidePage);
    }
#endif
}