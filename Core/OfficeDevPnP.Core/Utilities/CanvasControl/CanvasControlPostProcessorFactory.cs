using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities.CanvasControl.Processors;

namespace OfficeDevPnP.Core.Utilities.CanvasControl
{
#if !SP2013 && !SP2016
    public class CanvasControlPostProcessorFactory
    {
        /// <summary>
        /// Resolves client control web part by type
        /// </summary>
        /// <param name="canvasControl">CanvasControl object</param>
        /// <returns>Returns PassThroughProcessor object</returns>
        public static ICanvasControlPostProcessor Resolve(Framework.Provisioning.Model.CanvasControl canvasControl)
        {
            if (canvasControl.Type == WebPartType.List)
            {
                return new ListControlPostProcessor(canvasControl);
            }

            return new CanvasControlPassThroughProcessor();
        }
    }
#endif
}