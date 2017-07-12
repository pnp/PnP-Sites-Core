using System;
using System.Net;
using System.Net.Http;
using System.Web.Http.Controllers;
using System.Web.Http.Filters;

namespace OfficeDevPnP.Core.WebAPI
{
    /// <summary>
    /// Class deals with WebAPI context filter attribute and excutes action based on HttpActionContext
    /// </summary>
    public class WebAPIContextFilterAttribute : ActionFilterAttribute
    {
        /// <summary>
        /// Method executes on HTTP action
        /// </summary>
        /// <param name="actionContext">HttpActionContext object</param>
        public override void OnActionExecuting(HttpActionContext actionContext)
        {
            if (actionContext == null)
            {
                throw new ArgumentNullException("actionContext");
            }

            if (WebAPIHelper.HasCacheEntry(actionContext.ControllerContext))
            {
                return;
            }
            else
            {
                actionContext.Response = actionContext.Request.CreateErrorResponse(HttpStatusCode.MethodNotAllowed, CoreResources.Services_AccessDenied);
                return;
            }
        }
    }
}
