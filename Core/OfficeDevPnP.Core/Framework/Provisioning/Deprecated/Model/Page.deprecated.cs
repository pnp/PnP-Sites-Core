using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class Page
    {
        [Obsolete("Instead of this member, please use WelcomePage property of the WebSettings object.")]
        public bool WelcomePage
        {
            get
            {
                if (this.ParentTemplate != null && this.ParentTemplate.WebSettings != null)
                {
                    return (this.Url == this.ParentTemplate.WebSettings.WelcomePage);
                }
                else
                {
                    return (false);
                }
            }
            set
            {
                // Initialize the WebSettings property if it is not already there
                if (this.ParentTemplate != null && this.ParentTemplate.WebSettings == null)
                {
                    this.ParentTemplate.WebSettings = new WebSettings();
                }

                if (this.ParentTemplate != null && this.ParentTemplate.WebSettings != null)
                {
                    this.ParentTemplate.WebSettings.WelcomePage = this.Url;
                }
            }
        }

        [Obsolete("Instead of this constructor, please use the one without the WelcomePage property")]
        public Page(string url, bool overwrite, WikiPageLayout layout, IEnumerable<WebPart> webParts, bool welcomePage = false, ObjectSecurity security = null) :
            this(url, overwrite, layout, webParts, welcomePage, security, null)
        {
        }

        [Obsolete("Instead of this constructor, please use the one without the WelcomePage property")]
        public Page(string url, bool overwrite, WikiPageLayout layout, IEnumerable<WebPart> webParts, bool welcomePage = false, ObjectSecurity security = null, Dictionary<String, String> fields = null) :
            this(url, overwrite, layout, webParts, security, fields)
        {
        }
    }
}