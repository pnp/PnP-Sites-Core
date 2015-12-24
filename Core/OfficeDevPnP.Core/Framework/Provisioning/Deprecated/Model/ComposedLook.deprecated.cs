using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class ComposedLook
    {
        /// <summary>
        /// Gets or sets the Site Logo
        /// </summary>
        [Obsolete("Instead of this member, please use SiteLogo property of the WebSettings object.")]
        public string SiteLogo
        {
            get
            {
                return (this.ParentTemplate.WebSettings.SiteLogo);
            }
            set
            {
                this.ParentTemplate.WebSettings.SiteLogo = value;
            }
        }

        /// <summary>
        /// Gets or sets the AlternateCSS
        /// </summary>
        [Obsolete("Instead of this member, please use AlternateCSS property of the WebSettings object.")]
        public string AlternateCSS
        {
            get
            {
                return null;
                // return (this.ParentTemplate.WebSettings.AlternateCSS);
            }
            set
            {
                // this.ParentTemplate.WebSettings.AlternateCSS = value;
            }
        }
    }
}
