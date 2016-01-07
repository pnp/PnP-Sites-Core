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
                if (this.ParentTemplate != null && this.ParentTemplate.WebSettings != null)
                {
                    return (this.ParentTemplate.WebSettings.SiteLogo);
                }
                else
                {
                    return (null);
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
                    this.ParentTemplate.WebSettings.SiteLogo = value;
                }
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
                if (this.ParentTemplate != null && this.ParentTemplate.WebSettings != null)
                {
                    return (this.ParentTemplate.WebSettings.AlternateCSS);
                }
                else
                {
                    return (null);
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
                    this.ParentTemplate.WebSettings.AlternateCSS = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the MasterPage for the Composed Look
        /// </summary>
        [Obsolete("Instead of this member, please use MasterPageUrl property of the WebSettings object.")]
        public string MasterPage
        {
            get
            {
                if (this.ParentTemplate != null && this.ParentTemplate.WebSettings != null)
                {
                    return (this.ParentTemplate.WebSettings.MasterPageUrl);
                }
                else
                {
                    return (null);
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
                    this.ParentTemplate.WebSettings.MasterPageUrl = value;
                }
            }
        }
    }
}
