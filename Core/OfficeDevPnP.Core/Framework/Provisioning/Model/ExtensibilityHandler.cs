using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object for Extensiblity Call out
    /// </summary>
    public class ExtensibilityHandler : BaseModel, IEquatable<ExtensibilityHandler>
    {
        #region Properties

        public bool Enabled
        {
            get;
            set;
        }

        public string Assembly
        {
            get;
            set;
        }

        public string Type
        {
            get;
            set;
        }

        public string Configuration { get; set; }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                (this.Assembly != null ? this.Assembly.GetHashCode() : 0),
                (this.Configuration != null ? this.Configuration.GetHashCode() : 0),
                this.Enabled.GetHashCode(),
                (this.Type != null ? this.Type.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ExtensibilityHandler))
            {
                return (false);
            }
            return (Equals((ExtensibilityHandler)obj));
        }

        public bool Equals(ExtensibilityHandler other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Assembly == other.Assembly &&
                this.Configuration == other.Configuration &&
                this.Enabled == other.Enabled &&
                this.Type == other.Type);
        }

        #endregion
    }
}
