using System;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class View : BaseModel, IEquatable<View>
    {
        #region Private Members
        private string _schemaXml = string.Empty;
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets a value that specifies the XML Schema representing the View type.
        /// </summary>
        public string SchemaXml
        {
            get { return this._schemaXml; }
            set { this._schemaXml = value; }
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            XElement element = PrepareViewForCompare(this.SchemaXml);
            return (element != null ? element.ToString().GetHashCode() : 0);
        }

        /// <summary>
        /// Compares object with View
        /// </summary>
        /// <param name="obj">Object that represents View</param>
        /// <returns>true if the current object is equal to the View</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is View))
            {
                return (false);
            }
            return (Equals((View)obj));
        }

        /// <summary>
        /// Compares View object based on currentXml and otherXml
        /// </summary>
        /// <param name="other">View object</param>
        /// <returns>true if the View object is equal to the current object; otherwise, false.</returns>
        public bool Equals(View other)
        {
            if (other == null)
            {
                return (false);
            }

            XElement currentXml = PrepareViewForCompare(this.SchemaXml);
            XElement otherXml = PrepareViewForCompare(other.SchemaXml);
            return (XNode.DeepEquals(currentXml, otherXml));
        }

        private XElement PrepareViewForCompare(string schemaXML)
        {
            XElement element = XElement.Parse(schemaXML);
            if (element.Attribute("Name") != null)
            {
                Guid nameGuid = Guid.Empty;
                if (Guid.TryParse(element.Attribute("Name").Value, out nameGuid))
                {
                    // Temporary remove guid
                    element.Attribute("Name").Remove();
                }
            }
            if (element.Attribute("Url") != null)
            {
                element.Attribute("Url").Remove();
            }
            if (element.Attribute("ImageUrl") != null)
            {
                var index = element.Attribute("ImageUrl").Value.IndexOf("rev=", StringComparison.InvariantCultureIgnoreCase);

                if (index > -1)
                {
                    // Remove ?rev=23 in URL
                    Regex regex = new Regex("\\?rev=([0-9])\\w+");
                    element.SetAttributeValue("ImageUrl", regex.Replace(element.Attribute("ImageUrl").Value, ""));
                }
            }

            return element;
        }

        #endregion
    }
}
