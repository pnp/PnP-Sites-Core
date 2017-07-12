using System;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Represents a Field XML Markup that is used to define information about a field
    /// </summary>
    public partial class Field : BaseModel, IEquatable<Field>
    {
        #region Private Members
        private string _schemaXml = string.Empty;
        #endregion

        #region Public Properties

        /// <summary>
        /// Gets a value that specifies the XML Schema representing the Field type.
        /// <seealso>
        ///     <cref>https://msdn.microsoft.com/en-us/library/office/ff407271.aspx</cref>
        /// </seealso>
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
            XElement element = PrepareFieldForCompare(this.SchemaXml);
            return element.ToString().GetHashCode();
        }

        /// <summary>
        /// Compares object with Field
        /// </summary>
        /// <param name="obj">Object</param>
        /// <returns>true if the current object is equal to the Field</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Field))
            {
                return (false);
            }
            return (Equals((Field)obj));
        }

        /// <summary>
        /// Compares Field object based on currentXMl and otherXml
        /// </summary>
        /// <param name="other">Object that represents Field</param>
        /// <returns>true if the Field object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Field other)
        {
            if (other == null)
            {
                return (false);
            }

            XElement currentXml = PrepareFieldForCompare(this.SchemaXml);
            XElement otherXml = PrepareFieldForCompare(other.SchemaXml);
            return (XNode.DeepEquals(currentXml, otherXml));
        }

        private XElement PrepareFieldForCompare(string schemaXML)
        {
            XElement element = XElement.Parse(schemaXML);
            if (element.Attribute("SourceID") != null)
            {
                element.Attribute("SourceID").Remove();
            }

            return element;
        }
        #endregion
    }
}
