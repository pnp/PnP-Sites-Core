using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Represents the Header of a Client Side page
    /// </summary>
    public partial class ClientSidePageHeader : BaseModel, IEquatable<ClientSidePageHeader>
    {
        #region Public Members

        /// <summary>
        /// Defines the type of the header for the client side page
        /// </summary>
        public ClientSidePageHeaderType Type { get; set; }

        /// <summary>
        /// Defines the server-relative URL of the image for the header of the current client side page
        /// </summary>
        public String ServerRelativeImageUrl { get; set; }

        /// <summary>
        /// Defines the x-translate of the image for the header of the current client side page.
        /// </summary>
        public Double? TranslateX { get; set; }

        /// <summary>
        /// Defines the y-translate of the image for the header of the current client side page.
        /// </summary>
        public Double? TranslateY { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                this.Type.GetHashCode(),
                this.ServerRelativeImageUrl.GetHashCode(),
                this.TranslateX.GetHashCode(),
                this.TranslateY.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ClientSidePageHeader class
        /// </summary>
        /// <param name="obj">Object that represents ClientSidePageHeader</param>
        /// <returns>Checks whether object is ClientSidePageHeader class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ClientSidePageHeader))
            {
                return (false);
            }
            return (Equals((ClientSidePageHeader)obj));
        }

        /// <summary>
        /// Compares ClientSidePageHeader object based on Type, ServerRelativeImageUrl, TranslateX, and TranslateY
        /// </summary>
        /// <param name="other">ClientSidePageHeader Class object</param>
        /// <returns>true if the ClientSidePageHeader object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ClientSidePageHeader other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Type == other.Type &&
                this.ServerRelativeImageUrl == other.ServerRelativeImageUrl &&
                this.TranslateX == other.TranslateX &&
                this.TranslateY == other.TranslateY
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the possible values for the client side page header type
    /// </summary>
    public enum ClientSidePageHeaderType
    {
        /// <summary>
        /// The client side page does not hava any header
        /// </summary>
        None,
        /// <summary>
        /// Default client side page header
        /// </summary>
        Default,
        /// <summary>
        /// The client side page has a custom header
        /// </summary>
        Custom,
    }
}
