using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a ClientSidePage
    /// </summary>
    public partial class ClientSidePage : BaseModel, IEquatable<ClientSidePage>
    {
        #region Private Members

        private CanvasZoneCollection _zones;

        #endregion

        #region Public Members

        /// <summary>
        /// Gets or sets the zones
        /// </summary>
        public CanvasZoneCollection Zones
        {
            get { return _zones; }
            private set { _zones = value; }
        }

        /// <summary>
        /// Defines the Pages Library of the Client Side Page, required attribute.
        /// </summary>
        public String PagesLibrary { get; set; }

        /// <summary>
        /// Defines whether to promote the page as a news article, optional attribute
        /// </summary>
        public Boolean PromoteAsNewsArticle { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for ClientSidePage class
        /// </summary>
        public ClientSidePage()
        {
            this._zones = new CanvasZoneCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.Zones.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                PagesLibrary?.GetHashCode() ?? 0,
                PromoteAsNewsArticle.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ClientSidePage class
        /// </summary>
        /// <param name="obj">Object that represents ClientSidePage</param>
        /// <returns>Checks whether object is ClientSidePage class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ClientSidePage))
            {
                return (false);
            }
            return (Equals((ClientSidePage)obj));
        }

        /// <summary>
        /// Compares ClientSidePage object based on Zones, PagesLibrary, and PromoteAsNewsArticle
        /// </summary>
        /// <param name="other">ClientSidePage Class object</param>
        /// <returns>true if the ClientSidePage object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ClientSidePage other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Zones.DeepEquals(other.Zones)  &&
                this.PagesLibrary == other.PagesLibrary &&
                this.PromoteAsNewsArticle == other.PromoteAsNewsArticle
                );
        }

        #endregion
    }
}
