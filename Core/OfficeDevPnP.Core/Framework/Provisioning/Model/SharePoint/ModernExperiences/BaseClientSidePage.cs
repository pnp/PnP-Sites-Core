using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a base Client Side Page
    /// </summary>
    public abstract partial class BaseClientSidePage : BaseModel, IEquatable<BaseClientSidePage>
    {
        #region Private Members

        private CanvasSectionCollection _sections;
        private ObjectSecurity _security = null;
        private ClientSidePageHeader _header;

        #endregion

        #region Public Members

        /// <summary>
        /// Gets or sets the sections
        /// </summary>
        public CanvasSectionCollection Sections
        {
            get { return _sections; }
            private set { _sections = value; }
        }

        /// <summary>
        /// Defines the Security rules for the client-side Page
        /// </summary>
        public ObjectSecurity Security
        {
            get { return this._security; }
            internal set
            {
                if (this._security != null)
                {
                    this._security.ParentTemplate = null;
                }
                this._security = value;
                if (this._security != null)
                {
                    this._security.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Defines whether to promote the page as a news article, optional attribute
        /// </summary>
        public Boolean PromoteAsNewsArticle { get; set; }

        /// <summary>
        /// Defines whether to promote the page as a template, optional attribute
        /// </summary>
        public Boolean PromoteAsTemplate { get; set; }

        /// <summary>
        /// Defines whether the page can be overwritten if it exists
        /// </summary>
        public Boolean Overwrite { get; set; }

        /// <summary>
        /// Defines the Layout for the client-side page
        /// </summary>
        public String Layout { get; set; }

        /// <summary>
        /// Defines whether to publish the client-side page or not
        /// </summary>
        public Boolean Publish { get; set; }

        /// <summary>
        /// Defines whether the page will have comments enabled or not
        /// </summary>
        public Boolean EnableComments { get; set; }

        /// <summary>
        /// Defines the Title for the client-side page
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Defines the ContentTypeID for the client-side page
        /// </summary>
        public String ContentTypeID { get; set; }

        /// <summary>
        /// Defines the Header for the client-side page
        /// </summary>
        public ClientSidePageHeader Header
        {
            get
            {
                return (this._header);
            }
            set
            {
                if (this._header != null)
                {
                    this._header.ParentTemplate = null;
                }
                this._header = value;
                if (this._header != null)
                {
                    this._header.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Defines the page fields values, if any
        /// </summary>
        public Dictionary<String, String> FieldValues { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Defines property bag properties for the client side page
        /// </summary>
        public Dictionary<String, String> Properties { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Defines the URL of the thumbnail for the client side page
        /// </summary>
        public String ThumbnailUrl { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for BaseClientSidePage class
        /// </summary>
        public BaseClientSidePage()
        {
            this._sections = new CanvasSectionCollection(this.ParentTemplate);
            Security = new ObjectSecurity();
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|",
                this.Sections.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.PromoteAsNewsArticle.GetHashCode(),
                this.Overwrite.GetHashCode(),
                this.Layout?.GetHashCode() ?? 0,
                this.Publish.GetHashCode(),
                this.EnableComments.GetHashCode(),
                this.Title?.GetHashCode() ?? 0,
                this.FieldValues.Aggregate(0, (acc, next) => acc += (next.Value != null ? next.Value.GetHashCode() : 0)),
                this.ContentTypeID.GetHashCode(),
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.PromoteAsTemplate.GetHashCode(),
                this.ThumbnailUrl.GetHashCode(),
                this.GetInheritedHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Returns the HashCode of the members of any inherited type
        /// </summary>
        /// <returns></returns>
        protected abstract int GetInheritedHashCode();

        /// <summary>
        /// Compares object with BaseClientSidePage class
        /// </summary>
        /// <param name="obj">Object that represents BaseClientSidePage</param>
        /// <returns>Checks whether object is BaseClientSidePage class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is BaseClientSidePage))
            {
                return (false);
            }
            return (Equals((BaseClientSidePage)obj));
        }

        /// <summary>
        /// Compares BaseClientSidePage object based on Sections, PromoteAsNewsArticle, Overwrite, Layout, 
        /// Publish, EnableComments, Title, Properties, PromoteAsTemplate, and ThumbnailUrl
        /// </summary>
        /// <param name="other">BaseClientSidePage Class object</param>
        /// <returns>true if the BaseClientSidePage object is equal to the current object; otherwise, false.</returns>
        public bool Equals(BaseClientSidePage other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Sections.DeepEquals(other.Sections) &&
                this.PromoteAsNewsArticle == other.PromoteAsNewsArticle &&
                this.Overwrite == other.Overwrite &&
                this.Layout == other.Layout &&
                this.Publish == other.Publish &&
                this.EnableComments == other.EnableComments &&
                this.Title == other.Title &&
                this.FieldValues.DeepEquals(other.FieldValues) &&
                this.ContentTypeID == other.ContentTypeID &&
                this.Properties.DeepEquals(other.Properties) &&
                this.PromoteAsTemplate == other.PromoteAsTemplate &&
                this.ThumbnailUrl == other.ThumbnailUrl &&
                this.EqualsInherited(other)
                );
        }

        /// <summary>
        /// Compares the members of any inherited type
        /// </summary>
        /// <returns></returns>
        protected abstract bool EqualsInherited(BaseClientSidePage other);

        #endregion
    }
}
