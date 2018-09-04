using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a DocumentSet Template for creating multiple DocumentSet instances
    /// </summary>
    public partial class DocumentSetTemplate : BaseModel, IEquatable<DocumentSetTemplate>
    {
        #region Private Members

        private List<String> _allowedContentTypes = new List<String>();
        private DefaultDocumentCollection _defaultDocuments;
        private List<Guid> _sharedFields = new List<Guid>();
        private List<Guid> _welcomePageFields = new List<Guid>();

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for DocumentSetTemplate class
        /// </summary>
        public DocumentSetTemplate()
        {
            _defaultDocuments = new DefaultDocumentCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor for DocumentSetTemplate class
        /// </summary>
        /// <param name="welcomePage">Welcome Page of the DocumentSet</param>
        /// <param name="allowedContentTypes">Content Types allowed for the DocumentSet</param>
        /// <param name="defaultDocuments">Default documents for the DocumentSet</param>
        /// <param name="sharedFields">Shared Fields for the DocumentSet</param>
        /// <param name="welcomePageFields">Welcome Page Fields for the DocumentSet</param>
        public DocumentSetTemplate(String welcomePage, IEnumerable<String> allowedContentTypes = null, IEnumerable<DefaultDocument> defaultDocuments = null, IEnumerable<Guid> sharedFields = null, IEnumerable<Guid> welcomePageFields = null) : 
            this()
        {
            if (!String.IsNullOrEmpty(welcomePage))
            {
                this.WelcomePage = welcomePage;
            }
            if (allowedContentTypes != null)
            {
                this._allowedContentTypes.AddRange(allowedContentTypes);
            }
            this.DefaultDocuments.AddRange(defaultDocuments);
            if (sharedFields != null)
            {
                this._sharedFields.AddRange(sharedFields);
            }
            if (welcomePageFields != null)
            {
                this._welcomePageFields.AddRange(welcomePageFields);
            }
        }

        #endregion

        #region Public Members

        /// <summary>
        /// The list of allowed Content Types for the Document Set
        /// </summary>
        public List<String> AllowedContentTypes
        {
            get { return this._allowedContentTypes; }
            private set { this._allowedContentTypes = value; }
        }

        /// <summary>
        /// The list of default Documents for the Document Set
        /// </summary>
        public DefaultDocumentCollection DefaultDocuments
        {
            get { return this._defaultDocuments; }
            private set { this._defaultDocuments = value; }
        }

        /// <summary>
        /// The list of Shared Fields for the Document Set
        /// </summary>
        public List<Guid> SharedFields
        {
            get { return this._sharedFields; }
            private set { this._sharedFields = value; }
        }

        /// <summary>
        /// The list of Welcome Page Fields for the Document Set
        /// </summary>
        public List<Guid> WelcomePageFields
        {
            get { return this._welcomePageFields; }
            private set { this._welcomePageFields = value; }
        }

        /// <summary>
        /// Defines the custom WelcomePage for the Document Set
        /// </summary>
        public String WelcomePage { get; set; }

        /// <summary>
        /// The RemoveExistingContentTypes flag for the Allowed Content Types of the current Document Set, optional attribute.
        /// </summary>
        public Boolean RemoveExistingContentTypes { get; set; }

        /// <summary>
        /// Gets a value that specifies the XML Documents settings for the Document Set.
        /// <seealso>
        ///     <cref>https://msdn.microsoft.com/en-us/library/office/aa543904.aspx</cref>
        /// </seealso>
        /// </summary>
        public XElement XmlDocuments { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|",
                this.AllowedContentTypes.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.DefaultDocuments.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.SharedFields.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.WelcomePageFields.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.RemoveExistingContentTypes.GetHashCode(),
                this.XmlDocuments?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with DocumentSetTemplate
        /// </summary>
        /// <param name="obj">Object that represents DocumentSetTemplate</param>
        /// <returns>true if the current object is equal to the DocumentSetTemplate</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is DocumentSetTemplate))
            {
                return (false);
            }
            return (Equals((DocumentSetTemplate)obj));
        }

        /// <summary>
        /// Compares DocumentSetTemplate object based on AllowedContentTypes, DefaultDocuments, SharedFields, WelcomePageFields, RemoveExistingContentTypes, and XmlDocuments properties.
        /// </summary>
        /// <param name="other">DocumentSetTemplate object</param>
        /// <returns>true if the DocumentSetTemplate object is equal to the current object; otherwise, false.</returns>
        public bool Equals(DocumentSetTemplate other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AllowedContentTypes.DeepEquals(other.AllowedContentTypes) &&
                    this.DefaultDocuments.DeepEquals(other.DefaultDocuments) &&
                    this.SharedFields.DeepEquals(other.SharedFields) &&
                    this.WelcomePageFields.DeepEquals(other.WelcomePageFields) &&
                    this.RemoveExistingContentTypes == other.RemoveExistingContentTypes &&
                    this.XmlDocuments == other.XmlDocuments
                );
        }

        #endregion
    }
}
