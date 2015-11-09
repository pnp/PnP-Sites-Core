using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that specifies the properties of the new list.
    /// </summary>
    public partial class ListInstance : IEquatable<ListInstance>
    {
        #region Constructors

        public ListInstance() { }

        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows) :
                this(contentTypeBindings, views, fields, fieldRefs, dataRows, null, null, null, null)
        {
        }

        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows, Dictionary<String, String> fieldDefaults, ObjectSecurity security, List<Localization> listLocalizations, List<Localization> fieldsLocalizations)
        {
            if (contentTypeBindings != null)
            {
                this.ContentTypeBindings.AddRange(contentTypeBindings);
            }

            if (views != null)
            {
                this.Views.AddRange(views);
            }

            if (fields != null)
            {
                this.Fields.AddRange(fields);
            }

            if (fieldRefs != null)
            {
                this._fieldRefs.AddRange(fieldRefs);
            }
            if (dataRows != null)
            {
                this._dataRows.AddRange(dataRows);
            }
            if (fieldDefaults != null)
            {
                this._fieldDefaults = fieldDefaults;
            }
            if (security != null)
            {
                this._security = security;
            }
            if (listLocalizations != null)
            {
                this._listLocalizations.AddRange(listLocalizations);
            }
            if (fieldsLocalizations != null)
            {
                this._fieldsLocalizations.AddRange(fieldsLocalizations);
            }
        }

        #endregion

        #region Private Members
        private List<ContentTypeBinding> _ctBindings = new List<ContentTypeBinding>();
        private List<View> _views = new List<View>();
        private List<Field> _fields = new List<Field>();
        private List<FieldRef> _fieldRefs = new List<FieldRef>();
        private List<DataRow> _dataRows = new List<DataRow>();
        private List<Localization> _listLocalizations = new List<Localization>();
        private List<Localization> _fieldsLocalizations = new List<Localization>();
        private Dictionary<String, String> _fieldDefaults = new Dictionary<String, String>();
        private ObjectSecurity _security = null;
        private bool _enableFolderCreation = true;
        private bool _enableAttachments = true;
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the list title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the description of the list
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets a value that specifies the identifier of the document template for the new list.
        /// </summary>
        public string DocumentTemplate { get; set; }

        /// <summary>
        /// Gets or sets a value that specifies whether the new list is displayed on the Quick Launch of the site.
        /// </summary>
        public bool OnQuickLaunch { get; set; }

        /// <summary>
        /// Gets or sets a value that specifies the list server template of the new list.
        /// https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.listtemplatetype.aspx
        /// </summary>
        public int TemplateType { get; set; }

        /// <summary>
        /// Gets or sets a value that specifies whether the new list is displayed on the Quick Launch of the site.
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Gets or sets whether verisioning is enabled on the list
        /// </summary>
        public bool EnableVersioning { get; set; }

        /// <summary>
        /// Gets or sets whether minor verisioning is enabled on the list
        /// </summary>
        public bool EnableMinorVersions { get; set; }

        /// <summary>
        /// Gets or sets the DraftVersionVisibility for the list
        /// </summary>
        public int DraftVersionVisibility { get; set; }

        /// <summary>
        /// Gets or sets whether moderation/content approval is enabled on the list
        /// </summary>
        public bool EnableModeration { get; set; }

        /// <summary>
        /// Gets or sets the MinorVersionLimit  for versioning, just in case it is enabled on the list
        /// </summary>
        public int MinorVersionLimit { get; set; }

        /// <summary>
        /// Gets or sets the MinorVersionLimit  for verisioning, just in case it is enabled on the list
        /// </summary>
        public int MaxVersionLimit { get; set; }

        /// <summary>
        /// Gets or sets whether existing content types should be removed
        /// </summary>
        public bool RemoveExistingContentTypes { get; set; }

        /// <summary>
        /// Gets or sets whether existing views should be removed
        /// </summary>
        public bool RemoveExistingViews { get; set; }

        /// <summary>
        /// Gets or sets whether content types are enabled
        /// </summary>
        public bool ContentTypesEnabled { get; set; }

        /// <summary>
        /// Gets or sets whether to hide the list
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// Gets or sets whether attachments are enabled. Defaults to true.
        /// </summary>
        public bool EnableAttachments
        {
            get { return _enableAttachments; }
            set { _enableAttachments = value; }
        }

        /// <summary>
        /// Gets or sets whether folder is enabled. Defaults to true.
        /// </summary>
        public bool EnableFolderCreation
        {
            get { return _enableFolderCreation; }
            set { _enableFolderCreation = value; }
        }
        /// <summary>
        /// Gets or sets the content types to associate to the list
        /// </summary>
        public List<ContentTypeBinding> ContentTypeBindings
        {
            get { return this._ctBindings; }
            private set { this._ctBindings = value; }
        }

        /// <summary>
        /// Gets or sets the content types to associate to the list
        /// </summary>
        public List<View> Views
        {
            get { return this._views; }
            private set { this._views = value; }
        }

        public List<Field> Fields
        {
            get { return this._fields; }
            private set { this._fields = value; }
        }

        public List<FieldRef> FieldRefs
        {
            get { return this._fieldRefs; }
            private set { this._fieldRefs = value; }
        }

        public Guid TemplateFeatureID { get; set; }

        public List<DataRow> DataRows
        {
            get { return this._dataRows; }
            private set { this._dataRows = value; }
        }

        public List<Localization> ListLocalizations
        {
            get { return this._listLocalizations; }
            private set { this._listLocalizations = value; }
        }

        public List<Localization> FieldsLocalizations
        {
            get { return this._fieldsLocalizations; }
            private set { this._fieldsLocalizations = value; }
        }

        /// <summary>
        /// Defines a list of default values for the Fields of the List Instance
        /// </summary>
        public Dictionary<String, String> FieldDefaults
        {
            get { return this._fieldDefaults; }
            private set { this._fieldDefaults = value; }
        }

        /// <summary>
        /// Defines the Security rules for the List Instance
        /// </summary>
        public ObjectSecurity Security
        {
            get { return this._security; }
            set { this._security = value; }
        }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}|{21}|{22}|{23}|",
                this.ContentTypesEnabled.GetHashCode(),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                (this.DocumentTemplate != null ? this.DocumentTemplate.GetHashCode() : 0),
                this.EnableVersioning.GetHashCode(),
                this.Hidden.GetHashCode(),
                this.MaxVersionLimit.GetHashCode(),
                this.MinorVersionLimit.GetHashCode(),
                this.OnQuickLaunch.GetHashCode(),
                this.EnableAttachments.GetHashCode(),
                this.EnableFolderCreation.GetHashCode(),
                this.RemoveExistingContentTypes.GetHashCode(),
                this.TemplateType.GetHashCode(),
                (this.Title != null ? this.Title.GetHashCode() : 0),
                (this.Url != null ? this.Url.GetHashCode() : 0),
                (this.TemplateFeatureID != null ? this.TemplateFeatureID.GetHashCode() : 0),
                this.RemoveExistingViews.GetHashCode(),
                this.EnableMinorVersions.GetHashCode(),
                this.EnableModeration.GetHashCode(),
                this.ContentTypeBindings.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Views.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Fields.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.FieldRefs.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.ListLocalizations.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.FieldsLocalizations.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ListInstance))
            {
                return (false);
            }
            return (Equals((ListInstance)obj));
        }

        public bool Equals(ListInstance other)
        {
            return (this.ContentTypesEnabled == other.ContentTypesEnabled &&
                this.Description == other.Description &&
                this.DocumentTemplate == other.DocumentTemplate &&
                this.EnableVersioning == other.EnableVersioning &&
                this.EnableMinorVersions == other.EnableMinorVersions &&
                this.EnableModeration == other.EnableModeration &&
                this.Hidden == other.Hidden &&
                this.MaxVersionLimit == other.MaxVersionLimit &&
                this.MinorVersionLimit == other.MinorVersionLimit &&
                this.OnQuickLaunch == other.OnQuickLaunch &&
                this.EnableAttachments == other.EnableAttachments &&
                this.EnableFolderCreation == other.EnableFolderCreation &&
                this.RemoveExistingContentTypes == other.RemoveExistingContentTypes &&
                this.TemplateType == other.TemplateType &&
                this.Title == other.Title &&
                this.Url == other.Url &&
                this.TemplateFeatureID == other.TemplateFeatureID &&
                this.RemoveExistingViews == other.RemoveExistingViews &&
                this.ContentTypeBindings.DeepEquals(other.ContentTypeBindings) &&
                this.Views.DeepEquals(other.Views) &&
                this.Fields.DeepEquals(other.Fields) &&
                this.FieldRefs.DeepEquals(other.FieldRefs) &&
                this.ListLocalizations.DeepEquals(other.ListLocalizations) &&
                this.FieldsLocalizations.DeepEquals(other.FieldsLocalizations));
        }

        #endregion
    }
}
