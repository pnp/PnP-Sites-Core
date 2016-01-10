using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that specifies the properties of the new list.
    /// </summary>
    public partial class ListInstance : BaseModel, IEquatable<ListInstance>
    {
        #region Constructors

        public ListInstance()
        {
            this._ctBindings = new ContentTypeBindingCollection(this.ParentTemplate);
            this._views = new ViewCollection(this.ParentTemplate);
            this._fields = new FieldCollection(this.ParentTemplate);
            this._fieldRefs = new FieldRefCollection(this.ParentTemplate);
            this._dataRows = new DataRowCollection(this.ParentTemplate);
            this._folders = new FolderCollection(this.ParentTemplate);
        }

        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows) :
                this(contentTypeBindings, views, fields, fieldRefs, dataRows, null, null, null)
        {
        }

        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows, Dictionary<String, String> fieldDefaults, ObjectSecurity security) :
                this(contentTypeBindings, views, fields, fieldRefs, dataRows, fieldDefaults, security, null)
        {
        }

        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows, Dictionary<String, String> fieldDefaults, ObjectSecurity security, List<Folder> folders) : 
            this()
        {
            this.ContentTypeBindings.AddRange(contentTypeBindings);
            this.Views.AddRange(views);
            this.Fields.AddRange(fields);
            this.FieldRefs.AddRange(fieldRefs);
            this.DataRows.AddRange(dataRows);
            if (fieldDefaults != null)
            {
                this._fieldDefaults = fieldDefaults;
            }
            if (security != null)
            {
                this.Security = security;
            }
            this.Folders.AddRange(folders);
        }

        #endregion

        #region Private Members
        private ContentTypeBindingCollection _ctBindings;
        private ViewCollection _views;
        private FieldCollection _fields;
        private FieldRefCollection _fieldRefs;
        private DataRowCollection _dataRows;
        private Dictionary<String, String> _fieldDefaults = new Dictionary<String, String>();
        private ObjectSecurity _security = null;
        private FolderCollection _folders;
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
        public ContentTypeBindingCollection ContentTypeBindings
        {
            get { return this._ctBindings; }
            private set { this._ctBindings = value; }
        }

        /// <summary>
        /// Gets or sets the content types to associate to the list
        /// </summary>
        public ViewCollection Views
        {
            get { return this._views; }
            private set { this._views = value; }
        }

        public FieldCollection Fields
        {
            get { return this._fields; }
            private set { this._fields = value; }
        }

        public FieldRefCollection FieldRefs
        {
            get { return this._fieldRefs; }
            private set { this._fieldRefs = value; }
        }

        public Guid TemplateFeatureID { get; set; }

        public DataRowCollection DataRows
        {
            get { return this._dataRows; }
            private set { this._dataRows = value; }
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
            set
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
        /// Defines a collection of folders (eventually nested) that 
        /// will be provisioned into the target list/library
        /// </summary>
        public FolderCollection Folders
        {
            get { return this._folders; }
            private set { this._folders = value; }
        }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}|{21}|{22}|{23}|{24}",
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
                this.FieldDefaults.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.Security != null ? this.Security.GetHashCode() : 0),
                this.Folders.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
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
            if (other == null)
            {
                return (false);
            }

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
                this.FieldDefaults.DeepEquals(other.FieldDefaults) &&
                (this.Security != null ? this.Security.Equals(other.Security) : true) &&
                this.Folders.DeepEquals(other.Folders)
                );
        }

        #endregion
    }
}
