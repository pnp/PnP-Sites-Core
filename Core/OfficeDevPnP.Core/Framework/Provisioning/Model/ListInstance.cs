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
        /// <summary>
        /// Constructor for ListInstance class
        /// </summary>
        public ListInstance()
        {
            this._ctBindings = new ContentTypeBindingCollection(this.ParentTemplate);
            this._views = new ViewCollection(this.ParentTemplate);
            this._fields = new FieldCollection(this.ParentTemplate);
            this._fieldRefs = new FieldRefCollection(this.ParentTemplate);
            this._dataRows = new DataRowCollection(this.ParentTemplate);
            this._folders = new FolderCollection(this.ParentTemplate);
            this._userCustomActions = new CustomActionCollection(this.ParentTemplate);
            this._webhooks = new WebhookCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor for ListInstance class
        /// </summary>
        /// <param name="contentTypeBindings">ContentType Bindings of the list</param>
        /// <param name="views">Views of the list</param>
        /// <param name="fields">Fields of the list</param>
        /// <param name="fieldRefs">FieldRefs of the list</param>
        /// <param name="dataRows">DataRows of the list</param>
        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows) :
                this(contentTypeBindings, views, fields, fieldRefs, dataRows, null, null, null)
        {
        }

        /// <summary>
        /// Constructor for ListInstance class
        /// </summary>
        /// <param name="contentTypeBindings">ContentType Bindings of the list</param>
        /// <param name="views">View of the list</param>
        /// <param name="fields">Fields of the list</param>
        /// <param name="fieldRefs">FieldRefs of the list</param>
        /// <param name="dataRows">DataRows of the list</param>
        /// <param name="fieldDefaults">FieldDefaults of the list</param>
        /// <param name="security">Security Rules of the list</param>
        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows, Dictionary<String, String> fieldDefaults, ObjectSecurity security) :
                this(contentTypeBindings, views, fields, fieldRefs, dataRows, fieldDefaults, security, null)
        {
        }

        /// <summary>
        /// Constructor for ListInstance class
        /// </summary>
        /// <param name="contentTypeBindings">ContentTypeBindings  of the list</param>
        /// <param name="views">Views of the list</param>
        /// <param name="fields">Fields of the list</param>
        /// <param name="fieldRefs">FieldRefs of the list</param>
        /// <param name="dataRows">DataRows of the list</param>
        /// <param name="fieldDefaults">FieldDefaults of the list</param>
        /// <param name="security">Security Rules of the list</param>
        /// <param name="folders">List Folders</param>
        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows, Dictionary<String, String> fieldDefaults, ObjectSecurity security, List<Folder> folders) :
                this(contentTypeBindings, views, fields, fieldRefs, dataRows, fieldDefaults, security, folders, null)
        {
        }

        /// <summary>
        /// Constructor for the ListInstance class
        /// </summary>
        /// <param name="contentTypeBindings">ContentTypeBindings of the list</param>
        /// <param name="views">Views of the list</param>
        /// <param name="fields">Fields of the list</param>
        /// <param name="fieldRefs">FieldRefs of the list</param>
        /// <param name="dataRows">DataRows of the list</param>
        /// <param name="fieldDefaults">FieldDefaults of the list</param>
        /// <param name="security">Security Rules of the list</param>
        /// <param name="folders">List Folders</param>
        /// <param name="userCustomActions">UserCustomActions of the list</param>
        public ListInstance(IEnumerable<ContentTypeBinding> contentTypeBindings,
            IEnumerable<View> views, IEnumerable<Field> fields, IEnumerable<FieldRef> fieldRefs, List<DataRow> dataRows, Dictionary<String, String> fieldDefaults, ObjectSecurity security, List<Folder> folders, List<CustomAction> userCustomActions) :
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
            this.UserCustomActions.AddRange(userCustomActions);
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
        private CustomActionCollection _userCustomActions;
        private WebhookCollection _webhooks;
        private IRMSettings _IRMSettings;
        private Dictionary<String, String> _dataSource = new Dictionary<String, String>();
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
        /// Gets or sets whether to force checkout of documents in the library
        /// </summary>
        public bool ForceCheckout { get; set; } = false;

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
        /// Gets or sets the views associated to the list
        /// </summary>
        public ViewCollection Views
        {
            get { return this._views; }
            private set { this._views = value; }
        }

        /// <summary>
        /// Gets or sets the Fields associated to the list
        /// </summary>
        public FieldCollection Fields
        {
            get { return this._fields; }
            private set { this._fields = value; }
        }

        /// <summary>
        /// Gets or sets the FieldRefs associated to the list
        /// </summary>
        public FieldRefCollection FieldRefs
        {
            get { return this._fieldRefs; }
            private set { this._fieldRefs = value; }
        }

        /// <summary>
        /// Gets or sets the Guid for TemplateFeature
        /// </summary>
        public Guid TemplateFeatureID { get; set; }

        /// <summary>
        /// Gets or sets the DataRows associated to the list
        /// </summary>
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

        /// <summary>
        /// Defines a collection of user custom actions that 
        /// will be provisioned into the target list/library
        /// </summary>
        public CustomActionCollection UserCustomActions
        {
            get { return this._userCustomActions; }
            private set { this._userCustomActions = value; }
        }

        public WebhookCollection Webhooks
        {
            get { return this._webhooks; }
            private set { this._webhooks = value; }
        }

        public IRMSettings IRMSettings
        {
            get { return this._IRMSettings; }
            set
            {
                if (this._IRMSettings != null)
                {
                    this._IRMSettings.ParentTemplate = null;
                }
                this._IRMSettings = value;
                if (this._IRMSettings != null)
                {
                    this._IRMSettings.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Defines if the current list or library has to be included in crawling, optional attribute.
        /// </summary>
        public Boolean NoCrawl { get; set; }

        /// <summary>
        /// Defines the current list UI/UX experience (valid for SPO only).
        /// </summary>
        public ListExperience ListExperience { get; set; }

        /// <summary>
        /// Defines a value that specifies the location of the default display form for the list.
        /// </summary>
        public String DefaultDisplayFormUrl { get; set; }

        /// <summary>
        /// Defines a value that specifies the URL of the edit form to use for list items in the list.
        /// </summary>
        public String DefaultEditFormUrl { get; set; }

        /// <summary>
        /// Defines a value that specifies the location of the default new form for the list.
        /// </summary>
        public String DefaultNewFormUrl { get; set; }

        /// <summary>
        /// Defines a value that specifies the reading order of the list.
        /// </summary>
        public ListReadingDirection Direction { get; set; }

        /// <summary>
        /// Defines a value that specifies the URI for the icon of the list, optional attribute.
        /// </summary>
        public String ImageUrl { get; set; }

        /// <summary>
        /// Defines if IRM Expire property, optional attribute.
        /// </summary>
        public Boolean IrmExpire { get; set; }

        /// <summary>
        /// Defines the IRM Reject property, optional attribute.
        /// </summary>
        public Boolean IrmReject { get; set; }

        /// <summary>
        /// Defines a value that specifies a flag that a client application can use to determine whether to display the list, optional attribute.
        /// </summary>
        public Boolean IsApplicationList { get; set; }

        /// <summary>
        /// Defines the Read Security property, optional attribute.
        /// </summary>
        public Int32 ReadSecurity { get; set; }

        /// <summary>
        /// Defines the Write Security property, optional attribute.
        /// </summary>
        public Int32 WriteSecurity { get; set; }

        /// <summary>
        /// Defines a value that specifies the data validation criteria for a list item, optional attribute.
        /// </summary>
        public String ValidationFormula { get; set; }

        /// <summary>
        /// Defines a value that specifies the error message returned when data validation fails for a list item, optional attribute.
        /// </summary>
        public String ValidationMessage { get; set; }

        /// <summary>
        /// Defines a list of Data Source properties for the List Instance
        /// </summary>
        public Dictionary<String, String> DataSource
        {
            get { return this._dataSource; }
            private set { this._dataSource = value; }
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}|{21}|{22}|{23}|{24}|{25}|{26}|{27}|{28}|{29}|{30}|{31}|{32}|{33}|{34}|{35}|{36}|{37}|{38}|{39}|{40}|{41}|{42}|{43}|",
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
                this.ForceCheckout.GetHashCode(),
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
                this.Folders.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.UserCustomActions.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Webhooks.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                (this.IRMSettings != null ? this.IRMSettings.GetHashCode() : 0),
                this.NoCrawl.GetHashCode(),
                this.ListExperience.GetHashCode(),
                this.DefaultDisplayFormUrl?.GetHashCode() ?? 0,
                this.DefaultEditFormUrl?.GetHashCode() ?? 0,
                this.DefaultNewFormUrl?.GetHashCode() ?? 0,
                this.Direction.GetHashCode(),
                this.ImageUrl?.GetHashCode() ?? 0,
                this.IrmExpire.GetHashCode(),
                this.IrmReject.GetHashCode(),
                this.IsApplicationList.GetHashCode(),
                this.ReadSecurity.GetHashCode(),
                this.ValidationFormula?.GetHashCode() ?? 0,
                this.ValidationMessage?.GetHashCode() ?? 0,
                this.DataSource.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.WriteSecurity.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ListInstance
        /// </summary>
        /// <param name="obj">Object that represents ListInstance</param>
        /// <returns>true if the current object is equal to the ListInstance</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ListInstance))
            {
                return (false);
            }
            return (Equals((ListInstance)obj));
        }

        /// <summary>
        /// Compares ListInstance object based on ContentTypesEnabled, Description, DocumentTemplate, EnableVersioning, EnableMinorVersions, EnableModeration, Hidden, 
        /// MaxVersionLimit, MinorVersionLimit, OnQuickLaunch, EnableAttachments, EnableFolderCreation, ForceCheckOut, RemoveExistingContentTypes, TemplateType,
        /// Title, Url, TemplateFeatureID, RemoveExistingViews, ContentTypeBindings, View, Fields, FieldRefs, FieldDefaults, Security, Folders, UserCustomActions, 
        /// Webhooks, IRMSettings, DefaultDisplayFormUrl, DefaultEditFormUrl, DefaultNewFormUrl, Direction, ImageUrl, IrmExpire, IrmReject, IsApplicationList,
        /// ReadSecurity, ValidationFormula, ValidationMessage, DataSource, and WriteSecurity properties.
        /// </summary>
        /// <param name="other">ListInstance object</param>
        /// <returns>true if the ListInstance object is equal to the current object; otherwise, false.</returns>
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
                this.ForceCheckout == other.ForceCheckout &&
                this.RemoveExistingContentTypes == other.RemoveExistingContentTypes &&
                this.TemplateType == other.TemplateType &&
                this.Title == other.Title &&
                this.Url == other.Url &&
                this.TemplateFeatureID == other.TemplateFeatureID &&
                this.RemoveExistingViews == other.RemoveExistingViews &&
                this.ContentTypeBindings.DeepEquals(other.ContentTypeBindings) &&
                // Only do a deep view compare on non system lists to avoid subtle changes in OOB view XML to popup system lists in the generated model
                (this.IsApplicationList == false ? this.Views.DeepEquals(other.Views) : true) &&
                // Only do a deep field compare on non system lists to avoid subtle changes in OOB field XML to popup system lists in the generated model
                (this.IsApplicationList == false ? this.Fields.DeepEquals(other.Fields) : true) &&
                this.FieldRefs.DeepEquals(other.FieldRefs) &&
                this.FieldDefaults.DeepEquals(other.FieldDefaults) &&
                ((this.Security != null && other.Security != null) ? this.Security.Equals(other.Security) : true) &&
                this.Folders.DeepEquals(other.Folders) &&
                this.UserCustomActions.DeepEquals(other.UserCustomActions) &&
                this.Webhooks.DeepEquals(other.Webhooks) &&
                (this.IRMSettings != null ? this.IRMSettings.Equals(other.IRMSettings) : true) &&
                this.NoCrawl == other.NoCrawl &&
                this.ListExperience == other.ListExperience &&
                this.DefaultDisplayFormUrl == other.DefaultDisplayFormUrl &&
                this.DefaultEditFormUrl == other.DefaultEditFormUrl &&
                this.DefaultNewFormUrl == other.DefaultNewFormUrl &&
                this.Direction == other.Direction &&
                this.ImageUrl == other.ImageUrl &&
                this.IrmExpire == other.IrmExpire &&
                this.IrmReject == other.IrmReject &&
                this.IsApplicationList == other.IsApplicationList &&
                this.ReadSecurity == other.ReadSecurity &&
                this.ValidationFormula == other.ValidationFormula &&
                this.ValidationMessage == other.ValidationMessage &&
                this.DataSource.DeepEquals(other.DataSource) &&
                this.WriteSecurity == other.WriteSecurity
            );
        }

        #endregion
    }

    public enum ListExperience
    {
        /// <summary>
        ///  SPO will automatically define the right experience based on the settings of the current list, it is the default value.
        /// </summary>
        Auto,
        /// <summary>
        /// The Classic experience will be forced for the current list.
        /// </summary>
        ClassicExperience,
        /// <summary>
        /// The Modern experience will be forced for the current list.
        /// </summary>
        NewExperience,
    }

    public enum ListReadingDirection
    {
        /// <summary>
        /// None
        /// </summary>
        None,
        /// <summary>
        /// Left to Right
        /// </summary>
        LTR,
        /// <summary>
        /// Right to Left
        /// </summary>
        RTL,
    }
}
