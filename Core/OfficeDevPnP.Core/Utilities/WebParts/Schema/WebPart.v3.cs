namespace OfficeDevPnP.Core.Utilities.WebParts.Schema
{
    /// <summary>
    /// Contains WebParts element which contains list of webparts in the schema
    /// </summary>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("XmlSchemaClassGenerator", "1.0.0.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute("webParts", AnonymousType = true)]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlRootAttribute("webParts")]
    public partial class WebParts
    {

        /// <summary>
        /// Represents single WebPart element in the page
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("webPart", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public WebPart WebPart { get; set; }
    }

    /// <summary>
    /// Contains WebPart element attributes in the schema
    /// </summary>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("XmlSchemaClassGenerator", "1.0.0.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute("webPart", Namespace = "http://schemas.microsoft.com/WebPart/v3")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class WebPart
    {

        /// <summary>
        /// Represents metaData element of WebPart element
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("metaData", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public WebPartMetaData MetaData { get; set; }

        /// <summary>
        /// Represents data element of Webpart element
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("data", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public WebPartData Data { get; set; }
    }

    /// <summary>
    /// Contains webpart metadata attributes in the schema
    /// </summary>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("XmlSchemaClassGenerator", "1.0.0.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute("WebPartMetaData", Namespace = "http://schemas.microsoft.com/WebPart/v3", AnonymousType = true)]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class WebPartMetaData
    {

        /// <summary>
        /// Represents type element of webpart metadata
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("type", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public WebPartMetaDataType Type { get; set; }

        /// <summary>
        /// Represents importErrorMessage element of webpart metadata
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("importErrorMessage", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, DataType = "string")]
        public string ImportErrorMessage { get; set; }
    }

    /// <summary>
    /// Contains webpart metadata type element attributes
    /// </summary>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("XmlSchemaClassGenerator", "1.0.0.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute("WebPartMetaDataType", Namespace = "http://schemas.microsoft.com/WebPart/v3", AnonymousType = true)]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class WebPartMetaDataType
    {

        /// <summary>
        /// Represents name attribute of type element of webpart metadata
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute("name", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, DataType = "string")]
        public string Name { get; set; }

        /// <summary>
        /// Represents src attribute of type element of webpart metadata
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute("src", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, DataType = "string")]
        public string Src { get; set; }
    }

    /// <summary>
    /// Contains webpart data elements
    /// </summary>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("XmlSchemaClassGenerator", "1.0.0.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute("WebPartData", Namespace = "http://schemas.microsoft.com/WebPart/v3", AnonymousType = true)]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class WebPartData
    {

        /// <summary>
        /// Represents properties element of webpart data
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("properties", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PropertyContainerType Properties { get; set; }

        /// <summary>
        /// Represents genericWebPartProperties element of webpart data
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("genericWebPartProperties", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PropertyContainerType GenericWebPartProperties { get; set; }
    }

    /// <summary>
    /// Contains Property container type elements 
    /// </summary>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("XmlSchemaClassGenerator", "1.0.0.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute("propertyContainerType", Namespace = "http://schemas.microsoft.com/WebPart/v3")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class PropertyContainerType
    {

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        private System.Collections.ObjectModel.Collection<PropertyType> _property;

        /// <summary>
        /// Represents property element
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("property", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public System.Collections.ObjectModel.Collection<PropertyType> Property
        {
            get
            {
                return this._property;
            }
        }

        /// <summary>
        /// <para xml:lang="en">Gets a value indicating whether the Property collection is empty.</para>
        /// </summary>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool PropertySpecified
        {
            get
            {
                return (this.Property.Count != 0);
            }
        }

        /// <summary>
        /// <para xml:lang="en">Initializes a new instance of the <see cref="PropertyContainerType" /> class.</para>
        /// </summary>
        public PropertyContainerType()
        {
            this._property = new System.Collections.ObjectModel.Collection<PropertyType>();
            this._ipersonalizable = new System.Collections.ObjectModel.Collection<PropertyType>();
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        private System.Collections.ObjectModel.Collection<PropertyType> _ipersonalizable;

        /// <summary>
        /// Gets collection of webpart property types.
        /// </summary>
        [System.Xml.Serialization.XmlArrayAttribute("ipersonalizable", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        [System.Xml.Serialization.XmlArrayItemAttribute("property", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public System.Collections.ObjectModel.Collection<PropertyType> Ipersonalizable
        {
            get
            {
                return this._ipersonalizable;
            }
        }

        /// <summary>
        /// <para xml:lang="en">Gets a value indicating whether the Ipersonalizable collection is empty.</para>
        /// </summary>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool IpersonalizableSpecified
        {
            get
            {
                return (this.Ipersonalizable.Count != 0);
            }
        }
    }

    /// <summary>
    /// Contains Property type elements
    /// </summary>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("XmlSchemaClassGenerator", "1.0.0.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute("propertyType", Namespace = "http://schemas.microsoft.com/WebPart/v3")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class PropertyType
    {

        /// <summary>
        /// <para xml:lang="en">Gets or sets the text value.</para>
        /// </summary>
        [System.Xml.Serialization.XmlTextAttribute(DataType = "string")]
        public string Value { get; set; }

        /// <summary>
        /// Represents webpart name
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute("name", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, DataType = "string")]
        public string Name { get; set; }

        /// <summary>
        /// Represents webpart type
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute("type", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, DataType = "string")]
        public string Type { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Null property
        /// </summary>
        [System.Xml.Serialization.XmlAttributeAttribute("null", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, DataType = "boolean")]
        public bool Null { get; set; }

        /// <summary>
        /// <para xml:lang="en">Gets or sets a value indicating whether the Null property is specified.</para>
        /// </summary>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool NullSpecified { get; set; }
    }

    /// <summary>
    /// Contains properties of personalized webparts
    /// </summary>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("XmlSchemaClassGenerator", "1.0.0.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute("PropertyContainerTypeIpersonalizable", Namespace = "http://schemas.microsoft.com/WebPart/v3", AnonymousType = true)]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class PropertyContainerTypeIpersonalizable
    {

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        private System.Collections.ObjectModel.Collection<PropertyType> _property;

        /// <summary>
        /// Gets collection of webpart property types
        /// </summary>
        [System.Xml.Serialization.XmlElementAttribute("property", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public System.Collections.ObjectModel.Collection<PropertyType> Property
        {
            get
            {
                return this._property;
            }
        }

        /// <summary>
        /// <para xml:lang="en">Gets a value indicating whether the Property collection is empty.</para>
        /// </summary>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool PropertySpecified
        {
            get
            {
                return (this.Property.Count != 0);
            }
        }

        /// <summary>
        /// <para xml:lang="en">Initializes a new instance of the <see cref="PropertyContainerTypeIpersonalizable" /> class.</para>
        /// </summary>
        public PropertyContainerTypeIpersonalizable()
        {
            this._property = new System.Collections.ObjectModel.Collection<PropertyType>();
        }
    }
}
