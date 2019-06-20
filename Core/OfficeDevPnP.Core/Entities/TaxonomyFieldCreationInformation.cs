using Microsoft.SharePoint.Client.Taxonomy;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Hodlds properties for Taxonomy filed
    /// </summary>
    public class TaxonomyFieldCreationInformation : FieldCreationInformation
    {
        private bool _multiValue = false;
       
        /// <summary>
        /// Allows multiple values for Taxonomy field
        /// </summary>
        public bool MultiValue 
        {
            get {
                return _multiValue;
            }
            set
            {
                if (value)
                {
                    FieldType = "TaxonomyFieldTypeMulti";
                }
                else
                {
                    FieldType = "TaxonomyFieldType";
                }
                _multiValue = value;
            }
        }

        /// <summary>
        /// Represents an item in the TermStore
        /// </summary>
        public TaxonomyItem TaxonomyItem { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public TaxonomyFieldCreationInformation()
            : base("TaxonomyFieldType")
        { }

    }

}
