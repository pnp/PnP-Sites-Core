using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of DataRow objects
    /// </summary>
    public partial class DataRowCollection : ProvisioningTemplateCollection<DataRow>
    {
        /// <summary>
        /// Optional attribute to declare the name of the Key Column, if any, used to identify any already existing DataRows.
        /// </summary>
        public String KeyColumn { get; set; }

        /// <summary>
        /// If the DataRow already exists on target list, this attribute defines whether 
        /// the DataRow will be overwritten or skipped.
        /// </summary>
        public UpdateBehavior UpdateBehavior { get; set; }

        /// <summary>
        /// Constructor for DataRowCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public DataRowCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }

    public enum UpdateBehavior
    {
        /// <summary>
        /// Any existing DataRow will be overwritten.
        /// </summary>
        Overwrite,
        /// <summary>
        /// Any existing DataRow will be skipped.
        /// </summary>
        Skip,   
    }
}
