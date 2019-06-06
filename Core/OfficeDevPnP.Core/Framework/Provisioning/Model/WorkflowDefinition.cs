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
    /// Defines a Workflow Definition to provision
    /// </summary>
    public partial class WorkflowDefinition : BaseModel, IEquatable<WorkflowDefinition>
    {
        #region Private Members

        private Dictionary<String, String> _properties = new Dictionary<String, String>();

        #endregion

        #region Constructors
        /// <summary>
        /// Default Constructor
        /// </summary>
        public WorkflowDefinition() { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="properties">Dictionary of WorkflowDefinition properties</param>
        public WorkflowDefinition(Dictionary<String, String> properties)
        {
            if (properties != null)
            {
                this._properties = properties;
            }
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the Properties of the Workflows to provision
        /// </summary>
        public Dictionary<String, String> Properties
        {
            get { return this._properties; }
            private set {  this._properties = value; }
        }

        /// <summary>
        /// Defines the FormField XML of the Workflow to provision
        /// </summary>
        public String FormField { get; set; }

        /// <summary>
        /// Defines the ID of the Workflow Definition for the current Subscription
        /// </summary>
        public Guid Id { get; set; }

        /// <summary>
        /// Defines the URL of the Workflow Association page
        /// </summary>
        public String AssociationUrl { get; set; }

        /// <summary>
        /// The Description of the Workflow
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// The Display Name of the Workflow
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// Defines the DraftVersion of the Workflow, optional attribute.
        /// </summary>
        public String DraftVersion { get; set; }

        /// <summary>
        /// Defines the URL of the Workflow Initiation page
        /// </summary>
        public String InitiationUrl { get; set; }

        /// <summary>
        /// Defines if the Workflow is Published, optional attribute.
        /// </summary>
        public Boolean Published { get; set; }

        /// <summary>
        /// Defines if the Workflow requires the Association Form
        /// </summary>
        public Boolean RequiresAssociationForm { get; set; }

        /// <summary>
        /// Defines if the Workflow requires the Initiation Form
        /// </summary>
        public Boolean RequiresInitiationForm { get; set; }

        /// <summary>
        /// Defines the Scope Restriction for the Workflow
        /// </summary>
        public String RestrictToScope { get; set; }

        /// <summary>
        /// Defines the Type of Scope Restriction for the Workflow
        /// </summary>
        public String RestrictToType { get; set; }

        /// <summary>
        /// Defines path of the XAML of the Workflow to provision
        /// </summary>
        public String XamlPath { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Get hash code of WorkflowDefinition
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|",
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.FormField != null ? this.FormField.GetHashCode() : 0),
                (this.Id != null ? this.Id.GetHashCode() : 0),
                (this.AssociationUrl != null ? this.AssociationUrl.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                (this.DisplayName != null ? this.DisplayName.GetHashCode() : 0),
                (this.InitiationUrl != null ? this.InitiationUrl.GetHashCode() : 0),
                this.RequiresAssociationForm.GetHashCode(),
                this.RequiresInitiationForm.GetHashCode(),
                (this.RestrictToScope != null ? this.RestrictToScope.GetHashCode() : 0),
                (this.RestrictToType != null ? this.RestrictToType.GetHashCode() : 0),
                (this.XamlPath != null ? this.XamlPath.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares WorkflowDefinition with other WorkflowDefinition
        /// </summary>
        /// <param name="obj">WorkflowDefinition object</param>
        /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is WorkflowDefinition))
            {
                return (false);
            }
            return (Equals((WorkflowDefinition)obj));
        }

        /// <summary>
        /// Compares WorkflowDefinition with other WorkflowDefinition
        /// </summary>
        /// <param name="other">WorkflowDefinition object</param>
        /// <returns>true if the WorkflowDefinition object is equal to the current object; otherwise, false.</returns>
        public bool Equals(WorkflowDefinition other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.Properties.DeepEquals(other.Properties) &&
                this.FormField == other.FormField &&
                this.Id == other.Id &&
                this.AssociationUrl == other.AssociationUrl &&
                this.Description == other.Description &&
                this.DisplayName == other.DisplayName &&
                this.InitiationUrl == other.InitiationUrl &&
                this.RequiresAssociationForm == other.RequiresAssociationForm &&
                this.RequiresInitiationForm == other.RequiresInitiationForm &&
                this.RestrictToScope == other.RestrictToScope &&
                this.RestrictToType == other.RestrictToType &&
                this.XamlPath == other.XamlPath
                );
        }

        #endregion
    }
}
