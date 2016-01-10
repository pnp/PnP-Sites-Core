using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a Workflow Subscription to provision
    /// </summary>
    public partial class WorkflowSubscription : BaseModel, IEquatable<WorkflowSubscription>
    {
        #region Private Members

        private Dictionary<String, String> _propertyDefinitions = new Dictionary<String, String>();
        private List<String> _eventTypes = new List<String>();

        #endregion

        #region Constructors

        public WorkflowSubscription() { }

        public WorkflowSubscription(Dictionary<String, String> propertyDefinitions)
        {
            if (propertyDefinitions != null)
            {
                this._propertyDefinitions = propertyDefinitions;
            }
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the Property Definitions of the Workflows to provision
        /// </summary>
        public Dictionary<string, string> PropertyDefinitions
        {
            get { return this._propertyDefinitions; }
            private set { this._propertyDefinitions = value; }
        }

        /// <summary>
        /// Defines the ID of the Workflow Definition for the current Subscription
        /// </summary>
        public Guid DefinitionId { get; set; }

        /// <summary>
        /// Defines the ID of the target list/library for the current Subscription, 
        /// </summary>
        /// <remarks>
        /// Optional and if it is missing, the workflow subscription will 
        /// be at Site level
        /// </remarks>
        public String ListId { get; set; }

        /// <summary>
        /// Defines if the Workflow Definition is enabled for the current Subscription
        /// </summary>
        public Boolean Enabled { get; set; }

        /// <summary>
        /// Defines the ID of the Event Source for the current Subscription
        /// </summary>
        public String EventSourceId { get; set; }

        /// <summary>
        /// Defines the list of events that will start the workflow instance
        /// </summary>
        /// <remarks>
        /// Possible values in the list: WorkflowStartEvent, ItemAddedEvent, ItemUpdatedEvent
        /// </remarks>
        public List<String> EventTypes
        {
            get { return this._eventTypes; }
            set { this._eventTypes = value; }
        }

        /// <summary>
        /// Defines if the Workflow can be manually started bypassing the activation limit
        /// </summary>
        public Boolean ManualStartBypassesActivationLimit { get; set; }

        /// <summary>
        /// Defines the Name of the Workflow Subscription
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// Defines the Parent ContentType Id of the Workflow Subscription
        /// </summary>
        public String ParentContentTypeId { get; set; }

        /// <summary>
        /// Defines the Status Field Name of the Workflow Subscription
        /// </summary>
        public String StatusFieldName { get; set; }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|",
                this.PropertyDefinitions.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.DefinitionId != null ? this.DefinitionId.GetHashCode() : 0),
                (this.ListId != null ? this.ListId.GetHashCode() : 0),
                this.Enabled.GetHashCode(),
                (this.EventSourceId != null ? this.EventSourceId.GetHashCode() : 0),
                this.EventTypes.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.ManualStartBypassesActivationLimit.GetHashCode(),
                (this.Name != null ? this.Name.GetHashCode() : 0),
                (this.ParentContentTypeId != null ? this.ParentContentTypeId.GetHashCode() : 0),
                (this.StatusFieldName != null ? this.StatusFieldName.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is WorkflowSubscription))
            {
                return (false);
            }
            return (Equals((WorkflowSubscription)obj));
        }

        public bool Equals(WorkflowSubscription other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.PropertyDefinitions.DeepEquals(other.PropertyDefinitions) &&
                this.DefinitionId == other.DefinitionId &&
                this.ListId == other.ListId &&
                this.Enabled == other.Enabled &&
                this.EventSourceId == other.EventSourceId &&
                this.EventTypes.DeepEquals(other.EventTypes) &&
                this.ManualStartBypassesActivationLimit == other.ManualStartBypassesActivationLimit &&
                this.Name == other.Name &&
                this.ParentContentTypeId == other.ParentContentTypeId &&
                this.StatusFieldName == other.StatusFieldName
            );
        }

        #endregion
    }
}
