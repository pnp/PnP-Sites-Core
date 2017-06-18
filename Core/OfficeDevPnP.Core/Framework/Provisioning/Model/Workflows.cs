using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Workflows to provision
    /// </summary>
    public partial class Workflows: BaseModel
    {
        #region Private Members

        private WorkflowDefinitionCollection _workflowDefinitions;
        private WorkflowSubscriptionCollection _workflowSubscriptions;

        #endregion

        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public Workflows()
        {
            this._workflowDefinitions = new Model.WorkflowDefinitionCollection(this.ParentTemplate);
            this._workflowSubscriptions = new Model.WorkflowSubscriptionCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="workflowDefinitions">Collection of workflow definitions</param>
        /// <param name="workflowSubscriptions">Collection of workflow subscriptions</param>
        public Workflows(IEnumerable<WorkflowDefinition> workflowDefinitions = null, IEnumerable<WorkflowSubscription> workflowSubscriptions = null) : 
            this()
        {
            this.WorkflowDefinitions.AddRange(workflowDefinitions);
            this.WorkflowSubscriptions.AddRange(workflowSubscriptions);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the Workflows Definitions to provision
        /// </summary>
        public WorkflowDefinitionCollection WorkflowDefinitions
        {
            get { return this._workflowDefinitions; }
            private set { this._workflowDefinitions = value; }
        }

        /// <summary>
        /// Defines the Workflows Subscriptions to provision
        /// </summary>
        public WorkflowSubscriptionCollection WorkflowSubscriptions
        {
            get { return this._workflowSubscriptions; }
            private set { this._workflowSubscriptions = value; }
        }

        #endregion
    }
}
