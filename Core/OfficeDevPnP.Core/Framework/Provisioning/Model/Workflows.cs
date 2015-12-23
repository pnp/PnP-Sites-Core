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

        private WorkflowDefinitionsCollection _workflowDefinitions;
        private WorkflowSubscriptionsCollection _workflowSubscriptions;

        #endregion

        #region Constructors

        public Workflows()
        {
            this._workflowDefinitions = new Model.WorkflowDefinitionsCollection(this.ParentTemplate);
            this._workflowSubscriptions = new Model.WorkflowSubscriptionsCollection(this.ParentTemplate);
        }

        public Workflows(IEnumerable<WorkflowDefinition> workflowDefinitions = null, IEnumerable<WorkflowSubscription> workflowSubscriptions = null) : this()
        {
            if (workflowDefinitions != null)
            {
                this._workflowDefinitions.AddRange(workflowDefinitions);
            }
            if (workflowSubscriptions != null)
            {
                this._workflowSubscriptions.AddRange(workflowSubscriptions);
            }
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the Workflows Definitions to provision
        /// </summary>
        public WorkflowDefinitionsCollection WorkflowDefinitions
        {
            get { return this._workflowDefinitions; }
            private set { this._workflowDefinitions = value; }
        }

        /// <summary>
        /// Defines the Workflows Subscriptions to provision
        /// </summary>
        public WorkflowSubscriptionsCollection WorkflowSubscriptions
        {
            get { return this._workflowSubscriptions; }
            private set { this._workflowSubscriptions = value; }
        }

        #endregion
    }
}
