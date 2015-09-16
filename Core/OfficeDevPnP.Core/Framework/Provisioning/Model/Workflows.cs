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
    public class Workflows
    {
        #region Private Members

        private List<WorkflowDefinition> _workflowDefinitions = new List<WorkflowDefinition>();
        private List<WorkflowSubscription> _workflowSubscriptions = new List<WorkflowSubscription>();

        #endregion

        #region Constructors

        public Workflows() { }

        public Workflows(IEnumerable<WorkflowDefinition> workflowDefinitions = null, IEnumerable<WorkflowSubscription> workflowSubscriptions = null)
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
        public List<WorkflowDefinition> WorkflowDefinitions
        {
            get { return this._workflowDefinitions; }
            private set { this._workflowDefinitions = value; }
        }

        /// <summary>
        /// Defines the Workflows Subscriptions to provision
        /// </summary>
        public List<WorkflowSubscription> WorkflowSubscriptions
        {
            get { return this._workflowSubscriptions; }
            private set { this._workflowSubscriptions = value; }
        }

        #endregion
    }
}
