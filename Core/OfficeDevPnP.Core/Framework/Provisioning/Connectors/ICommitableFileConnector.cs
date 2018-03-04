using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors
{
    /// <summary>
    /// Interface for File Connectors
    /// </summary>
    public interface ICommitableFileConnector
    {
        /// <summary>
        /// Commits the file
        /// </summary>
        void Commit();
    }
}
