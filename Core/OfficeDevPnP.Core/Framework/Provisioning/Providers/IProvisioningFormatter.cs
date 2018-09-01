//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace OfficeDevPnP.Core.Framework.Provisioning.Providers
//{
//    /// <summary>
//    /// Interface for basic capabilites that any Provisioning Formatter should provide/support
//    /// </summary>
//    public interface IProvisioningFormatter
//    {
//        /// <summary>
//        /// Method to validate the content of a formatted provisioning instance
//        /// </summary>
//        /// <param name="provisioning">The formatted provisioning instance as a Stream</param>
//        /// <returns>Boolean result of the validation</returns>
//        Boolean IsValidProvisioning(Stream provisioning);

//        /// <summary>
//        /// Method to format a Provisioning instance into a formatted provisioning
//        /// </summary>
//        /// <param name="provisioning">The input Provisioning object</param>
//        /// <returns>The output formatted provisioning as a Stream</returns>
//        Stream ToFormattedProvisioning(Model.ProvisioningHierarchy provisioning);

//        /// <summary>
//        /// Method to convert a formatted provisioning into a Provisioning object
//        /// </summary>
//        /// <param name="template">The input formatted provisioning as a Stream</param>
//        /// <returns>The output Provisioning object</returns>
//        Model.ProvisioningHierarchy ToProvisioning(Stream template);
//    }
//}
