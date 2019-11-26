using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers
{
    interface ITemplateFormatterWithValidation : ITemplateFormatter
    {
        /// <summary>
        /// Method to validate the content of a formatted template instance
        /// </summary>
        /// <param name="template">The formatted template instance as a Stream</param>
        /// <returns>A validation result of the validation</returns>
        ValidationResult GetValidationResults(Stream template);
    }
}
