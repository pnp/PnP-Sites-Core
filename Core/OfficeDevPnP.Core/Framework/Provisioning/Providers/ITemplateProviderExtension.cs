using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers
{
    /// <summary>
    /// Interface for extending the XMLTemplateProvider while retrieving a template
    /// </summary>
    public interface ITemplateProviderExtension
    {
        /// <summary>
        /// Initialization method to setup the extension object
        /// </summary>
        /// <param name="settings"></param>
        void Initialize(Object settings);

        /// <summary>
        /// Method invoked before deserializing the template from the source repository
        /// </summary>
        /// <param name="stream">The source stream</param>
        /// <returns>The resulting stream, after pre-processing</returns>
        Stream PreProcessGetTemplate(Stream stream);

        /// <summary>
        /// Method invoked after deserializing the template from the source repository
        /// </summary>
        /// <param name="template">The just deserialized template</param>
        /// <returns>The resulting template, after post-processing</returns>
        ProvisioningTemplate PostProcessGetTemplate(ProvisioningTemplate template);

        /// <summary>
        /// Method invoked before serializing the template and before it is saved onto the target repository
        /// </summary>
        /// <param name="template">The template that is going to be serialized</param>
        /// <returns>The resulting template, after pre-processing</returns>
        ProvisioningTemplate PreProcessSaveTemplate(ProvisioningTemplate template);

        /// <summary>
        /// Method invoked after serializing the template and before it is saved onto the target repository
        /// </summary>
        /// <param name="stream">The source stream</param>
        /// <returns>The resulting stream, after pre-processing</returns>
        Stream PostProcessSaveTemplate(Stream stream);

        /// <summary>
        /// Declares whether the object supports pre-processing during GetTemplate
        /// </summary>
        Boolean SupportsGetTemplatePreProcessing { get; }

        /// <summary>
        /// Declares whether the object supports post-processing during GetTemplate
        /// </summary>
        Boolean SupportsGetTemplatePostProcessing { get; }

        /// <summary>
        /// Declares whether the object supports pre-processing during SaveTemplate
        /// </summary>
        Boolean SupportsSaveTemplatePreProcessing { get; }

        /// <summary>
        /// Declares whether the object supports post-processing during SaveTemplate
        /// </summary>
        Boolean SupportsSaveTemplatePostProcessing { get; }
    }
}
