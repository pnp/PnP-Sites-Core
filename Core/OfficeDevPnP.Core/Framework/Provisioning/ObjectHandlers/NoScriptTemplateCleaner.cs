using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    public class NoScriptTemplateCleaner
    {
        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        private Web _web;
        public NoScriptTemplateCleaner(Web web)
        {
            _web = web;
        }
        public ProvisioningTemplate CleanUpBeforeProvisioning(ProvisioningTemplate template)
        {
            bool isNoScriptSite = _web.IsNoScriptSite();

            var listsToRemove = new List<ListInstance>();
            
            foreach(var templateList in template.Lists)
            { 
                if (isNoScriptSite && templateList.Url == "Style Library")
                {
                    listsToRemove.Add(templateList);
                    WriteMessage(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_List__0__is_Style_Library_of_NoScript_will_Skip, templateList.Title), ProvisioningMessageType.Warning);
                }
            }
            while(listsToRemove.Count > 0)
            {
                var listToRemove = listsToRemove[0];
                template.Lists.Remove(listToRemove);
                listsToRemove.RemoveAt(0);
            }

            return template;
        }

        public ProvisioningTemplate CleanUpAfterExtraction(ProvisioningTemplate template)
        {
            bool isNoScriptSite = _web.IsNoScriptSite();

            var listsToRemove = new List<ListInstance>();

            foreach (var templateList in template.Lists)
            {
                if (isNoScriptSite && templateList.Url == "Style Library")
                {
                    listsToRemove.Add(templateList);
                    WriteMessage(string.Format(CoreResources.Provisioning_ObjectHandlers_ListInstances_List__0__is_Style_Library_of_NoScript_will_Skip, templateList.Title), ProvisioningMessageType.Warning);
                }
            }
            while (listsToRemove.Count > 0)
            {
                var listToRemove = listsToRemove[0];
                template.Lists.Remove(listToRemove);
                listsToRemove.RemoveAt(0);
            }

            return template;
        }



        internal void WriteMessage(string message, ProvisioningMessageType messageType)
        {
            if (MessagesDelegate != null)
            {
                MessagesDelegate(message, messageType);
            }
        }
    }
}
