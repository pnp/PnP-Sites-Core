using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Pages;

namespace OfficeDevPnP.Core.Utilities.CanvasControl.Processors
{
#if !ONPREMISES
    /// <summary>
    /// Updates list id for List web part, to allow provision based on url in a dynamic provisioning scenario
    /// </summary>
    public class ListControlPostProcessor : ICanvasControlPostProcessor
    {
        private readonly IDictionary<string, object> _properties;

        /// <summary>
        /// Constructor for ListControlPostProcessor class
        /// </summary>
        /// <param name="control">Client control</param>
        public ListControlPostProcessor(Framework.Provisioning.Model.CanvasControl control)
        {
            _properties = JsonUtility.Deserialize<Dictionary<string, dynamic>>(control.JsonControlData);
        }


        /// <summary>
        /// Method for processing canvas control
        /// </summary>
        /// <param name="canvasControl">Canvas control object</param>
        /// <param name="clientSidePage">ClientSidePage object</param>
        public void Process(Framework.Provisioning.Model.CanvasControl canvasControl, ClientSidePage clientSidePage)
        {
            var web = GetWeb(clientSidePage);
            var list = GetList(web);

            if (list == null)
            {
                return;
            }

            list.EnsureProperties(l => l.Id, l => l.RootFolder, l => l.RootFolder.Name);

            SetProperty("selectedListId", list.Id);
            SetProperty("selectedListUrl", list.RootFolder.Name);

            canvasControl.JsonControlData = JsonUtility.Serialize(_properties);
        }

        private List GetList(Web web)
        {
            // grab list based on url
            var listUrlProperty = GetProperty("selectedListUrl") as string;
            if (!string.IsNullOrWhiteSpace(listUrlProperty))
            {
                return web.GetList(listUrlProperty);
            }

            // grab list based on list id
            var listIdProperty = GetProperty("selectedListId") as string;
            Guid listId;
            if (TryParseGuidProperty(listIdProperty, out listId))
            {
                return web.Lists.GetById(listId);
            }

            // grab list based on list title
            var listDisplayName = GetProperty("listTitle") as string;
            if (!string.IsNullOrWhiteSpace(listDisplayName))
            {
                return web.GetListByTitle(listDisplayName);
            }

            return null;
        }

        private Web GetWeb(ClientSidePage clientSidePage)
        {
            return clientSidePage.Context.Web;
        }

        private bool TryParseGuidProperty(string guid, out Guid id)
        {
            if (!string.IsNullOrWhiteSpace(guid) && Guid.TryParse(guid, out id) && !id.Equals(Guid.Empty))
            {
                return true;
            }

            id = Guid.Empty;
            return false;
        }

        private object GetProperty(string name)
        {
            object value;
            return _properties.TryGetValue(name, out value) ? value : null;
        }

        private void SetProperty(string name, object value)
        {
            _properties[name] = value;
        }
    }
#endif
}