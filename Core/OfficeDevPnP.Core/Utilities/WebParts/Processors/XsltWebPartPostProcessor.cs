using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Utilities.WebParts.Schema;
using WebPart = OfficeDevPnP.Core.Utilities.WebParts.Schema.WebPart;

namespace OfficeDevPnP.Core.Utilities.WebParts.Processors
{
    /// <summary>
    /// Updates view for XsltListViewWebPart using schema definition provided
    /// Instead of using default view for XsltListViewWebPart, it tries to resolve view from schema and updates hidden view created by XsltListViewWebPart
    /// </summary>
    public class XsltWebPartPostProcessor : IWebPartPostProcessor
    {
        private readonly IList<PropertyType> _properties;

        /// <summary>
        /// Constructor for XsltWebPartPostProcessor class
        /// </summary>
        /// <param name="schema">Webpart object</param>
        public XsltWebPartPostProcessor(WebPart schema)
        {
            _properties = schema.Data.Properties.Property.ToList();
        }

        /// <summary>
        /// Method to process webpart
        /// </summary>
        /// <param name="wpDefinition">WebPartDefinition object</param>
        /// <param name="webPartPage">File object</param>
        public void Process(WebPartDefinition wpDefinition, File webPartPage)
        {
            var web = GetWeb(webPartPage);
            var list = GetList(web);

            if (list == null)
            {
                return;
            }

            var xsltHiddenView = list.GetViewById(wpDefinition.Id);
            if (xsltHiddenView == null)
            {
                return;
            }

            var listView = GetViewFromSchemaProperties(list);

            if (listView == null)
            {
                UpdateHiddenViewFromWebPartSchema(xsltHiddenView);
            }
            else
            {
                UpdateHiddenView(xsltHiddenView, listView.ListViewXml);
            }
        }

        private void UpdateHiddenViewFromWebPartSchema(View xsltHiddenView)
        {
            var xmlDefinitionProperty = GetProperty("XmlDefinition");
            if (!string.IsNullOrEmpty(xmlDefinitionProperty?.Value))
            {
                UpdateHiddenView(xsltHiddenView, xmlDefinitionProperty.Value);
            }
        }

        private void UpdateHiddenView(View xsltHiddenView, string schemaXml)
        {
            var viewSchemaElement = XElement.Parse(schemaXml);
            var xml = string.Concat(viewSchemaElement.Elements());
            xsltHiddenView.ListViewXml = xml;
            xsltHiddenView.Update();
            xsltHiddenView.Context.ExecuteQueryRetry();
        }

        private View GetViewFromSchemaProperties(List list)
        {
            var viewIdProperty = GetProperty("ViewId");
            Guid viewId;
            if (TryParseGuidProperty(viewIdProperty, out viewId))
            {
                return list.GetViewById(viewId);
            }

            var viewGuidProperty = GetProperty("ViewGuid");
            if (TryParseGuidProperty(viewGuidProperty, out viewId))
            {
                return list.GetViewById(viewId);
            }

            var viewNameProperty = GetProperty("ViewName");

            if (!string.IsNullOrEmpty(viewNameProperty?.Value))
            {
                return list.GetViewByName(viewNameProperty.Value);
            }

            var viewUrlProperty = GetProperty("ViewUrl");

            if (!string.IsNullOrEmpty(viewUrlProperty?.Value))
            {
                View view;
                if (TryGetView(() =>
                {
                    list.Context.Load(list.Views);
                    list.Context.ExecuteQueryRetry();

                    foreach (View listView in list.Views)
                    {
                        if (!listView.Hidden &&
                            listView.ServerRelativeUrl.IndexOf(viewUrlProperty.Value, StringComparison.OrdinalIgnoreCase) !=
                            -1)
                        {
                            return listView;
                        }
                    }

                    throw new Exception("View not found");
                }, out view))
                {
                    return view;
                }
            }

            var xmlDefinitionProperty = GetProperty("XmlDefinition");

            if (!string.IsNullOrEmpty(xmlDefinitionProperty?.Value))
            {
                var viewSchemaElement = XElement.Parse(xmlDefinitionProperty.Value);
                var nameAttribute = viewSchemaElement.Attribute("Name");

                View view;
                if (nameAttribute != null && TryGetView(() =>
                {
                    var listView = list.Views.GetById(new Guid(nameAttribute.Value));
                    list.Context.Load(listView);
                    list.Context.ExecuteQueryRetry();

                    return listView;
                }, out view))
                {
                    return view;
                }

                var displayNameAttribute = viewSchemaElement.Attribute("DisplayName");
                if (!string.IsNullOrEmpty(displayNameAttribute?.Value) && TryGetView(() =>
                {
                    var listView = list.Views.GetByTitle(displayNameAttribute.Value);
                    list.Context.Load(listView);
                    list.Context.ExecuteQueryRetry();

                    return listView;
                }, out view))
                {
                    return view;
                }

                var urlAttribute = viewSchemaElement.Attribute("Url");

                if (urlAttribute != null && TryGetView(() =>
                {
                    list.Context.Load(list.Views);
                    list.Context.ExecuteQueryRetry();

                    foreach (View listView in list.Views)
                    {
                        if (!listView.Hidden && 
                        listView.ServerRelativeUrl.IndexOf(urlAttribute.Value, StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            return listView;
                        }
                    }

                    throw new Exception("View not found");

                }, out view))
                {
                    return view;
                }

            }

            return null;
        }

        private bool TryGetView(Func<View> action, out View view)
        {
            try
            {
                view = action();
                return true;
            }
            catch (Exception)
            {
                view = null;
                return false;
            }
        }

        private List GetList(Web web)
        {
            var listUrlProperty = GetProperty("ListUrl");

            if (!string.IsNullOrEmpty(listUrlProperty?.Value))
            {
                return web.GetListByUrl(listUrlProperty.Value);
            }

            var listIdProperty = GetProperty("ListId");
            Guid listId;
            if (TryParseGuidProperty(listIdProperty, out listId))
            {
                return web.Lists.GetById(listId);
            }

            var listNameProperty = GetProperty("ListName");
            if (TryParseGuidProperty(listNameProperty, out listId))
            {
                return web.Lists.GetById(listId);
            }

            var listDisplayName = GetProperty("ListDisplayName");
            if (!string.IsNullOrEmpty(listDisplayName?.Value))
            {
                return web.GetListByTitle(listDisplayName.Value);
            }

            return null;
        }

        private Web GetWeb(File webPartPage)
        {
            var context = (ClientContext)webPartPage.Context;

            var webIdProperty = GetProperty("WebId");
            Guid webId;

            if (TryParseGuidProperty(webIdProperty, out webId))
            {
                return context.Site.OpenWebById(webId);
            }

            return ((ClientContext) webPartPage.Context).Web;
        }

        private bool TryParseGuidProperty(PropertyType property, out Guid id)
        {
            if (!string.IsNullOrEmpty(property?.Value) && Guid.TryParse(property.Value, out id) && !id.Equals(Guid.Empty))
            {
                return true;
            }
            id = Guid.Empty;
            return false;
        }

        private PropertyType GetProperty(string name)
        {
            return _properties.FirstOrDefault(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
    }
}
