using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Navigation type from schema to model
    /// </summary>
    internal class NavigationFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        private String _navigationType;

        public NavigationFromSchemaToModelTypeResolver(String navigationType)
        {
            this._navigationType = navigationType;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object target = null;
            Object targetNavigation = null;
            var targetIsGlobal = false;

            Object navigation = null;

            switch (this._navigationType)
            {
                case "GlobalNavigation":
                    navigation = source.GetPublicInstancePropertyValue("GlobalNavigation");
                    targetIsGlobal = true;
                    break;
                case "CurrentNavigation":
                    navigation = source.GetPublicInstancePropertyValue("CurrentNavigation");
                    targetIsGlobal = false;
                    break;
            }

            if (navigation != null)
            {
                var navigationType = navigation.GetPublicInstancePropertyValue("NavigationType");
                switch (navigationType.ToString())
                {
                    case "Managed":
                        targetNavigation = new Model.ManagedNavigation();

                        var managedNavigation = navigation.GetPublicInstancePropertyValue("ManagedNavigation");
                        PnPObjectsMapper.MapProperties(managedNavigation, targetNavigation, resolvers, true);

                        if (targetIsGlobal)
                        {
                            target = new Model.GlobalNavigation(Model.GlobalNavigationType.Managed, null, (Model.ManagedNavigation)targetNavigation);
                        }
                        else
                        {
                            target = new Model.CurrentNavigation(Model.CurrentNavigationType.Managed, null, (Model.ManagedNavigation)targetNavigation);
                        }

                        break;
                    case "Structural":
                    case "StructuralLocal":
                        targetNavigation = new Model.StructuralNavigation();
                        var structuralNavigation = navigation.GetPublicInstancePropertyValue("StructuralNavigation");
                        var structuralNavigationNodes = structuralNavigation.GetPublicInstancePropertyValue("NavigationNode");

                        if (!resolvers.ContainsKey($"{targetNavigation.GetType().FullName}.NavigationNodes"))
                        {
                            resolvers.Add($"{targetNavigation.GetType().FullName}.NavigationNodes", new NavigationNodeFromSchemaToModelTypeResolver());
                        }

                        PnPObjectsMapper.MapProperties(structuralNavigation, targetNavigation, resolvers, true);

                        if (targetIsGlobal)
                        {
                            target = new Model.GlobalNavigation(Model.GlobalNavigationType.Structural, (Model.StructuralNavigation)targetNavigation, null);
                        }
                        else
                        {
                            target = new Model.CurrentNavigation(
                                navigationType.ToString() == "Structural" ? Model.CurrentNavigationType.Structural : Model.CurrentNavigationType.StructuralLocal,
                                (Model.StructuralNavigation)targetNavigation, null);
                        }

                        break;
                    case "Inherit":
                        if (targetIsGlobal)
                        {
                            target = new Model.GlobalNavigation(Model.GlobalNavigationType.Inherit, null, null);
                        }
                        else
                        {
                            target = new Model.CurrentNavigation(Model.CurrentNavigationType.Inherit, null, null);
                        }
                        break;
                }
                return (target);
            }
            return null;
        }
    }
}
