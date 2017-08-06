using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Navigation type from model to schema
    /// </summary>
    internal class NavigationFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        private String _navigationType;

        public NavigationFromModelToSchemaTypeResolver(String navigationType)
        {
            this._navigationType = navigationType;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object target = null;
            Model.BaseNavigationKind navigation = null;

            var globalNavigationTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.NavigationGlobalNavigation, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var globalNavigationType = Type.GetType(globalNavigationTypeName, true);
            var globalNavigationTypeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.NavigationGlobalNavigationNavigationType, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var globalNavigationTypeType = Type.GetType(globalNavigationTypeTypeName, true);
            var currentNavigationTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.NavigationCurrentNavigation, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var currentNavigationType = Type.GetType(currentNavigationTypeName, true);
            var currentNavigationTypeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.NavigationCurrentNavigationNavigationType, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var currentNavigationTypeType = Type.GetType(currentNavigationTypeTypeName, true);
            var managedNavigationTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ManagedNavigation, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var managedNavigationType = Type.GetType(managedNavigationTypeName, true);
            var structuralNavigationTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StructuralNavigation, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var structuralNavigationType = Type.GetType(structuralNavigationTypeName, true);
            
            var modelSource = source as Model.Navigation;
            if (modelSource != null)
            {
                switch (this._navigationType)
                {
                    case "GlobalNavigation":
                        navigation = modelSource.GlobalNavigation;

                        target = Activator.CreateInstance(globalNavigationType);
                        target.SetPublicInstancePropertyValue("NavigationType", Enum.Parse(globalNavigationTypeType, modelSource.GlobalNavigation.NavigationType.ToString()));

                        switch (modelSource.GlobalNavigation.NavigationType)
                        {
                            case Model.GlobalNavigationType.Managed:
                                var managedNavigation = Activator.CreateInstance(managedNavigationType);

                                PnPObjectsMapper.MapProperties(modelSource.GlobalNavigation.ManagedNavigation, managedNavigation, resolvers, true);
                                target.SetPublicInstancePropertyValue("ManagedNavigation", managedNavigation);

                                break;
                            case Model.GlobalNavigationType.Structural:
                                var structuralNavigation = Activator.CreateInstance(structuralNavigationType);

                                if (!resolvers.ContainsKey($"{structuralNavigation.GetType().FullName}.NavigationNode"))
                                {
                                    resolvers.Add($"{structuralNavigation.GetType().FullName}.NavigationNode", new NavigationNodeFromModelToSchemaTypeResolver());
                                }
                                PnPObjectsMapper.MapProperties(modelSource.GlobalNavigation.StructuralNavigation, structuralNavigation, resolvers, true);
                                target.SetPublicInstancePropertyValue("StructuralNavigation", structuralNavigation);

                                break;
                            case Model.GlobalNavigationType.Inherit:
                                break;
                        }

                        break;
                    case "CurrentNavigation":
                        navigation = modelSource.CurrentNavigation;

                        target = Activator.CreateInstance(currentNavigationType);
                        target.SetPublicInstancePropertyValue("NavigationType", Enum.Parse(currentNavigationTypeType, modelSource.CurrentNavigation.NavigationType.ToString()));

                        switch (modelSource.CurrentNavigation.NavigationType)
                        {
                            case Model.CurrentNavigationType.Managed:
                                var managedNavigation = Activator.CreateInstance(managedNavigationType);

                                PnPObjectsMapper.MapProperties(modelSource.CurrentNavigation.ManagedNavigation, managedNavigation, resolvers, true);
                                target.SetPublicInstancePropertyValue("ManagedNavigation", managedNavigation);

                                break;
                            case Model.CurrentNavigationType.Structural:
                            case Model.CurrentNavigationType.StructuralLocal:
                                var structuralNavigation = Activator.CreateInstance(structuralNavigationType);

                                if (!resolvers.ContainsKey($"{structuralNavigation.GetType().FullName}.NavigationNode"))
                                {
                                    resolvers.Add($"{structuralNavigation.GetType().FullName}.NavigationNode", new NavigationNodeFromModelToSchemaTypeResolver());
                                }
                                PnPObjectsMapper.MapProperties(modelSource.CurrentNavigation.StructuralNavigation, structuralNavigation, resolvers, true);
                                target.SetPublicInstancePropertyValue("StructuralNavigation", structuralNavigation);

                                break;
                            case Model.CurrentNavigationType.Inherit:
                                break;
                        }

                        break;
                }
            }

            return (target);
        }
    }
}
