using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class SearchsettingsValidator : ValidatorBase
    {
        public bool Validate(string source, string target)
        {
            bool isValid = false;

            SeacrhConfigSettings sourceSeacrhConfigSettings = GetSearchConfigSettings(source);
            SeacrhConfigSettings targetSeacrhConfigSettings = GetSearchConfigSettings(target);

            ///Comparing the source XML and target XML
            SearchQueryConfigurationSettings sQuery = sourceSeacrhConfigSettings.SearchQueryConfigurationSettings;
            SearchQueryConfigurationSettings tQuery = targetSeacrhConfigSettings.SearchQueryConfigurationSettings;

            //null check for Best bets elemen in source and target
            if (sQuery.BestBet != null)
            {
                if (!ValidateObjects(sQuery.BestBet, tQuery.BestBet, new List<string> { "Title", "Url" })) { return false; }
            }

            if (sQuery.QueryRules != null)
            {
                if (!ValidateObjects(sQuery.QueryRules, tQuery.QueryRules, new List<string> { "DisplayName", "IsActive", "IsBestBetOnly" })) { return false; }
            }

            if (sQuery.ResultTypes != null)
            {
                if (!ValidateObjects(sQuery.ResultTypes, tQuery.ResultTypes, new List<string> { "Name", "DisplayTemplateUrl", "IsDeleted" })) { return false; }
            }

            if (sQuery.Sources != null)
            {
                if (!ValidateObjects(sQuery.Sources, tQuery.Sources, new List<string> { "Name", "Active" })) { return false; }
            }

            SearchSchemaConfigSettings sSchema = sourceSeacrhConfigSettings.SearchSchemaConfigSettings;
            SearchSchemaConfigSettings tSchema = targetSeacrhConfigSettings.SearchSchemaConfigSettings;

            if (sSchema.SeacrhScehmaAliases.AliasDictionary != null)
            {
                if (!ValidateObjects(sSchema.SeacrhScehmaAliases.AliasDictionary, tSchema.SeacrhScehmaAliases.AliasDictionary, new List<string> { "Key", "Name" })) { return false; }
            }

            int sCount = 0;
            int tCount = 0;
            foreach (CrawledPropertyInfoCollection sourceCrawlProperty in sSchema.CategoriesAndCrawledProperties)
            {
                sCount++;
                foreach (CrawledPropertyInfoCollection targetCrawlProperty in tSchema.CategoriesAndCrawledProperties)
                {
                    if (sourceCrawlProperty.SearchSchemaConfigSettings_Dictionary != null)
                    {
                        isValid = ValidateObjects(sourceCrawlProperty.SearchSchemaConfigSettings_Dictionary, targetCrawlProperty.SearchSchemaConfigSettings_Dictionary, new List<string> { "Key", "Name", "CategoryName" });
                        if (isValid) { tCount++; break; }
                    }
                }
            }
            if (sCount != tCount) { return false; }


            if (sSchema.ManagedProperties.ManagedPropertiesDictionary != null)
            {
                if (!ValidateObjects(sSchema.ManagedProperties.ManagedPropertiesDictionary, tSchema.ManagedProperties.ManagedPropertiesDictionary, new List<string> { "Key", "Name" })) { return false; }
            }


            if (sSchema.Mappings.MappingsDictionary != null)
            {
                if (!ValidateObjects(sSchema.Mappings.MappingsDictionary, tSchema.Mappings.MappingsDictionary, new List<string> { "Key", "Name", "CrawledPropertyName", "CrawledPropset" })) { return false; }
            }

            if (sSchema.Overrides.OverridesDictionary != null)
            {
                if (!ValidateObjects(sSchema.Overrides.OverridesDictionary, tSchema.Overrides.OverridesDictionary, new List<string> { "Key", "Name" })) { return false; }
            }

            SearchRankingModelConfigurationSettings srcRankingModel = sourceSeacrhConfigSettings.SearchRankingModelConfigurationSettings;
            SearchRankingModelConfigurationSettings targetRankingModel = targetSeacrhConfigSettings.SearchRankingModelConfigurationSettings;

            isValid = XmlComparer.AreEqual(srcRankingModel.RankingModel, targetRankingModel.RankingModel).Success;
            if (!isValid) { return false; }

            return isValid;
        }

        private SeacrhConfigSettings GetSearchConfigSettings(string xml)
        {
            XmlDocument xDoc = LoadXML(xml);
            SeacrhConfigSettings searchConfigSettings = new SeacrhConfigSettings();
            searchConfigSettings.SearchQueryConfigurationSettings = GetSeacrhQueryConfiguration(xDoc, "SearchQueryConfigurationSettings");
            searchConfigSettings.SearchSchemaConfigSettings = GetSeacrhSchemaConfiguration(xDoc, "SearchSchemaConfigurationSettings");
            searchConfigSettings.SearchRankingModelConfigurationSettings = GetSearchRankingModelConfigurationSettings(xDoc, "SearchRankingModelConfigurationSettings", "RankingModels");
            return searchConfigSettings;
        }

        #region SeacrhQueryConfigurations
        private SearchQueryConfigurationSettings GetSeacrhQueryConfiguration(XmlDocument xDoc, string tagName)
        {
            SearchQueryConfigurationSettings searchQueryConfigurationSettings = new SearchQueryConfigurationSettings();
            foreach (XmlNode sNode in xDoc.GetElementsByTagName(tagName))
            {
                foreach (XmlNode sourceNode in sNode.ChildNodes)
                {
                    foreach (XmlNode srcNode in sourceNode.ChildNodes)
                    {
                        if (srcNode.Name == "BestBets")
                        {
                            if (srcNode.ChildNodes.Count > 0)
                            {
                                searchQueryConfigurationSettings.BestBet = GetBestBetsFromSearchQueryConfigurationSettings(srcNode);
                            }
                            continue;
                        }
                        if (srcNode.Name == "DefaultSourceId")
                        {
                            searchQueryConfigurationSettings.DefaultSourceId = GetSearchQueryConfigurationSettingsElementValues(srcNode);
                            continue;
                        }
                        if (srcNode.Name == "DefaultSourceIdSet")
                        {
                            searchQueryConfigurationSettings.DefaultSourceIdSet = GetSearchQueryConfigurationSettingsElementValues(srcNode);
                            continue;
                        }
                        if (srcNode.Name == "DeployToParent")
                        {
                            searchQueryConfigurationSettings.DeployToParent = GetSearchQueryConfigurationSettingsElementValues(srcNode);
                            continue;
                        }
                        if (srcNode.Name == "DisableInheritanceOnImport")
                        {
                            searchQueryConfigurationSettings.DisableInheritanceOnImport = GetSearchQueryConfigurationSettingsElementValues(srcNode);
                            continue;
                        }
                        if (srcNode.Name == "QueryRules")
                        {
                            if (srcNode.ChildNodes.Count > 0)
                            {
                                searchQueryConfigurationSettings.QueryRules = GetQueryRulesFromSearchQueryConfigurationSettings(srcNode);
                            }
                            continue;
                        }
                        if (srcNode.Name == "ResultTypes")
                        {
                            if (srcNode.ChildNodes.Count > 0)
                            {
                                searchQueryConfigurationSettings.ResultTypes = GetResultTypesFromSearchQueryConfigurationSettings(srcNode);
                            }
                            continue;
                        }
                        if (srcNode.Name == "Sources")
                        {
                            if (srcNode.ChildNodes.Count > 0)
                            {
                                searchQueryConfigurationSettings.Sources = GetSourcesFromSearchQueryConfigurationSettings(srcNode);
                            }
                            continue;
                        }
                    }
                }
                //Getting 2 child elements when calling GetElementsByTagName with SearchQueryConfigurationSettings element so will break 
                //for second loop by checking the searchQueryConfigurationSettings class object with null case
                if (searchQueryConfigurationSettings != null) break;
            }
            return searchQueryConfigurationSettings;
        }
        private List<BestBets> GetBestBetsFromSearchQueryConfigurationSettings(XmlNode srcNode)
        {
            List<BestBets> lstBestBets = new List<BestBets>();
            BestBets bestBets = new BestBets();
            foreach (XmlNode srcBestBetNode in srcNode.ChildNodes)
            {
                bestBets = new BestBets();
                foreach (XmlNode srcBestBet in srcBestBetNode.ChildNodes)
                {

                    if (srcBestBet.LocalName == "_Title")
                    {
                        bestBets.Title = srcBestBet.InnerText;
                        continue;
                    }
                    if (srcBestBet.LocalName == "_Url")
                    {
                        bestBets.Url = srcBestBet.InnerText;
                        continue;
                    }
                }
                lstBestBets.Add(bestBets);
            }
            return lstBestBets;
        }
        private List<QueryRule> GetQueryRulesFromSearchQueryConfigurationSettings(XmlNode srcNode)
        {
            List<QueryRule> lstQueryRules = new List<QueryRule>();
            QueryRule queryRule = new QueryRule();
            //for SearchQueryConfigurationSettngs
            foreach (XmlNode srcQueryRuleNode in srcNode.ChildNodes)
            {
                queryRule = new QueryRule();
                foreach (XmlNode srcQueryRule in srcQueryRuleNode.ChildNodes)
                {
                    if (srcQueryRule.LocalName == "_DisplayName")
                    {
                        queryRule.DisplayName = srcQueryRule.InnerText;
                        continue;
                    }
                    if (srcQueryRule.LocalName == "IsActive")
                    {
                        queryRule.IsActive = Convert.ToBoolean(srcQueryRule.InnerText);
                        continue;
                    }
                    if (srcQueryRule.LocalName == "_IsBestBetOnly")
                    {
                        queryRule.IsBestBetOnly = Convert.ToBoolean(srcQueryRule.InnerText);
                        continue;
                    }
                }

                lstQueryRules.Add(queryRule);
            }

            return lstQueryRules;
        }

        private List<ResultType> GetResultTypesFromSearchQueryConfigurationSettings(XmlNode srcNode)
        {
            List<ResultType> lstResultTypes = new List<ResultType>();
            ResultType resultType = null;
            foreach (XmlNode srcResultypeNode in srcNode.ChildNodes)
            {
                resultType = new ResultType();
                foreach (XmlNode srcResultType in srcResultypeNode.ChildNodes)
                {
                    if (srcResultType.LocalName == "DisplayTemplateUrl")
                    {
                        resultType.DisplayTemplateUrl = srcResultType.InnerText;
                        continue;
                    }
                    if (srcResultType.LocalName == "Name")
                    {
                        resultType.Name = srcResultType.InnerText;
                        continue;
                    }
                    if (srcResultType.LocalName == "IsDeleted")
                    {
                        resultType.IsDeleted = Convert.ToBoolean(srcResultType.InnerText);
                        continue;
                    }
                }

                lstResultTypes.Add(resultType);
            }

            return lstResultTypes;
        }

        private List<Source> GetSourcesFromSearchQueryConfigurationSettings(XmlNode srcNode)
        {
            List<Source> lstSources = new List<Source>();
            Source source = null;
            foreach (XmlNode srcSourceNode in srcNode.ChildNodes)
            {
                source = new Source();
                foreach (XmlNode srcSource in srcSourceNode.ChildNodes)
                {
                    if (srcSource.LocalName == "Name")
                    {
                        source.Name = srcSource.InnerText;
                        continue;
                    }
                    if (srcSource.LocalName == "Active")
                    {
                        source.Active = Convert.ToBoolean(srcSource.InnerText);
                        continue;
                    }
                    if (srcSource.LocalName == "ProviderId")
                    {
                        source.ProviderId = new Guid(srcSource.InnerText);
                        continue;
                    }
                }

                lstSources.Add(source);
            }

            return lstSources;
        }

        private static string GetSearchQueryConfigurationSettingsElementValues(XmlNode srcNode)
        {
            string attributeValue = string.Empty;
            if (srcNode.Attributes.Count > 0)
            {
                foreach (XmlAttribute srcAttr in srcNode.Attributes)
                {
                    attributeValue = srcAttr.OuterXml;
                    break;
                }
            }
            else
            {
                attributeValue = srcNode.InnerText;
            }

            return attributeValue;
        }

        #endregion

        #region SearchRankingModelConfigurationSettings
        private SearchRankingModelConfigurationSettings GetSearchRankingModelConfigurationSettings(XmlDocument xDoc, string tagName, string elementName)
        {
            SearchRankingModelConfigurationSettings SearchRankingSettings = new SearchRankingModelConfigurationSettings();
            string rankingModelXML = string.Empty;
            bool isElementFound = false;
            //for SearchRankingModelConfigurationSettings
            foreach (XmlNode sNode in xDoc.GetElementsByTagName(tagName))
            {
                foreach (XmlNode sourceNode in sNode.ChildNodes)
                {
                    if (sourceNode.Name == elementName)
                    {
                        rankingModelXML = sourceNode.OuterXml;
                        isElementFound = true;
                        break;

                    }
                    if (isElementFound) break;
                }
                if (isElementFound) break;
            }
            SearchRankingSettings.RankingModel = rankingModelXML;
            return SearchRankingSettings;
        }
        #endregion

        #region SeacrchScehma
        private SearchSchemaConfigSettings GetSeacrhSchemaConfiguration(XmlDocument xDoc, string tagName)
        {
            SearchSchemaConfigSettings searchSchemaConfigurationSettings = new SearchSchemaConfigSettings();

            foreach (XmlNode sNode in xDoc.GetElementsByTagName(tagName))
            {
                foreach (XmlNode sourceNode in sNode.ChildNodes)
                {
                    if (sourceNode.Name == "Aliases")
                    {
                        foreach (XmlNode srcNode in sourceNode.ChildNodes)
                        {
                            if (srcNode.LocalName == "LastItemName")
                            {
                                searchSchemaConfigurationSettings.SeacrhScehmaAliases = new Aliases();
                                searchSchemaConfigurationSettings.SeacrhScehmaAliases.LastItemName = srcNode.InnerText;
                                continue;
                            }
                            if (srcNode.LocalName == "dictionary")
                            {
                                searchSchemaConfigurationSettings.SeacrhScehmaAliases.AliasDictionary = GetSeacrhSchemeaConfigurationAliasesSettings(srcNode);
                                continue;
                            }
                        }
                    }
                    if (sourceNode.Name == "CategoriesAndCrawledProperties")
                    {
                        searchSchemaConfigurationSettings.CategoriesAndCrawledProperties = GetSearchSchemaCrawledProperties(sourceNode);
                        continue;
                    }
                    if (sourceNode.Name == "ManagedProperties")
                    {
                        searchSchemaConfigurationSettings.ManagedProperties = GetSearchSchemaManagedProperties(sourceNode);
                        continue;
                    }
                    if (sourceNode.Name == "Mappings")
                    {
                        searchSchemaConfigurationSettings.Mappings = GetSearchSchemaMappings(sourceNode);
                        continue;
                    }
                    if (sourceNode.Name == "Overrides")
                    {
                        searchSchemaConfigurationSettings.Overrides = GetSearchSchemaOverrides(sourceNode);
                        continue;
                    }
                }
                //Geetting 2 child elements when calling GetElementsByTagName with SearchQueryConfigurationSettings element so will break 
                //for second loop by checking the searchQueryConfigurationSettings class object with null case
                if (searchSchemaConfigurationSettings != null) break;
            }

            return searchSchemaConfigurationSettings;
        }

        private List<Aliases_Dictionary> GetSeacrhSchemeaConfigurationAliasesSettings(XmlNode srcNode)
        {
            List<Aliases_Dictionary> lstSearchSchemaConfigSettingsAliases = new List<Aliases_Dictionary>();

            Aliases_Dictionary SearchSchemaConfigSettingsAliases = null;
            foreach (XmlNode srcSourceNode in srcNode.ChildNodes)
            {
                SearchSchemaConfigSettingsAliases = new Aliases_Dictionary();
                foreach (XmlNode srcSource in srcSourceNode.ChildNodes)
                {

                    if (srcSource.LocalName == "Key")
                    {
                        SearchSchemaConfigSettingsAliases.Key = srcSource.InnerText;
                        continue;
                    }
                    if (srcSource.LocalName == "Value")
                    {
                        foreach (XmlNode srcAlias in srcSource.ChildNodes)
                        {
                            if (srcAlias.LocalName == "Name")
                            {
                                SearchSchemaConfigSettingsAliases.Name = srcAlias.InnerText;
                                continue;
                            }
                            if (srcAlias.LocalName == "ManagedPid")
                            {
                                SearchSchemaConfigSettingsAliases.ManagedPid = srcAlias.InnerText;
                                continue;
                            }
                            if (srcAlias.LocalName == "SchemaId")
                            {
                                SearchSchemaConfigSettingsAliases.SchemaId = srcAlias.InnerText;
                                continue;
                            }
                        }
                    }
                }
                lstSearchSchemaConfigSettingsAliases.Add(SearchSchemaConfigSettingsAliases);
            }
            return lstSearchSchemaConfigSettingsAliases;
        }

        private List<CrawledPropertyInfoCollection> GetSearchSchemaCrawledProperties(XmlNode srcNode)
        {
            List<CrawledPropertyInfoCollection> lstCategoriesAndCrawledProperties = new List<CrawledPropertyInfoCollection>();
            CrawledPropertyInfoCollection CategoriesAndCrawledProperties = null;
            SearchSchemaConfigSettings_Dictionary schemaDic = null;
            List<SearchSchemaConfigSettings_Dictionary> lstschemaDic = null;

            foreach (XmlNode srcSource in srcNode.ChildNodes)
            {
                CategoriesAndCrawledProperties = new CrawledPropertyInfoCollection();
                foreach (XmlNode srcCrawledSource in srcSource.ChildNodes)
                {
                    if (srcCrawledSource.LocalName == "Key")
                    {
                        CategoriesAndCrawledProperties.Key = new Guid(srcCrawledSource.InnerText);
                        continue;
                    }
                    if (srcCrawledSource.LocalName == "Value")
                    {
                        foreach (XmlNode srcCrwaled in srcCrawledSource.ChildNodes)
                        {
                            if (srcCrwaled.LocalName == "LastItemName")
                            {
                                CategoriesAndCrawledProperties.LastItemName = srcCrwaled.InnerText;
                                continue;
                            }
                            if (srcCrwaled.LocalName == "dictionary")
                            {
                                lstschemaDic = new List<SearchSchemaConfigSettings_Dictionary>();
                                foreach (XmlNode srcCrwaledString in srcCrwaled.ChildNodes)
                                {
                                    schemaDic = new SearchSchemaConfigSettings_Dictionary();
                                    foreach (XmlNode srcCrwaledStringKey in srcCrwaledString.ChildNodes)
                                    {
                                        if (srcCrwaledStringKey.LocalName == "Key")
                                        {
                                            schemaDic.Key = srcCrwaledStringKey.InnerText;
                                            continue;
                                        }
                                        if (srcCrwaledStringKey.LocalName == "Value")
                                        {
                                            foreach (XmlNode srcCrwaledStringKeyNode in srcCrwaledStringKey.ChildNodes)
                                            {
                                                if (srcCrwaledStringKeyNode.LocalName == "Name")
                                                {
                                                    schemaDic.Name = srcCrwaledStringKeyNode.InnerText;
                                                    continue;
                                                }
                                                if (srcCrwaledStringKeyNode.LocalName == "CategoryName")
                                                {
                                                    schemaDic.CategoryName = srcCrwaledStringKeyNode.InnerText;
                                                    continue;
                                                }
                                                if (srcCrwaledStringKeyNode.LocalName == "IsImplicit")
                                                {
                                                    schemaDic.IsImplicit = Convert.ToBoolean(srcCrwaledStringKeyNode.InnerText);
                                                    continue;
                                                }
                                                if (srcCrwaledStringKeyNode.LocalName == "IsMappedToContents")
                                                {
                                                    schemaDic.IsMappedToContents = Convert.ToBoolean(srcCrwaledStringKeyNode.InnerText);
                                                    continue;
                                                }
                                                if (srcCrwaledStringKeyNode.LocalName == "IsNameEnum")
                                                {
                                                    schemaDic.IsNameEnum = Convert.ToBoolean(srcCrwaledStringKeyNode.InnerText);
                                                    continue;
                                                }
                                            }
                                        }

                                    }
                                    lstschemaDic.Add(schemaDic);
                                }
                                continue;
                            }
                        }
                    }
                }
                CategoriesAndCrawledProperties.SearchSchemaConfigSettings_Dictionary = lstschemaDic;
                lstCategoriesAndCrawledProperties.Add(CategoriesAndCrawledProperties);
            }

            return lstCategoriesAndCrawledProperties;
        }

        private static ManagedProperties GetSearchSchemaManagedProperties(XmlNode srcNode)
        {
            ManagedProperties managedProperties = new ManagedProperties();
            ManagedProperties_Dictionary managedProperty = null;
            foreach (XmlNode srcSourceNode in srcNode.ChildNodes)
            {
                if (srcSourceNode.LocalName == "LastItemName")
                {
                    managedProperties.LastItemName = srcSourceNode.InnerText;
                    continue;
                }
                if (srcSourceNode.ParentNode.LocalName == "TotalCount")
                {
                    managedProperties.TotalCount = Convert.ToInt32(srcSourceNode.InnerText);
                    continue;
                }
                if (srcSourceNode.LocalName == "dictionary")
                {
                    managedProperties.ManagedPropertiesDictionary = new List<ManagedProperties_Dictionary>();

                    foreach (XmlNode srcSource in srcSourceNode.ChildNodes)
                    {
                        if (srcSource.HasChildNodes)
                        {
                            managedProperty = new ManagedProperties_Dictionary();

                            foreach (XmlNode srcManagedProperty in srcSource.ChildNodes)
                            {
                                if (srcManagedProperty.LocalName == "Key")
                                {
                                    managedProperty.Key = srcManagedProperty.InnerText;
                                    continue;
                                }
                                foreach (XmlNode srcManagedPropertyValuedict in srcManagedProperty.ChildNodes)
                                {
                                    if (srcManagedPropertyValuedict.LocalName == "Name")
                                    {
                                        managedProperty.Name = srcManagedPropertyValuedict.InnerText;
                                        continue;
                                    }
                                    if (srcManagedPropertyValuedict.LocalName == "IsImplicit")
                                    {
                                        managedProperty.IsImplicit = Convert.ToBoolean(srcManagedPropertyValuedict.InnerText);
                                        continue;
                                    }
                                }
                                managedProperties.ManagedPropertiesDictionary.Add(managedProperty);
                            }
                        }
                    }
                }

            }

            return managedProperties;
        }

        private static Mappings GetSearchSchemaMappings(XmlNode srcNode)
        {
            Mappings mappings = new Mappings();
            Mappings_Dictionary mappingProperty = null;
            foreach (XmlNode srcSourceNode in srcNode.ChildNodes)
            {
                if (srcSourceNode.LocalName == "LastItemName")
                {
                    mappings.LastItemName = srcSourceNode.InnerText;
                    continue;
                }
                if (srcSourceNode.LocalName == "dictionary")
                {
                    mappings.MappingsDictionary = new List<Mappings_Dictionary>();

                    foreach (XmlNode srcSource in srcSourceNode.ChildNodes)
                    {
                        if (srcSource.HasChildNodes)
                        {
                            mappingProperty = new Mappings_Dictionary();

                            foreach (XmlNode srcManagedProperty in srcSource.ChildNodes)
                            {
                                if (srcManagedProperty.LocalName == "Key")
                                {
                                    mappingProperty.Key = srcManagedProperty.InnerText;
                                    continue;
                                }
                                foreach (XmlNode srcManagedPropertyValuedict in srcManagedProperty.ChildNodes)
                                {
                                    if (srcManagedPropertyValuedict.LocalName == "Name")
                                    {
                                        mappingProperty.Name = srcManagedPropertyValuedict.InnerText;
                                        continue;
                                    }
                                    if (srcManagedPropertyValuedict.LocalName == "CrawledPropertyName")
                                    {
                                        mappingProperty.CrawledPropertyName = srcManagedPropertyValuedict.InnerText;
                                        continue;
                                    }
                                    if (srcManagedPropertyValuedict.LocalName == "CrawledPropset")
                                    {
                                        mappingProperty.CrawledPropset = srcManagedPropertyValuedict.InnerText;
                                        continue;
                                    }
                                    if (srcManagedPropertyValuedict.LocalName == "ManagedPid")
                                    {
                                        mappingProperty.ManagedPid = srcManagedPropertyValuedict.InnerText;
                                        continue;
                                    }
                                    if (srcManagedPropertyValuedict.LocalName == "MappingOrder")
                                    {
                                        mappingProperty.MappingOrder = srcManagedPropertyValuedict.InnerText;
                                        continue;
                                    }
                                    if (srcManagedPropertyValuedict.LocalName == "SchemaId")
                                    {
                                        mappingProperty.SchemaId = srcManagedPropertyValuedict.InnerText;
                                        continue;
                                    }
                                }
                                mappings.MappingsDictionary.Add(mappingProperty);
                            }
                        }
                    }
                }
            }

            return mappings;
        }

        private static Overrides GetSearchSchemaOverrides(XmlNode srcNode)
        {
            Overrides overrides = new Overrides();
            Overrides_Dictionary overridesProperty = null;
            foreach (XmlNode srcSourceNode in srcNode.ChildNodes)
            {
                if (srcSourceNode.LocalName == "LastItemName")
                {
                    overrides.LastItemName = srcSourceNode.InnerText;
                    continue;
                }
                if (srcSourceNode.LocalName == "dictionary")
                {
                    overrides.OverridesDictionary = new List<Overrides_Dictionary>();

                    foreach (XmlNode srcSource in srcSourceNode.ChildNodes)
                    {
                        if (srcSource.HasChildNodes)
                        {
                            overridesProperty = new Overrides_Dictionary();

                            foreach (XmlNode srcManagedProperty in srcSource.ChildNodes)
                            {
                                if (srcManagedProperty.LocalName == "Key")
                                {
                                    overridesProperty.Key = srcManagedProperty.InnerText;
                                    continue;
                                }
                                foreach (XmlNode srcManagedPropertyValuedict in srcManagedProperty.ChildNodes)
                                {
                                    if (srcManagedPropertyValuedict.LocalName == "Name")
                                    {
                                        overridesProperty.Name = srcManagedPropertyValuedict.InnerText;
                                        continue;
                                    }
                                }
                                overrides.OverridesDictionary.Add(overridesProperty);
                            }
                        }
                    }
                }
            }

            return overrides;
        }
        #endregion

        private static XmlDocument LoadXML(string xml)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);
            return xmlDoc;
        }
        #region SeacrhSettings Classes
        private class QueryRule
        {
            public string DisplayName { get; set; }
            public bool IsActive { get; set; }
            public bool IsBestBetOnly { get; set; }
        }

        private class ResultType
        {
            public string DisplayTemplateUrl { get; set; }
            public string Name { get; set; }
            public bool IsDeleted { get; set; }
        }

        private class Source
        {
            public bool Active { get; set; }
            public string Name { get; set; }
            public Guid ProviderId { get; set; }
        }

        private class BestBets
        {
            public string Title { get; set; }
            public string Url { get; set; }
        }

        private class Aliases
        {
            public string LastItemName { get; set; }
            public List<Aliases_Dictionary> AliasDictionary { get; set; }
        }

        private class Aliases_Dictionary
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public string ManagedPid { get; set; }
            public string SchemaId { get; set; }
        }

        private class CrawledPropertyInfoCollection
        {
            public Guid Key { get; set; }
            public string LastItemName { get; set; }
            public List<SearchSchemaConfigSettings_Dictionary> SearchSchemaConfigSettings_Dictionary { get; set; }
        }

        private class SearchSchemaConfigSettings_Dictionary
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public string CategoryName { get; set; }
            public bool IsImplicit { get; set; }
            public bool IsMappedToContents { get; set; }
            public bool IsNameEnum { get; set; }
        }

        private class ManagedProperties
        {
            public string LastItemName { get; set; }
            public List<ManagedProperties_Dictionary> ManagedPropertiesDictionary { get; set; }
            public int TotalCount { get; set; }
        }

        private class ManagedProperties_Dictionary
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public bool IsImplicit { get; set; }
        }

        private class Mappings
        {
            public string LastItemName { get; set; }
            public List<Mappings_Dictionary> MappingsDictionary { get; set; }
        }

        private class Mappings_Dictionary
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public string CrawledPropertyName { get; set; }
            public string CrawledPropset { get; set; }
            public string ManagedPid { get; set; }
            public string MappingOrder { get; set; }
            public string SchemaId { get; set; }
        }

        private class Overrides
        {
            public string LastItemName { get; set; }
            public List<Overrides_Dictionary> OverridesDictionary { get; set; }
        }

        private class Overrides_Dictionary
        {
            public string Key { get; set; }
            public string Name { get; set; }
        }

        private class SearchSchemaConfigSettings
        {
            public Aliases SeacrhScehmaAliases { get; set; }
            public List<CrawledPropertyInfoCollection> CategoriesAndCrawledProperties { get; set; }
            public ManagedProperties ManagedProperties { get; set; }
            public Mappings Mappings { get; set; }
            public Overrides Overrides { get; set; }
        }

        private class SearchQueryConfigurationSettings
        {
            public List<BestBets> BestBet { get; set; }
            public string DefaultSourceId { get; set; }
            public string DefaultSourceIdSet { get; set; }
            public string DeployToParent { get; set; }
            public string DisableInheritanceOnImport { get; set; }
            public List<QueryRule> QueryRules { get; set; }
            public List<ResultType> ResultTypes { get; set; }
            public List<Source> Sources { get; set; }
        }
        private class SearchRankingModelConfigurationSettings
        {
            public string RankingModel { get; set; }
        }
        private class SeacrhConfigSettings
        {
            public SearchQueryConfigurationSettings SearchQueryConfigurationSettings { get; set; }
            public SearchSchemaConfigSettings SearchSchemaConfigSettings { get; set; }
            public SearchRankingModelConfigurationSettings SearchRankingModelConfigurationSettings { get; set; }

        }
        #endregion
    }
}
