using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Diagnostics;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class for taxonomy extension methods
    /// </summary>
    [Guid("8A8AEA7A-7C25-4138-9C83-2584028868C5")]
    public static partial class TaxonomyExtensions
    {
        #region Taxonomy Management
        private static readonly Regex TrimSpacesRegex = new Regex("\\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex InvalidDescriptionRegex = new Regex("[\t]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex InvalidNameRegex = new Regex("[;\"<>|&\\t]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// The default Taxonomy Guid Label Delimiter
        /// </summary>
        public const string TaxonomyGuidLabelDelimiter = "|";

        /// <summary>
        /// Creates a new term group, in the specified term store.
        /// </summary>
        /// <param name="termStore">the term store to use</param>
        /// <param name="groupName">Name of the term group</param>
        /// <param name="groupId">(Optional) ID of the group; if not provided a random GUID is used</param>
        /// <param name="groupDescription">(Optional) Description of the term group</param>
        /// <returns>The created term group</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static TermGroup CreateTermGroup(this TermStore termStore, string groupName, Guid groupId = default(Guid), string groupDescription = null)
        {
            if (string.IsNullOrEmpty(groupName)) { throw new ArgumentNullException(nameof(groupName)); }

            TermGroup termGroup;
            groupName = NormalizeName(groupName);
            ValidateName(groupName, "groupName");

            // Create Group
            if (groupId == Guid.Empty)
            {
                groupId = Guid.NewGuid();
            }

            if (!termStore.IsObjectPropertyInstantiated("Name"))
            {
                // get instances to root web, since we are processing currently sub site 
                termStore.Context.Load(termStore);
                termStore.Context.ExecuteQueryRetry();
            }
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_CreateTermGroup0InStore1, groupName, termStore.Name);
            termGroup = termStore.CreateGroup(groupName, groupId);
            termStore.Context.Load(termGroup, g => g.Name, g => g.Id, g => g.Description);
            termStore.Context.ExecuteQueryRetry();

            // Apply description
            bool changed = false;
            if (groupDescription != null && !string.Equals(termGroup.Description, groupDescription))
            {
                try
                {
                    ValidateDescription(groupDescription, "groupDescription");
                    termGroup.Description = groupDescription;
                    changed = true;
                }
                catch (Exception ex)
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ExceptionUpdateDescriptionGroup01, termGroup.Name, termGroup.Id, ex.Message);
                }
            }
            if (changed)
            {
                Log.Debug(Constants.LOGGING_SOURCE, "Updating term group");
                termStore.Context.ExecuteQueryRetry();
                //termStore.CommitAll();
            }

            return termGroup;
        }

        /// <summary>
        /// Ensures the named group exists, returning a reference to the group, and creating or updating as necessary.
        /// </summary>
        /// <param name="site">Site connected to the term store to use</param>
        /// <param name="groupName">Name of the term group</param>
        /// <param name="groupId">(Optional) ID of the group; if not provided the parameter is ignored, a random GUID is used if necessary to create the group, otherwise if the ID differs a warning is logged</param>
        /// <param name="groupDescription">(Optional) Description of the term group; if null or not provided the parameter is ignored, otherwise the group is updated as necessary to match the description; passing an empty string will clear the description</param>
        /// <returns>The required term group</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static TermGroup EnsureTermGroup(this Site site, string groupName, Guid groupId = default(Guid), string groupDescription = null)
        {
            if (string.IsNullOrEmpty(groupName)) { throw new ArgumentNullException("groupName"); }

            TaxonomySession session = TaxonomySession.GetTaxonomySession(site.Context);
            var termStore = session.GetDefaultSiteCollectionTermStore();
            site.Context.Load(termStore, s => s.Name, s => s.Id);

            bool changed = false;
            TermGroup termGroup = null;
            groupName = NormalizeName(groupName);
            ValidateName(groupName, "groupName");

            // Find or create group
            IEnumerable<TermGroup> groups = site.Context.LoadQuery(termStore.Groups.Include(g => g.Name, g => g.Id, g => g.Description));
            site.Context.ExecuteQueryRetry();
            if (groupId != Guid.Empty)
            {
                termGroup = groups.FirstOrDefault(g => g.Id == groupId);
            }
            if (termGroup == null)
            {
                termGroup = groups.FirstOrDefault(g => string.Equals(g.Name, groupName, StringComparison.OrdinalIgnoreCase));
            }

            if (termGroup == null)
            {
                if (groupId == Guid.Empty)
                {
                    groupId = Guid.NewGuid();
                }
                Log.Info(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_CreateTermGroup0InStore1, groupName, termStore.Name);
                termGroup = termStore.CreateGroup(groupName, groupId);
                site.Context.Load(termGroup, g => g.Name, g => g.Id, g => g.Description);
                site.Context.ExecuteQueryRetry();
            }
            else
            {
                // Check ID (if retrieved by name and ID is different)
                if (groupId != Guid.Empty && termGroup.Id != groupId)
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_TermGroup0Id1DoesNotMatchSpecifiedId2, termGroup.Name, termGroup.Id, groupId);
                }
            }
            // Apply name (if retrieved by ID and name has changed)
            if (!string.Equals(termGroup.Name, groupName))
            {
                termGroup.Name = groupName;
                changed = true;
            }
            // Apply description
            if (groupDescription != null && !string.Equals(termGroup.Description, groupDescription))
            {
                try
                {
                    ValidateDescription(groupDescription, "groupDescription");
                    termGroup.Description = groupDescription;
                    changed = true;
                }
                catch (Exception ex)
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ExceptionUpdateDescriptionGroup01, termGroup.Name, termGroup.Id, ex.Message);
                }
            }
            if (changed)
            {
                Log.Debug(Constants.LOGGING_SOURCE, "Updating term group");
                site.Context.ExecuteQueryRetry();
                //termStore.CommitAll();
            }
            return termGroup;
        }

        /// <summary>
        /// Ensures the named term set exists, returning a reference to the set, and creating or updating as necessary.
        /// </summary>
        /// <param name="parentGroup">Group to check or create the term set in</param>
        /// <param name="termSetName">Name of the term set</param>
        /// <param name="termSetId">(Optional) ID of the term set; if not provided the parameter is ignored, a random GUID is used if necessary to create the term set, otherwise if the ID differs a warning is logged</param>
        /// <param name="lcid">(Optional) Default language of the term set; if not provided the default of the associate term store is used</param>
        /// <param name="description">(Optional) Description of the term set; if null or not provided the parameter is ignored, otherwise the term set is updated as necessary to match the description; passing an empty string will clear the description</param>
        /// <param name="isOpen">(Optional) Whether the term store is open for new term creation or not</param>
        /// <param name="termSetContact">(Optional) E-mail address for term suggestions and feedback</param>
        /// <param name="termSetOwner">Owner of termset</param>
        /// <returns>The required term set</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static TermSet EnsureTermSet(this TermGroup parentGroup, string termSetName, Guid termSetId = default(Guid), int? lcid = null, string description = null, bool? isOpen = null, string termSetContact = null, string termSetOwner = null)
        {
            bool changed = false;
            TermSet termSet = null;
            termSetName = NormalizeName(termSetName);
            ValidateName(termSetName, "termSetName");

            // Find or create term set
            parentGroup.Context.Load(parentGroup, g => g.Name, g => g.Id);
            IEnumerable<TermSet> termSets = parentGroup.Context.LoadQuery(parentGroup.TermSets.Include(g => g.Name, g => g.Id, g => g.Description, g => g.IsOpenForTermCreation, g => g.Contact, g => g.Owner));
            parentGroup.Context.ExecuteQueryRetry();
            if (termSetId != Guid.Empty)
            {
                termSet = termSets.FirstOrDefault(s => s.Id == termSetId);
            }
            if (termSet == null)
            {
                termSet = termSets.FirstOrDefault(s => string.Equals(s.Name, termSetName, StringComparison.OrdinalIgnoreCase));
            }

            if (termSet == null)
            {
                if (termSetId == Guid.Empty)
                {
                    termSetId = Guid.NewGuid();
                }
                if (lcid.HasValue)
                {
                    var termStore = parentGroup.TermStore;
                    parentGroup.Context.Load(termStore, ts => ts.Languages);
                    parentGroup.Context.ExecuteQueryRetry();
                    if (!termStore.Languages.Contains(lcid.Value))
                    {
                        termStore.AddLanguage(lcid.Value);
                    }
                }
                else
                {
                    var termStore = parentGroup.TermStore;
                    parentGroup.Context.Load(termStore, ts => ts.DefaultLanguage);
                    parentGroup.Context.ExecuteQueryRetry();
                    lcid = termStore.DefaultLanguage;
                }
                Log.Info(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_CreateTermSet0InGroup1, termSetName, parentGroup.Name);
                termSet = parentGroup.CreateTermSet(termSetName, termSetId, lcid.Value);
                parentGroup.Context.Load(termSet, g => g.Name, g => g.Id, g => g.Description, g => g.IsOpenForTermCreation, g => g.Contact, g => g.Owner);
                parentGroup.Context.ExecuteQueryRetry();
            }
            else
            {
                if (termSetId != Guid.Empty && termSet.Id != termSetId)
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_TermSet0Id1DoesNotMatchSpecifiedId2, termSet.Name, termSet.Id, termSetId);
                }
            }
            // Apply name (if retrieved by ID and name has changed)
            if (!string.Equals(termSet.Name, termSetName))
            {
                termSet.Name = termSetName;
                changed = true;
            }
            // Apply description
            if (description != null && (termSet.Description != description))
            {
                try
                {
                    ValidateDescription(description, "termSetDescription");
                    termSet.Description = description;
                    changed = true;
                }
                catch (Exception ex)
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ExceptionUpdateDescriptionSet01, termSet.Name, termSet.Id, ex.Message);
                }
            }
            // Other settings
            if (isOpen.HasValue && (termSet.IsOpenForTermCreation != isOpen.Value))
            {
                termSet.IsOpenForTermCreation = isOpen.Value;
                changed = true;
            }
            if (termSetContact != null && termSet.Contact != termSetContact)
            {
                termSet.Contact = termSetContact;
                changed = true;
            }
            if (termSetOwner != null && termSet.Owner != termSetOwner)
            {
                termSet.Owner = termSetOwner;
                changed = true;
            }

            // TODO: Add Stakeholders
            //if (settings.EnvironmentSettings.TermSetStakeholder) {
            //    foreach (user in settings.EnvironmentSettings.TermSetStakeholder) {
            //        Write-Host "Adding term set stakeholder 'user'."
            //        termSet.AddStakeholder(user)
            //    }
            //}

            // Update (if changed)
            if (changed)
            {
                //Diagnostics.TraceVerbose("Committing term set creation");
                Log.Debug(Constants.LOGGING_SOURCE, "Updating term set");
                parentGroup.Context.ExecuteQueryRetry();
            }
            return termSet;
        }

        /// <summary>
        /// Private method used for resolving taxonomy term set for taxonomy field
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <returns>Returns TermStore object</returns>
        private static TermStore GetDefaultTermStore(Web web)
        {
            TermStore termStore = null;
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(web.Context);
            web.Context.Load(taxonomySession,
                ts => ts.TermStores.Include(
                    store => store.Name,
                    store => store.Groups.Include(
                        group => group.Name
                        )
                    )
                );
            web.Context.ExecuteQueryRetry();
            if (taxonomySession != null)
            {
                termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            }

            return termStore;
        }

        /// <summary>
        /// Returns a new taxonomy session for the current site
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <returns>Returns TaxonomySession object</returns>
        public static TaxonomySession GetTaxonomySession(this Site site)
        {
            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(site.Context);
            site.Context.Load(tSession);
            site.Context.ExecuteQueryRetry();
            return tSession;
        }

        /// <summary>
        /// Returns the default keywords termstore for the current site
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <returns>Returns TermStore object</returns>
        public static TermStore GetDefaultKeywordsTermStore(this Site site)
        {
            TaxonomySession session = TaxonomySession.GetTaxonomySession(site.Context);
            var termStore = session.GetDefaultKeywordsTermStore();
            site.Context.Load(termStore);
            site.Context.ExecuteQueryRetry();

            return termStore;
        }

        /// <summary>
        /// Returns the default site collection termstore
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <returns>Returns TermStore object</returns>
        public static TermStore GetDefaultSiteCollectionTermStore(this Site site)
        {
            TaxonomySession session = TaxonomySession.GetTaxonomySession(site.Context);
            var termStore = session.GetDefaultSiteCollectionTermStore();
            site.Context.Load(termStore);
            site.Context.ExecuteQueryRetry();

            return termStore;
        }

        /// <summary>
        /// Gets the URL of the content type syndication hub, if it exists.
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <returns>Returns the URL of the content type syndication hub</returns>
        public static string GetContentTypePublishingHub(this Site site)
        {
            TaxonomySession session = TaxonomySession.GetTaxonomySession(site.Context);
            var termStore = session.GetDefaultSiteCollectionTermStore();
            site.Context.Load(termStore, s => s.ContentTypePublishingHub);
            site.Context.ExecuteQueryRetry();

            return termStore.ContentTypePublishingHub;
        }

        /// <summary>
        /// Finds a termset by name
        /// </summary>
        /// <param name="site">The current site</param>
        /// <param name="name">The name of the termset</param>
        /// <param name="lcid">The locale ID for the termset to return, defaults to 1033</param>
        /// <returns>Returns collection of TermSet</returns>
        public static TermSetCollection GetTermSetsByName(this Site site, string name, int lcid = 1033)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException("name");

            TaxonomySession session = TaxonomySession.GetTaxonomySession(site.Context);
            TermStore store = session.GetDefaultSiteCollectionTermStore();
            var termsets = store.GetTermSetsByName(name, lcid);
            site.Context.Load(termsets);
            site.Context.ExecuteQueryRetry();
            return termsets;
        }


        /// <summary>
        /// Finds a termgroup by name
        /// </summary>
        /// <param name="site">The current site</param>
        /// <param name="name">The name of the termgroup</param>
        /// <returns>Returns TermGroup object</returns>
        public static TermGroup GetTermGroupByName(this Site site, string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException("name");

            TaxonomySession session = TaxonomySession.GetTaxonomySession(site.Context);
            var store = session.GetDefaultSiteCollectionTermStore();
            IEnumerable<TermGroup> groups = site.Context.LoadQuery(store.Groups.Include(g => g.Name, g => g.Id, g => g.TermSets)).Where(g => g.Name == name);
            site.Context.ExecuteQueryRetry();
            return groups.FirstOrDefault();
        }

        /// <summary>
        /// Gets the named term group, if it exists in the term store.
        /// </summary>
        /// <param name="termStore">The term store to use</param>
        /// <param name="groupName">Name of the term group</param>
        /// <returns>The requested term group, or null if it does not exist</returns>
        public static TermGroup GetTermGroupByName(this TermStore termStore, string groupName)
        {
            if (string.IsNullOrEmpty(groupName)) { throw new ArgumentNullException(nameof(groupName)); }

            TermGroup termGroup;
            groupName = NormalizeName(groupName);
            ValidateName(groupName, "groupName");

            // Find group
            var groups = termStore.Context.LoadQuery(termStore.Groups.Include(g => g.Name, g => g.Id, g => g.Description));
            termStore.Context.ExecuteQueryRetry();
            termGroup = groups.FirstOrDefault(g => string.Equals(g.Name, groupName, StringComparison.OrdinalIgnoreCase));
            return termGroup;
        }

        /// <summary>
        /// Finds a termgroup by its ID
        /// </summary>
        /// <param name="site">The current site</param>
        /// <param name="termGroupId">The ID of the termgroup</param>
        /// <returns>Returns TermGroup object</returns>
        public static TermGroup GetTermGroupById(this Site site, Guid termGroupId)
        {
            if (termGroupId == null || termGroupId.Equals(Guid.Empty))
            {
                throw new ArgumentNullException("termGroupId");
            }

            TaxonomySession session = TaxonomySession.GetTaxonomySession(site.Context);
            var store = session.GetDefaultSiteCollectionTermStore();
            IEnumerable<TermGroup> groups = site.Context.LoadQuery(store.Groups.Include(g => g.Name, g => g.Id, g => g.TermSets)).Where(g => g.Id == termGroupId);
            site.Context.ExecuteQueryRetry();
            return groups.FirstOrDefault();
        }

        /// <summary>
        /// Gets a Taxonomy Term by Name
        /// </summary>
        /// <param name="site">The site to process</param>
        /// <param name="termSetId">Guid of a TermSet</param>
        /// <param name="term">Term name</param>
        /// <returns>Returns Term object</returns>
        public static Term GetTermByName(this Site site, Guid termSetId, string term)
        {
            if (string.IsNullOrEmpty(term))
                throw new ArgumentNullException(nameof(term));

            TermCollection termMatches = null;
            ExceptionHandlingScope scope = new ExceptionHandlingScope(site.Context);

            string termId = string.Empty;
            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(site.Context);
            TermStore ts = tSession.GetDefaultSiteCollectionTermStore();
            TermSet tset = ts.GetTermSet(termSetId);

            var lmi = new LabelMatchInformation(site.Context);

            lmi.Lcid = ts.EnsureProperty(tstore => tstore.DefaultLanguage);
            lmi.TrimUnavailable = true;
            lmi.TermLabel = term;

            termMatches = tset.GetTerms(lmi);
            site.Context.Load(tSession);
            site.Context.Load(ts);
            site.Context.Load(tset);
            site.Context.Load(termMatches);

            site.Context.ExecuteQueryRetry();

            if (termMatches.AreItemsAvailable)
            {
                return termMatches.FirstOrDefault();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Adds a term to a given termset
        /// </summary>
        /// <param name="site">The current site</param>
        /// <param name="termSetId">The ID of the termset</param>
        /// <param name="term">The label of the new term to create</param>
        /// <returns>Returns Term object</returns>
        public static Term AddTermToTermset(this Site site, Guid termSetId, string term)
        {
            return AddTermToTermset(site, termSetId, term, Guid.NewGuid());
        }

        /// <summary>
        /// Adds a term to a given termset
        /// </summary>
        /// <param name="site">The current site</param>
        /// <param name="termSetId">The ID of the termset</param>
        /// <param name="term">The label of the new term to create</param>
        /// <param name="termId">The ID of the term to create</param>
        /// <returns>Returns Term object</returns>
        public static Term AddTermToTermset(this Site site, Guid termSetId, string term, Guid termId)
        {
            if (string.IsNullOrEmpty(term))
                throw new ArgumentNullException(nameof(term));

            Term t = null;
            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(site.Context);
            TermStore ts = tSession.GetDefaultSiteCollectionTermStore();
            TermSet tset = ts.GetTermSet(termSetId);

            t = tset.CreateTerm(term, ts.EnsureProperty(tstore => tstore.DefaultLanguage), termId);
            //site.Context.Load(tSession);
            //site.Context.Load(ts);
            //site.Context.Load(tset);
            site.Context.Load(t);

            site.Context.ExecuteQueryRetry();

            return t;
        }

        /// <summary>
        ///  Imports an array of | delimited strings into the deafult site collection termstore. Specify strings in this format:
        ///  TermGroup|TermSet|Term
        ///  
        ///  E.g. "Locations|Nordics|Sweden"
        ///  
        /// </summary>
        /// <param name="site">The current site</param>
        /// <param name="termLines">Array of TermLines</param>
        /// <param name="lcid">Locale identifier (LCID) for the language</param>
        /// <param name="delimiter">delimeter which seperates terms</param>
        /// <param name="synchronizeDeletions">Remove tags that are not present in the import</param>
        public static void ImportTerms(this Site site, string[] termLines, int lcid, string delimiter = "|", bool synchronizeDeletions = false)
        {
            termLines.ValidateNotNullOrEmpty("termLines");

            var clientContext = site.Context;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();

            ImportTerms(site, termLines, lcid, termStore, delimiter, synchronizeDeletions);
        }

        /// <summary>
        ///  Imports an array of | delimited strings into the deafult site collection termstore. Specify strings in this format:
        ///  TermGroup|TermSet|Term
        ///  
        ///  E.g. "Locations|Nordics|Sweden"
        ///  
        /// </summary>
        /// <param name="site">The current site</param>
        /// <param name="termLines">Array of TermLines</param>
        /// <param name="lcid">Locale identifier (LCID) for the language</param>
        /// <param name="termStore">The termstore to import the terms into</param>
        /// <param name="delimiter">delimeter which seperates terms</param>
        /// <param name="synchronizeDeletions">Remove tags that are not present in the import</param>
        public static void ImportTerms(this Site site, string[] termLines, int lcid, TermStore termStore, string delimiter = "|", bool synchronizeDeletions = false)
        {
            var groupDict = new Dictionary<TermGroup, Dictionary<string, List<string>>>();

            var clientContext = site.Context;
            if (termStore.ServerObjectIsNull == true)
            {
                clientContext.Load(termStore);
                clientContext.ExecuteQueryRetry();
            }
            clientContext.Load(termStore);
            clientContext.ExecuteQueryRetry();

            foreach (var line in termLines)
            {
                // Find termgroup
                var items = line.Split(new[] { delimiter }, StringSplitOptions.None);
                if (items.Any())
                {
                    Dictionary<string, List<string>> termsets = null;

                    var groupItem = items[0];
                    var groupName = groupItem;
                    var groupId = Guid.Empty;
                    if (groupItem.IndexOf(";#", StringComparison.Ordinal) > -1)
                    {
                        groupName = groupItem.Split(new[] { ";#" }, StringSplitOptions.None)[0];
                        groupId = new Guid(groupItem.Split(new[] { ";#" }, StringSplitOptions.None)[1]);
                    }
                    TermGroup termGroup = null;
                    // Cached?
                    if (groupDict.Any())
                    {
                        KeyValuePair<TermGroup, Dictionary<string, List<string>>> groupDictItem;
                        if (groupId != Guid.Empty)
                        {
                            groupDictItem = groupDict.FirstOrDefault(tg => tg.Key.Id == groupId);

                            termGroup = groupDictItem.Key;
                            termsets = groupDictItem.Value;
                        }
                        else
                        {
                            groupDictItem = groupDict.FirstOrDefault(tg => tg.Key.Name == groupName);

                            termGroup = groupDictItem.Key;
                            termsets = groupDictItem.Value;
                        }
                    }
                    if (termGroup == null)
                    {
                        if (groupId != Guid.Empty)
                        {

                            termGroup = termStore.Groups.GetById(groupId);
                        }
                        else
                        {
                            termGroup = termStore.Groups.GetByName(NormalizeName(groupName));
                        }
                        try
                        {
                            clientContext.Load(termGroup);
                            clientContext.ExecuteQueryRetry();
                            groupDict.Add(termGroup, new Dictionary<string, List<string>>());
                            termsets = new Dictionary<string, List<string>>();
                        }
                        catch
                        {

                        }
                    }
                    if (termGroup.ServerObjectIsNull == null)
                    {
                        if (groupId == Guid.Empty)
                        {
                            groupId = Guid.NewGuid();
                        }
                        termGroup = termStore.CreateGroup(NormalizeName(groupName), groupId);
                        termsets = new Dictionary<string, List<string>>();
                        clientContext.Load(termGroup);
                        clientContext.ExecuteQueryRetry();

                        groupDict.Add(termGroup, new Dictionary<string, List<string>>());

                    }
                    var sb = new StringBuilder();
                    if (items.Length > 1)
                    {
                        var termSetName = items[1];
                        termSetName = termSetName.Replace(";#", "|");
                        sb.AppendFormat("{0},,{1},True,,", termSetName, lcid);

                        // Termset = position 1
                        for (var q = 0; q < 7; q++)
                        {
                            var item = "";
                            if (items.Length > q + 2)
                            {
                                item = items[q + 2];
                                item = item.Replace(";#", "|");
                            }
                            sb.AppendFormat("{0},", NormalizeName(item));
                        }
                        if (termsets != null)
                        {
                            if (termsets.ContainsKey(termSetName)) {
                                termsets[termSetName].Add(sb.ToString());
                            }
                            else
                            {
                                termsets.Add(termSetName, new List<string>() { sb.ToString() });
                            }

                            groupDict[termGroup] = termsets;
                        }
                    }
                }
            }
            foreach (var groupDictItem in groupDict)
            {
                var termGroup = groupDictItem.Key as TermGroup;
                foreach (var termset in groupDictItem.Value)
                {
                    using (var memoryStream = new MemoryStream())
                    using (var streamWriter = new StreamWriter(memoryStream))
                    {
                        // Header
                        streamWriter.WriteLine(@"""Term Set Name"",""Term Set Description"",""LCID"",""Available for Tagging"",""Term Description"",""Level 1 Term"",""Level 2 Term"",""Level 3 Term"",""Level 4 Term"",""Level 5 Term"",""Level 6 Term"",""Level 7 Term""");

                        // Items
                        foreach (var termLine in termset.Value)
                        {
                            streamWriter.WriteLine(termLine);
                        }
                        streamWriter.Flush();
                        memoryStream.Position = 0;
                        termGroup.ImportTermSet(memoryStream, synchroniseDeletions: synchronizeDeletions);
                    }
                }
            }
        }

        private static Term AddTermToTerm(this Term term, int lcid, string termLabel, Guid termId)
        {
            var clientContext = term.Context;
            if (term.ServerObjectIsNull == true)
            {
                clientContext.Load(term);
                clientContext.ExecuteQueryRetry();
            }
            Term subTerm = null;
            if (termId != Guid.Empty)
            {
                subTerm = term.Terms.GetById(termId);
            }
            else
            {
                subTerm = term.Terms.GetByName(NormalizeName(termLabel));
            }
            clientContext.Load(term);
            try
            {
                clientContext.ExecuteQueryRetry();
            }
            catch { }

            clientContext.Load(subTerm);
            try
            {
                clientContext.ExecuteQueryRetry();
            }
            catch { }
            if (subTerm.ServerObjectIsNull == null)
            {
                if (termId == Guid.Empty) termId = Guid.NewGuid();
                subTerm = term.CreateTerm(NormalizeName(termLabel), lcid, termId);
                clientContext.Load(subTerm);
                clientContext.ExecuteQueryRetry();
            }
            return subTerm;
        }

        /// <summary>
        /// Imports terms from a term set file, updating with any new terms, in the same format at that used by the web interface import ability.
        /// </summary>
        /// <param name="termGroup">Group to create the term set within</param>
        /// <param name="filePath">Local path to the file to import</param>
        /// <param name="termSetId">GUID to use for the term set; if Guid.Empty is passed then a random GUID is generated and used</param>
        /// <param name="synchroniseDeletions">(Optional) Whether to also synchronise deletions; that is, remove any terms not in the import file; default is no (false)</param>
        /// <param name="termSetIsOpen">(Optional) Whether the term set should be marked open; if not passed, then the existing setting is not changed</param>
        /// <param name="termSetContact">(Optional) Contact for the term set; if not provided, the existing setting is retained</param>
        /// <param name="termSetOwner">(Optional) Owner for the term set; if not provided, the existing setting is retained</param>
        /// <returns>The created, or updated, term set</returns>
        /// <remarks>
        /// <para>
        /// The format of the file is the same as that used by the import function in the 
        /// web interface. A sample file can be obtained from the web interface.
        /// </para>
        /// <para>
        /// This is a CSV file, with the following headings:
        /// </para>
        /// <para>
        /// <code>Term Set Name,Term Set Description,LCID,Available for Tagging,Term Description,Level 1 Term,Level 2 Term,Level 3 Term,Level 4 Term,Level 5 Term,Level 6 Term,Level 7 Term</code>
        /// </para>
        /// <para>
        /// The first data row must contain the Term Set Name, Term Set Description, and LCID, and should also contain the first term. 
        /// </para>
        /// <para>
        /// It is recommended that a fixed GUID be used as the termSetId, to allow the term set to be easily updated (so do not pass Guid.Empty).
        /// </para>
        /// <para>
        /// In contrast to the web interface import, this is not a one-off import but runs synchronisation logic allowing updating of an existing Term Set.
        /// When synchronising, any existing terms are matched (with Term Description and Available for Tagging updated as necessary),
        /// any new terms are added in the correct place in the hierarchy, and (if synchroniseDeletions is set) any terms not in the imported file 
        /// are removed.
        /// </para>
        /// <para>
        /// The import file also supports an expanded syntax for the Term Set Name and term names (Level 1 Term, Level 2 Term, etc).
        /// These columns support values with the format "Name|GUID", with the name and GUID separated by a pipe character (note that the pipe character is invalid to use within a taxomony item name).
        /// This expanded syntax is not required, but can be used to ensure all terms have fixed IDs.
        /// </para>
        /// </remarks>
        public static TermSet ImportTermSet(this TermGroup termGroup, string filePath, Guid termSetId = default(Guid), bool synchroniseDeletions = false, bool? termSetIsOpen = null, string termSetContact = null, string termSetOwner = null)
        {
            if (filePath == null) { throw new ArgumentNullException("filePath"); }
            if (string.IsNullOrWhiteSpace(filePath)) { throw new ArgumentException(CoreResources.TaxonomyExtensions_ImportTermSet_File_path_is_required_, "filePath"); }

            using (var fs = new FileStream(filePath, FileMode.Open))
            {
                return ImportTermSet(termGroup, fs, termSetId, synchroniseDeletions, termSetIsOpen, termSetContact, termSetOwner);
            }
        }

        /// <summary>
        /// Imports terms from a term set stream, updating with any new terms, in the same format at that used by the web interface import ability.
        /// </summary>
        /// <param name="termGroup">Group to create the term set within</param>
        /// <param name="termSetData">Stream containing the data to import</param>
        /// <param name="termSetId">GUID to use for the term set; if Guid.Empty is passed then a random GUID is generated and used</param>
        /// <param name="synchroniseDeletions">(Optional) Whether to also synchronise deletions; that is, remove any terms not in the import file; default is no (false)</param>
        /// <param name="termSetIsOpen">(Optional) Whether the term set should be marked open; if not passed, then the existing setting is not changed</param>
        /// <param name="termSetContact">(Optional) Contact for the term set; if not provided, the existing setting is retained</param>
        /// <param name="termSetOwner">(Optional) Owner for the term set; if not provided, the existing setting is retained</param>
        /// <returns>The created, or updated, term set</returns>
        /// <remarks>
        /// <para>
        /// The format of the file is the same as that used by the import function in the 
        /// web interface. A sample file can be obtained from the web interface.
        /// </para>
        /// <para>
        /// This is a CSV file, with the following headings:
        /// </para>
        /// <para>
        /// <code>Term Set Name,Term Set Description,LCID,Available for Tagging,Term Description,Level 1 Term,Level 2 Term,Level 3 Term,Level 4 Term,Level 5 Term,Level 6 Term,Level 7 Term</code>
        /// </para>
        /// <para>
        /// The first data row must contain the Term Set Name, Term Set Description, and LCID, and should also contain the first term. 
        /// </para>
        /// <para>
        /// It is recommended that a fixed GUID be used as the termSetId, to allow the term set to be easily updated (so do not pass Guid.Empty).
        /// </para>
        /// <para>
        /// In contrast to the web interface import, this is not a one-off import but runs synchronisation logic allowing updating of an existing Term Set.
        /// When synchronising, any existing terms are matched (with Term Description and Available for Tagging updated as necessary),
        /// any new terms are added in the correct place in the hierarchy, and (if synchroniseDeletions is set) any terms not in the imported file 
        /// are removed.
        /// </para>
        /// <para>
        /// The import file also supports an expanded syntax for the Term Set Name and term names (Level 1 Term, Level 2 Term, etc).
        /// These columns support values with the format "Name|GUID", with the name and GUID separated by a pipe character (note that the pipe character is invalid to use within a taxomony item name).
        /// This expanded syntax is not required, but can be used to ensure all terms have fixed IDs.
        /// </para>
        /// </remarks>
        public static TermSet ImportTermSet(this TermGroup termGroup, Stream termSetData, Guid termSetId = default(Guid), bool synchroniseDeletions = false, bool? termSetIsOpen = null, string termSetContact = null, string termSetOwner = null)
        {
            if (termSetData == null) { throw new ArgumentNullException("termSetData"); }

            Log.Info(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ImportTermSet);

            TermSet termSet = null;
            var importedTermIds = new Dictionary<Guid, object>();
            using (var reader = new StreamReader(termSetData))
            {
                bool allTermsAdded;
                termSet = ImportTermSetImplementation(termGroup, reader, termSetId, importedTermIds, termSetIsOpen, termSetContact, termSetOwner, out allTermsAdded);
            }

            if (synchroniseDeletions)
            {
                ImportTermSetRemoveExtraTerms(termSet, importedTermIds);
            }

            return termSet;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        private static TermSet ImportTermSetImplementation(this TermGroup parentGroup, TextReader reader, Guid termSetId, IDictionary<Guid, object> importedTermIds, bool? termSetIsOpen, string termSetContact, string termSetOwner, out bool allTermsAdded)
        {
            if (parentGroup == null)
            {
                throw new ArgumentNullException("parentGroup");
            }
            if (reader == null)
            {
                throw new ArgumentNullException("reader");
            }

            Log.Debug(Constants.LOGGING_SOURCE, "Begin import term set");

            TermSet termSet = null;

            int lcid = 0;

            int lineIndex = -1;
            allTermsAdded = true;
            checked
            {
                try
                {
                    string rowText;
                    while ((rowText = reader.ReadLine()) != null)
                    {
                        lineIndex++;
                        if (lineIndex == 0)
                        {
                            // Check file look vaguely like a CSV -- ensure the first line (headers) has some commas:
                            if (!rowText.Contains(","))
                            {
                                throw new ArgumentException(CoreResources.TaxonomyExtensions_ImportTermSetImplementation_Invalid_CSV_format__was_expecting_a_comma_in_the_first__header__line_, "reader");
                            }
                        }
                        else
                        {
                            // Process the second line (index=1), and then all non-blank lines
                            if (lineIndex <= 1 || !string.IsNullOrEmpty(rowText.Trim()))
                            {
                                var entries = ImportTermSetLineParse(rowText);
                                //lcid = this.GetImportLcid(termStore, lcid, lineIndex, entries);
                                if (termSet == null)
                                {
                                    if (lineIndex != 1)
                                    {
                                        throw new InvalidOperationException("Term set not created on first line.");
                                    }
                                    if (entries.Count > 0)
                                    {
                                        string termSetName = entries[0];
                                        // Accept extended format of "Name|Guid", noting that | is not an allowed character in the term name
                                        if (termSetName.Contains(TaxonomyGuidLabelDelimiter))
                                        {
                                            var split = termSetName.Split(new string[] { TaxonomyGuidLabelDelimiter }, StringSplitOptions.None);
                                            termSetName = split[0];
                                            termSetId = new Guid(split[1]);
                                        }
                                        string description = null;
                                        if (entries.Count > 1)
                                        {
                                            description = entries[1];
                                        }
                                        if (entries.Count > 2)
                                        {
                                            if (!Int32.TryParse(entries[2], NumberStyles.Integer | NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite,
                                                NumberFormatInfo.InvariantInfo, out lcid))
                                            {
                                                var termStore = parentGroup.TermStore;
                                                parentGroup.Context.Load(termStore, ts => ts.DefaultLanguage);
                                                parentGroup.Context.ExecuteQueryRetry();
                                                lcid = termStore.DefaultLanguage;
                                            }
                                        }
                                        termSet = parentGroup.EnsureTermSet(termSetName, termSetId, lcid, description, termSetIsOpen, termSetContact, termSetOwner);
                                        //termStore.CommitAll();
                                    }
                                }
                                var termAdded = ImportTermSetLineImport(entries, termSet, lcid, lineIndex + 1, importedTermIds);
                                if (!termAdded)
                                {
                                    allTermsAdded = false;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(
                        $"Exception on line {lineIndex + 1}: {ex.Message}",
                        ex);
                }
                Log.Debug(Constants.LOGGING_SOURCE, "End ImportTermSet");
                return termSet;
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        private static bool ImportTermSetLineImport(IList<string> entries, TermSet importTermSet, int lcid, int lineNumber, IDictionary<Guid, object> importedTermIds)
        {
            TermSetItem parentTermSetItem = null;
            Term term = null;
            int num = 0;
            bool success = true;
            bool result = false;
            bool termCreated = false;
            bool changed = false;
            if (entries == null || entries.Count <= 5)
            {
                return false;
            }
            num = 0;
            checked
            {
                string termName = null;
                Guid termId = Guid.Empty;
                // Find matching existing terms
                while (num < entries.Count - 5 && success)
                {
                    string termNameEntry = entries[5 + num];
                    if (string.IsNullOrEmpty(termNameEntry))
                    {
                        if (termCreated)
                        {
                            result = true;
                        }
                        break;
                    }
                    termName = null;
                    termId = Guid.Empty;
                    // Accept extended format of "Name|Guid", noting that | is not an allowed character in the term name
                    if (termNameEntry.Contains(TaxonomyGuidLabelDelimiter))
                    {
                        var split = termNameEntry.Split(new string[] { TaxonomyGuidLabelDelimiter }, StringSplitOptions.None);
                        termName = split[0];
                        termId = new Guid(split[1]);
                    }
                    else
                    {
                        termName = termNameEntry;
                    }
                    // Process the entry
                    if (termName.Length > 255)
                    {
                        termName = termName.Substring(0, 255);
                    }
                    termName = NormalizeName(termName);
                    try
                    {
                        ValidateName(termName, "name");
                    }
                    catch (ArgumentNullException)
                    {
                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ImportErrorName0Line1, new object[]
                        {
                            termName,
                            lineNumber
                        });
                        success = false;
                        break;
                    }
                    catch (ArgumentException)
                    {
                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ImportErrorName0Line1, new object[]
                        {
                            termName,
                            lineNumber
                        });
                        success = false;
                        break;
                    }
                    if (term == null)
                    {
                        parentTermSetItem = importTermSet;
                    }
                    else
                    {
                        parentTermSetItem = term;
                    }
                    term = null;
                    if (!parentTermSetItem.IsObjectPropertyInstantiated("Terms"))
                    {
                        parentTermSetItem.Context.Load(parentTermSetItem, i => i.Terms.Include(t => t.Id, t => t.Name, t => t.Description, t => t.IsAvailableForTagging));
                        parentTermSetItem.Context.ExecuteQueryRetry();
                    }
                    foreach (Term current in parentTermSetItem.Terms)
                    {
                        if (termId != Guid.Empty && current.Id == termId)
                        {
                            term = current;
                            break;
                        }
                        if (current.Name == termName)
                        {
                            term = current;
                            break;
                        }
                    }
                    if (term == null && parentTermSetItem != null)
                    {
                        if (termId == Guid.Empty)
                        {
                            termId = Guid.NewGuid();
                        }
                        Log.Info(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_CreateTerm01UnderParent2, termName, termId, parentTermSetItem.Name);
                        term = parentTermSetItem.CreateTerm(termName, lcid, termId);
                        parentTermSetItem.Context.Load(parentTermSetItem, i => i.Terms.Include(t => t.Id, t => t.Name, t => t.Description, t => t.IsAvailableForTagging));
                        parentTermSetItem.Context.Load(term, t => t.Id, t => t.Name, t => t.Description, t => t.IsAvailableForTagging);
                        parentTermSetItem.Context.ExecuteQueryRetry();
                        termCreated = true;
                        if (num == entries.Count - 5 - 1)
                        {
                            result = true;
                        }
                    }
                    if (term != null)
                    {
                        importedTermIds[term.Id] = null;
                    }
                    num++;
                }
                if (success && term != null)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(entries[3]))
                        {
                            var isAvailableForTagging = bool.Parse(entries[3]);
                            if (term.IsAvailableForTagging != isAvailableForTagging)
                            {
                                Log.Debug(Constants.LOGGING_SOURCE, "Setting IsAvailableForTagging = {1} for term '{0}'.", term.Name, isAvailableForTagging);
                                term.IsAvailableForTagging = isAvailableForTagging;
                                changed = true;
                            }
                        }
                        else
                        {
                            Log.Debug(Constants.LOGGING_SOURCE, "The available for tagging entry on line {0} is null or empty.", new object[]
                            {
                                lineNumber
                            });
                        }
                    }
                    catch (ArgumentNullException)
                    {
                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ImportErrorTaggingLine0, new object[]
                        {
                            lineNumber
                        });
                        success = false;
                    }
                    catch (FormatException)
                    {
                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ImportErrorTaggingLine0, new object[]
                        {
                            lineNumber
                        });
                        success = false;
                    }
                    string description = entries[4];
                    if (description.Length > 1000)
                    {
                        description = description.Substring(0, 1000);
                    }
                    if (!string.IsNullOrEmpty(description))
                    {
                        try
                        {
                            ValidateDescription(description, "description");
                            if (term.Description != description)
                            {
                                Log.Debug(Constants.LOGGING_SOURCE, "Updating description for term '{0}'.", term.Name);
                                term.SetDescription(description, lcid);
                                changed = true;
                            }
                        }
                        catch (ArgumentException)
                        {
                            Log.Error(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ImportErrorDescription0Line1, new object[]
                            {
                                description,
                                lineNumber
                            });
                            success = false;
                        }
                    }
                    if (term.Name != termName)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, "Updating name for term '{0}'.", term.Name);
                        term.Name = termName;
                        changed = true;
                    }
                    if (!success)
                    {
                        result = false;
                        Guid id = term.Id;
                        try
                        {
                            Log.Debug(Constants.LOGGING_SOURCE, "Was an issue; deleting");
                            term.DeleteObject();
                            changed = true;
                        }
                        catch (Exception ex)
                        {
                            Log.Error(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_ImportErrorDeleteId0Line1, new object[]
                            {
                                id,
                                lineNumber
                            }, ex.Message);
                        }
                    }
                    if (changed)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, "Updating term {0}", term.Id);
                        parentTermSetItem.Context.ExecuteQueryRetry();
                    }
                }
                return result || changed;
            }
        }

        private static IList<string> ImportTermSetLineParse(string line)
        {
            List<string> entries = new List<string>();
            char[] lineChars = line.ToCharArray();
            string entry = string.Empty;
            bool flagInsideQuotes = false;
            int charIndex = 0;
            checked
            {
                while (charIndex < line.Length)
                {
                    if (flagInsideQuotes || !string.IsNullOrEmpty(entry)
                        || (!char.IsWhiteSpace(lineChars[charIndex]) && lineChars[charIndex] != '"'))
                    {
                        if (flagInsideQuotes && lineChars[charIndex] == '"'
                            && (charIndex + 1 >= line.Length || lineChars[charIndex + 1] == ','))
                        {
                            // End of quotes (and either end of line or next char is comma)
                            flagInsideQuotes = false;
                        }
                        else
                        {

                            if (flagInsideQuotes && lineChars[charIndex] == '"')
                            {
                                if (lineChars[charIndex + 1] != '"')
                                {
                                    // End of quotes  and next char is not a comma!
                                    return null;
                                }
                                // Doubled (escaped) quotes
                                charIndex++;
                            }
                            if (flagInsideQuotes || lineChars[charIndex] != ',')
                            {
                                entry += lineChars[charIndex];
                            }
                            else
                            {
                                entry = entry.Trim();
                                entries.Add(entry);
                                entry = string.Empty;
                            }
                        }
                    }
                    else
                    {
                        if (lineChars[charIndex] == '"')
                        {
                            flagInsideQuotes = true;
                        }
                    }

                    charIndex++;
                }
                entry = entry.Trim();
                entries.Add(entry);

                return entries;
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        private static void ImportTermSetRemoveExtraTerms(TermSet termSet, IDictionary<Guid, object> importedTermIds)
        {
            Log.Debug(Constants.LOGGING_SOURCE, "Removing extra terms");
            var termsToDelete = new List<Term>();
            var allTerms = termSet.GetAllTerms();
            termSet.Context.Load(allTerms, at => at.Include(t => t.Id, t => t.Name));
            termSet.Context.ExecuteQueryRetry();
            foreach (var term in allTerms)
            {
                if (!importedTermIds.ContainsKey(term.Id))
                {
                    termsToDelete.Add(term);
                }
            }
            foreach (var termToDelete in termsToDelete)
            {
                try
                {
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.TaxonomyExtension_DeleteTerm01, termToDelete.Name, termToDelete.Id);
                    termToDelete.DeleteObject();
                    termSet.Context.ExecuteQueryRetry();
                }
                catch (ServerException ex)
                {
                    if (ex.Message.StartsWith("Taxonomy item instantiation failed."))
                    {
                        // This is a sucky way to check if the term was already deleted
                        Log.Debug(Constants.LOGGING_SOURCE, "Term id {0} already deleted.", termToDelete.Id);
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// Exports the full list of terms from all termsets in all termstores.
        /// </summary>
        /// <param name="site">The site to process</param>
        /// <param name="termSetId">The ID of the termset to export</param>
        /// <param name="includeId">if true, Ids of the the taxonomy items will be included</param>
        /// <param name="delimiter">if specified, this delimiter will be used. Notice that IDs will be delimited with ;# from the label</param>
        /// <param name="lcid">if specified, retrieve terms in the specificed language</param>
        /// <returns>Returns list of Termset strings</returns>
        public static List<string> ExportTermSet(this Site site, Guid termSetId, bool includeId, string delimiter = "|", int lcid = 0)
        {
            var termStore = site.GetDefaultSiteCollectionTermStore();

            return ExportTermSet(site, termSetId, includeId, termStore, delimiter, lcid);
        }

        /// <summary>
        /// Exports the full list of terms from all termsets in all termstores.
        /// </summary>
        /// <param name="site">The site to export the termsets from</param>
        /// <param name="termSetId">The ID of the termset to export</param>
        /// <param name="includeId">if true, Ids of the the taxonomy items will be included</param>
        /// <param name="termStore">The term store to export the termset from</param>
        /// <param name="delimiter">if specified, this delimiter will be used. Notice that IDs will be delimited with ;# from the label</param>
        /// <param name="lcid">if specified, retrieve terms in the specificed language</param>
        /// <returns>Returns list of Termset strings</returns>
        public static List<string> ExportTermSet(this Site site, Guid termSetId, bool includeId, TermStore termStore, string delimiter = "|", int lcid = 0)
        {
            var clientContext = site.Context;
            var termsString = new List<string>();
            TermCollection terms = null;
            TermSet termSet = null;

            if (termSetId != Guid.Empty)
            {
                termSet = termStore.GetTermSet(termSetId);
                terms = termSet.Terms;
                if (lcid != 0)
                {
                    clientContext.Load(terms, t => t.IncludeWithDefaultProperties(s => s.TermSet), t => t.IncludeWithDefaultProperties(s => s.TermSet.Group), t => t.IncludeWithDefaultProperties(s => s.Labels));
                    clientContext.Load(termSet, ts => ts.Names);
                }
                else
                {
                    clientContext.Load(terms, t => t.IncludeWithDefaultProperties(s => s.TermSet), t => t.IncludeWithDefaultProperties(s => s.TermSet.Group));
                }
            }

            clientContext.ExecuteQueryRetry();

            if (terms.Any())
            {
                foreach (var term in terms)
                {
                    var groupName = DenormalizeName(term.TermSet.Group.Name);
                    var termsetName = DenormalizeName(term.TermSet.Name);
                    var termName = DenormalizeName(term.Name);
                    if (lcid != 0)
                    {
                        var termSetLabel = termSet.Names.SingleOrDefault(n => n.Key == lcid + "");
                        if (!string.IsNullOrWhiteSpace(termSetLabel.Value))
                        {
                            termsetName = DenormalizeName(termSetLabel.Value);
                        }

                        var label = term.Labels.SingleOrDefault(l => l.Language == lcid);
                        if (label != null && !string.IsNullOrWhiteSpace(label.Value))
                        {
                            termName = DenormalizeName(label.Value);
                        }
                    }

                    var groupPath = string.Format("{0}{1}", groupName, (includeId) ? string.Format(";#{0}", term.TermSet.Group.Id.ToString()) : "");
                    var termsetPath = string.Format("{0}{1}", termsetName, (includeId) ? string.Format(";#{0}", term.TermSet.Id.ToString()) : "");
                    var termPath = string.Format("{0}{1}", termName, (includeId) ? string.Format(";#{0}", term.Id.ToString()) : "");
                    termsString.Add(string.Format("{0}{3}{1}{3}{2}", groupPath, termsetPath, termPath, delimiter));

                    if (term.TermsCount > 0)
                    {
                        var subTermPath = string.Format("{0}{3}{1}{3}{2}", groupPath, termsetPath, termPath, delimiter);

                        termsString.AddRange(ParseSubTerms(subTermPath, term, includeId, delimiter, clientContext));
                    }

                    termsString.Add(string.Format("{0}{3}{1}{3}{2}", groupPath, termsetPath, termPath, delimiter));
                }
            }

            return termsString.Distinct().ToList<string>();
        }

        /// <summary>
        /// Exports the full list of terms from all termsets in all termstores.
        /// </summary>
        /// <param name="site">The site to process</param>
        /// <param name="includeId">if true, Ids of the the taxonomy items will be included</param>
        /// <param name="delimiter">if specified, this delimiter will be used. Notice that IDs will be delimited with ;# from the label</param>
        /// <returns>Returns list of Term strings</returns>
        public static List<string> ExportAllTerms(this Site site, bool includeId, string delimiter = "|")
        {
            var clientContext = site.Context;

            var termsString = new List<string>();

            TaxonomySession taxonomySession = taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);

            clientContext.ExecuteQueryRetry();

            var termStores = taxonomySession.TermStores;
            clientContext.Load(termStores, t => t.IncludeWithDefaultProperties(s => s.Groups));
            clientContext.ExecuteQueryRetry();
            foreach (var termStore in termStores)
            {
                foreach (var termGroup in termStore.Groups)
                {
                    var termSets = termGroup.TermSets;
                    clientContext.Load(termSets, t => t.IncludeWithDefaultProperties(s => s.Terms));
                    clientContext.ExecuteQueryRetry();
                    var termGroupName = DenormalizeName(termGroup.Name);
                    var groupPath = $"{termGroupName}{((includeId) ? $";#{termGroup.Id}" : "")}";
                    foreach (var set in termSets)
                    {
                        var setName = DenormalizeName(set.Name);
                        var termsetPath = string.Format("{0}{3}{1}{2}", groupPath, setName, (includeId) ? $";#{set.Id}" : "", delimiter);
                        foreach (var term in set.Terms)
                        {
                            var termName = DenormalizeName(term.Name);
                            var termPath = string.Format("{0}{3}{1}{2}", termsetPath, termName, (includeId) ?
                                $";#{term.Id.ToString()}"
                                : "", delimiter);
                            termsString.Add(termPath);

                            if (term.TermsCount > 0)
                            {
                                termsString.AddRange(ParseSubTerms(termPath, term, includeId, delimiter, clientContext));
                            }
                        }
                    }
                }
            }

            return termsString.Distinct().ToList<string>();
        }

        private static List<string> ParseSubTerms(string subTermPath, Term term, bool includeId, string delimiter, ClientRuntimeContext clientContext)
        {
            var items = new List<string>();
            if (term.ServerObjectIsNull == null || term.ServerObjectIsNull == false)
            {
                clientContext.Load(term.Terms);
                clientContext.ExecuteQueryRetry();
            }

            foreach (var subTerm in term.Terms)
            {
                var termName = DenormalizeName(subTerm.Name);
                var termPath = string.Format("{0}{3}{1}{2}", subTermPath, termName, (includeId) ? string.Format(";#{0}", subTerm.Id.ToString()) : "", delimiter);

                items.Add(termPath);

                if (term.TermsCount > 0)
                {
                    items.AddRange(ParseSubTerms(termPath, subTerm, includeId, delimiter, clientContext));
                }

            }
            return items;
        }

        /// <summary>
        /// Normalizes a Taxonomy name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string NormalizeName(string name)
        {
            if (name == null) return (string)null;
            name = TrimSpacesRegex.Replace(name, " ").Replace('&', '＆');

            if (!name.Contains(",") || !name.StartsWith("\"") || !name.EndsWith("\""))
            {
                name = name.Replace('"', '＂');
            }
            return name;
        }

        /// <summary>
        /// Denormalizes a Taxonomy name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string DenormalizeName(string name)
        {
            if (name == null)
                return (string)null;

            name = TrimSpacesRegex.Replace(name, " ").Replace('＆', '&').Replace('＂', '"');
            if (name.Contains(",") && !name.StartsWith("\"") && !name.EndsWith("\""))
            {
                name = '"' + name + '"'; //Add quotes for terms with comma, if not parsing breaks on import
            }
            return name;
        }

        /// <summary>
        /// Returns a taxonomy item by it's path, e.g. Group|Set|Term
        /// </summary>
        /// <param name="site">The current site</param>
        /// <param name="path">The path of the item to return</param>
        /// <param name="delimiter">The delimeter separating groups, sets and term in the path. Defaults to |</param>
        /// <returns>Returns TaxonomyItem object</returns>
        public static TaxonomyItem GetTaxonomyItemByPath(this Site site, string path, string delimiter = "|")
        {
            
            var context = site.Context;

            if (string.IsNullOrEmpty(path)) throw new ArgumentNullException("path");

            

            var pathSplit = path.Split(new string[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);


            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(context);
            TermStore ts = tSession.GetDefaultKeywordsTermStore();

            Term term = null;

            if (pathSplit.Length == 2 && Guid.TryParse(pathSplit[1], out Guid termid))
            {
                term = ts.GetTerm(termid);
                context.Load(term);
                context.ExecuteQueryRetry();
            }
            else
            {
                var groups = context.LoadQuery(ts.Groups);
                context.ExecuteQueryRetry();

                var group = groups.FirstOrDefault(l => l.Name.Equals(pathSplit[0], StringComparison.CurrentCultureIgnoreCase));
                if (group == null) return null;
                if (pathSplit.Length == 1) return group;

                var termSets = context.LoadQuery(group.TermSets);
                context.ExecuteQueryRetry();

                var termSet = termSets.FirstOrDefault(l => l.Name.Equals(pathSplit[1], StringComparison.CurrentCultureIgnoreCase));
                if (termSet == null) return null;
                if (pathSplit.Length == 2) return termSet;


                for (int i = 2; i < pathSplit.Length; i++)
                {
                    IEnumerable<Term> termColl = context.LoadQuery(i == 2 ? termSet.Terms : term.Terms);
                    context.ExecuteQueryRetry();

                    term = termColl.FirstOrDefault(l => l.Name.Equals(pathSplit[i], StringComparison.OrdinalIgnoreCase));

                    if (term == null) return null;
                }
            }

            return term;
        }

        private static void ValidateDescription(string description, string parameterName)
        {
            if (string.IsNullOrEmpty(description))
            {
                return;
            }
            if (InvalidDescriptionRegex.IsMatch(description))
            {
                throw new ArgumentException(string.Format("Invalid characters in description '{0}'.", new object[]
                {
                    description
                }), parameterName);
            }
            if (description.Length > 1000)
            {
                throw new ArgumentException(string.Format("Description exceeds maximum length (1000): '{0}'.", new object[]
                {
                    description
                }), parameterName);
            }
        }

        private static void ValidateName(string name, string parameterName)
        {
            if (string.IsNullOrEmpty(name)) { throw new ArgumentNullException(parameterName); }

            if (name.Length > 255 || InvalidNameRegex.IsMatch(name))
            {
                throw new ArgumentException(string.Format("Invalid taxonomy name '{0}'.", new object[]
                {
                    name
                }), parameterName);
            }
        }

        /// <summary>
        /// Ensures the specified label for the specified lcid exists.
        /// </summary>
        /// <param name="term">The term to ensure the label for</param>
        /// <param name="lcid">The LCID of the label to ensure</param>
        /// <param name="labelName">The name of the label to ensure</param>
        /// <param name="isDefault">Determines if the label should be the default</param>
        public static void EnsureLabel(this Term term, int lcid, string labelName, bool isDefault)
        {
            var clientContext = term.Context;

            if (term.ServerObjectIsNull == true)
            {
                clientContext.Load(term);
                clientContext.ExecuteQueryRetry();
            }

            clientContext.Load(term.Labels);
            clientContext.ExecuteQueryRetry();

            if (!term.Labels.Where(l => l.Language == lcid).Any(l => l.Value == labelName))
            {
                term.CreateLabel(labelName, lcid, isDefault);
                clientContext.ExecuteQueryRetry();
            }
        }

        #endregion

        #region Fields
        /// <summary>
        /// Sets a value in a taxonomy field
        /// </summary>
        /// <param name="item">The item to set the value to</param>
        /// <param name="TermPath">The path of the term in the shape of "TermGroupName|TermSetName|TermName"</param>
        /// <param name="fieldId">The id of the field</param>
        /// <param name="systemUpdate">If set to true, will do a system udpate to the item. Default value is false.</param>
        /// <exception cref="KeyNotFoundException"/>
        public static void SetTaxonomyFieldValueByTermPath(this ListItem item, string TermPath, Guid fieldId, bool systemUpdate = false)
        {
            var clientContext = item.Context as ClientContext;
            TaxonomyItem taxItem = clientContext.Site.GetTaxonomyItemByPath(TermPath);
            if (taxItem != null)
            {
                item.SetTaxonomyFieldValue(fieldId, taxItem.Name, taxItem.Id, systemUpdate);
            }
            else
            {
                throw new KeyNotFoundException("Taxonomy Term not found");
            }
        }

        /// <summary>
        /// Sets a value of a taxonomy field. To set an empty value set label to an empty string and termGuid to an empty GUID.
        /// </summary>
        /// <param name="item">The item to process</param>
        /// <param name="fieldId">The ID of the field to set</param>
        /// <param name="label">The label of the term to set</param>
        /// <param name="termGuid">The id of the term to set</param>
        /// <param name="systemUpdate">If set to true, will do a system udpate to the item. Default value is false.</param>
        public static void SetTaxonomyFieldValue(this ListItem item, Guid fieldId, string label, Guid termGuid, bool systemUpdate = false)
        {
            ClientContext clientContext = item.Context as ClientContext;

            var field = item.ParentList.Fields.GetById(fieldId);
            TaxonomyField taxField = clientContext.CastTo<TaxonomyField>(field);
            clientContext.ExecuteQueryRetry();

            if (string.IsNullOrEmpty(label) && termGuid.Equals(Guid.Empty))
            {
                taxField.SetFieldValueByLabelGuidPair(item, string.Empty, systemUpdate);
            }
            else
            {
                taxField.SetFieldValueByLabelGuidPair(item, $"{label}|{termGuid.ToString()}", systemUpdate);
            }
        }

        /// <summary>
        /// Sets a value of a taxonomy field that supports multiple values
        /// </summary>
        /// <param name="item">The item to process</param>
        /// <param name="fieldId">The ID of the field to set</param>
        /// <param name="termValues">The key and values of terms to set</param>
        /// <param name="systemUpdate">If set to true, will do a system udpate to the item. Default value is false.</param>
        public static void SetTaxonomyFieldValues(this ListItem item, Guid fieldId, IEnumerable<KeyValuePair<Guid, String>> termValues, bool systemUpdate = false)
        {
            ClientContext clientContext = item.Context as ClientContext;

            var field = item.ParentList.Fields.GetById(fieldId);
            TaxonomyField taxField = clientContext.CastTo<TaxonomyField>(field);
            clientContext.Load(taxField, tf => tf.AllowMultipleValues, tf => tf.StaticName);
            clientContext.ExecuteQueryRetry();

            if (taxField.AllowMultipleValues)
            {
                StringBuilder termValuesStringbuilder = new StringBuilder();
                bool skipSeperator = true;
                foreach (var term in termValues)
                {
                    if (skipSeperator)
                    {
                        skipSeperator = false;
                    }
                    else
                    {
                        termValuesStringbuilder.Append(';');
                    }
                    termValuesStringbuilder.Append(term.Value + "|" + term.Key.ToString());
                }

                taxField.SetFieldValueByLabelGuidPair(item, termValuesStringbuilder.ToString(), systemUpdate);
            }
            else
            {
                throw new ArgumentException(string.Format(CoreResources.TaxonomyExtensions_Field_Is_Not_Multivalues, taxField.StaticName));
            }
        }

        /// <summary>
        /// Sets a value of a taxonomy field.
        /// Value parameter is one or more label GUID pairs:
        /// Single value field (TaxonomyFieldType) - term label|term GUID
        /// Multi value field (TaxonomyFieldTypeMulti) - term label|term GUID;term label|term GUID;term label|term GUID...
        /// </summary>
        /// <param name="taxonomyField">The field to set</param>
        /// <param name="item">The item to process</param>
        /// <param name="value">The value to set on the taxonomy field</param>
        /// <param name="systemUpdate">If set to true, will do a system udpate to the item. Default value is false.</param>
        public static void SetFieldValueByLabelGuidPair(this TaxonomyField taxonomyField, ListItem item, string value, bool systemUpdate = false)
        {
            taxonomyField.EnsureProperties(f => f.TextField, f => f.AllowMultipleValues);

            //TaxonomyFieldValue.PopulateFromLabelGuidPairs(string) has not been exposed to CSOM.
            //This trick uses TaxonomyFieldValueCollection to get the same result for single value taxonomy fields.
            TaxonomyFieldValueCollection taxonomyValues = new TaxonomyFieldValueCollection(item.Context, null, taxonomyField);
            taxonomyValues.PopulateFromLabelGuidPairs(value);

            if (taxonomyField.AllowMultipleValues)
            {
                taxonomyField.SetFieldValueByValueCollection(item, taxonomyValues);
            }
            else
            {
                item.Context.Load(taxonomyValues);
                item.Context.ExecuteQueryRetry();
                if (taxonomyValues.Count > 0)
                {
                    taxonomyField.SetFieldValueByValue(item, taxonomyValues[0]);
                }
                else
                {
                    //Empty single value taxonomy value specified. Clear out existing value.
                    //It's not possible to clear out the value using an empty TaxonomyFieldValue directly.
                    //This is due to the fact that SetFieldValueByValue requires TermGuid or creatingField to be set on server side.
                    //It is not possible to set creatingField, so the only possible workaround is to use the predefined Guid with all 1's.
                    //This will leave data in the hidden field that must be cleared in the same query on the server.
                    //This approach is necessary to maintain data in TaxCatchAll field for other taxonomy fields in the list item.

                    Field hiddenField = item.ParentList.Fields.GetById(taxonomyField.TextField);
                    item.Context.Load(hiddenField,
                        tf => tf.InternalName
                        );
                    item.Context.ExecuteQueryRetry();

                    TaxonomyFieldValue taxonomyValue = new TaxonomyFieldValue();
                    taxonomyValue.Label = string.Empty;
                    taxonomyValue.TermGuid = "11111111-1111-1111-1111-111111111111";
                    taxonomyValue.WssId = -1;
                    taxonomyField.SetFieldValueByValue(item, taxonomyValue);

                    item[hiddenField.InternalName] = string.Empty;
                }
            }
            if (systemUpdate)
            {
#if !SP2013 && !SP2016
                item.SystemUpdate();
#else
                item.Update();
#endif
            }
            else
            {
                item.Update();
            }
            item.Context.ExecuteQueryRetry();
        }

        private static void CleanupTaxonomyHiddenField(Web web, FieldCollection fields, TaxonomyFieldCreationInformation fieldCreationInformation)
        {
            // if the Guid is empty then we'll have no issue
            if (fieldCreationInformation.Id != Guid.Empty)
            {
                FieldCollection _fields = fields;
                web.Context.Load(_fields, fc => fc.Include(f => f.Id, f => f.InternalName, f => f.Hidden));
                web.Context.ExecuteQueryRetry();
                var _field = _fields.FirstOrDefault(f => f.InternalName.Equals(fieldCreationInformation.InternalName));
                // if the field does not exist we assume the possiblity that it was created earlier then deleted and the hidden field was left behind
                // if the field does exist then return and let the calling process exception out when attempting to create it
                // this does not appear to be an issue with lists, just site columns, but it doesnt hurt to check
                if (_field == null)
                {
                    // The hidden field format is the id of the field itself with hyphens removed and the first character replaced
                    // with a random character, so get everything to the right of the first character and remove hyphens
                    var _hiddenField = fieldCreationInformation.Id.ToString().Replace("-", "").Substring(1);
                    _field = _fields.FirstOrDefault(f => f.InternalName.EndsWith(_hiddenField));
                    if (_field != null)
                    {
                        if (_field.Hidden)
                        {
                            // just in case the field itself is hidden, make sure it is not because depending on the current CU hidden fields may not be deletable
                            _field.Hidden = false;
                            _field.Update();
                        }
                        _field.DeleteObject();
                        web.Context.ExecuteQueryRetry();
                    }
                }
            }
        }
        /// <summary>
        /// Can be used to create taxonomy field remotely to web.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="fieldCreationInformation">Creation Information of the field</param>
        /// <returns>New taxonomy field</returns>
        public static Field CreateTaxonomyField(this Web web, TaxonomyFieldCreationInformation fieldCreationInformation)
        {
            fieldCreationInformation.InternalName.ValidateNotNullOrEmpty("internalName");
            fieldCreationInformation.DisplayName.ValidateNotNullOrEmpty("displayName");
            fieldCreationInformation.TaxonomyItem.ValidateNotNullOrEmpty("taxonomyItem");

            CleanupTaxonomyHiddenField(web, web.Fields, fieldCreationInformation);

            if (fieldCreationInformation.Id == Guid.Empty)
            {
                fieldCreationInformation.Id = Guid.NewGuid();
            }

            var showFieldAttribute = new KeyValuePair<string, string>();
            if (fieldCreationInformation.AdditionalAttributes != null)
            {
                showFieldAttribute = fieldCreationInformation.AdditionalAttributes.FirstOrDefault(a => a.Key == "ShowField");
            }
            if (showFieldAttribute.Key == null)
            {
                if (fieldCreationInformation.AdditionalAttributes == null)
                {
                    fieldCreationInformation.AdditionalAttributes = new List<KeyValuePair<string, string>>();
                }
                ((List<KeyValuePair<string, string>>)fieldCreationInformation.AdditionalAttributes).Add(new KeyValuePair<string, string>("ShowField", "Term1033"));
            }

            var _field = web.CreateField(fieldCreationInformation);

            WireUpTaxonomyFieldInternal(_field, fieldCreationInformation.TaxonomyItem, fieldCreationInformation.MultiValue);
            _field.Update();

            web.Context.ExecuteQueryRetry();

            return _field;

        }

        /// <summary>
        /// Removes a taxonomy field (site column) and its associated hidden field by internal name
        /// </summary>
        /// <param name="web">Web object were the field (site column) exists</param>
        /// <param name="internalName">Internal name of the taxonomy field (site column) to be removed</param>
        public static void RemoveTaxonomyFieldByInternalName(this Web web, string internalName)
        {
            FieldCollection fields = web.Fields;
            web.Context.Load(fields, fc => fc.Include(f => f.Id, f => f.InternalName));
            web.Context.ExecuteQueryRetry();

            Field field = fields.FirstOrDefault(f => f.InternalName == internalName);

            if (field != null)
            {
                field.DeleteObject();
                web.Update();
                web.Context.ExecuteQueryRetry();

                CleanupTaxonomyHiddenField(web, web.Fields, new TaxonomyFieldCreationInformation() { Id = field.Id, InternalName = field.InternalName });

            }
        }

        /// <summary>
        /// Removes a taxonomy field (site column) and its associated hidden field by id
        /// </summary>
        /// <param name="web">Web object were the field (site column) exists</param>
        /// <param name="id">Guid representing the id of the taxonomy field (site column) to be removed</param>
        public static void RemoveTaxonomyFieldById(this Web web, Guid id)
        {

            FieldCollection fields = web.Fields;
            web.Context.Load(fields, fc => fc.Include(f => f.Id, f => f.InternalName));
            web.Context.ExecuteQueryRetry();

            Field field = fields.FirstOrDefault(f => f.Id == id);

            if (field != null)
            {
                field.DeleteObject();
                web.Update();
                web.Context.ExecuteQueryRetry();

                CleanupTaxonomyHiddenField(web, web.Fields, new TaxonomyFieldCreationInformation() { Id = id, InternalName = field.InternalName });
            }
        }

        /// <summary>
        /// Can be used to create taxonomy field remotely in a list. 
        /// </summary>
        /// <param name="list">List to be processed</param>
        /// <param name="fieldCreationInformation">Creation information of the field</param>
        /// <returns>New taxonomy field</returns>
        public static Field CreateTaxonomyField(this List list, TaxonomyFieldCreationInformation fieldCreationInformation)
        {
            fieldCreationInformation.InternalName.ValidateNotNullOrEmpty("internalName");
            fieldCreationInformation.DisplayName.ValidateNotNullOrEmpty("displayName");
            fieldCreationInformation.TaxonomyItem.ValidateNotNullOrEmpty("taxonomyItem");

            CleanupTaxonomyHiddenField(list.ParentWeb, list.Fields, fieldCreationInformation);

            if (fieldCreationInformation.Id == Guid.Empty)
            {
                fieldCreationInformation.Id = Guid.NewGuid();
            }
            var showFieldAttribute = new KeyValuePair<string, string>();
            if (fieldCreationInformation.AdditionalAttributes != null)
            {
                showFieldAttribute = fieldCreationInformation.AdditionalAttributes.FirstOrDefault(a => a.Key == "ShowField");
            }
            if (showFieldAttribute.Key == null)
            {
                if (fieldCreationInformation.AdditionalAttributes == null)
                {
                    fieldCreationInformation.AdditionalAttributes = new List<KeyValuePair<string, string>>();
                }
                ((List<KeyValuePair<string, string>>)fieldCreationInformation.AdditionalAttributes).Add(new KeyValuePair<string, string>("ShowField", "Term1033"));
            }
            var _field = list.CreateField(fieldCreationInformation);

            WireUpTaxonomyFieldInternal(_field, fieldCreationInformation.TaxonomyItem, fieldCreationInformation.MultiValue);
            _field.Update();

            list.Context.ExecuteQueryRetry();

            return _field;
        }

        /// <summary>
        /// Wires up MMS field to the specified term set.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="field">Field to be wired up</param>
        /// <param name="mmsGroupName">Taxonomy group</param>
        /// <param name="mmsTermSetName">Term set name</param>
        /// <param name="multiValue">If true, create a multivalue field</param>
        public static void WireUpTaxonomyField(this Web web, Field field, string mmsGroupName, string mmsTermSetName, bool multiValue = false)
        {
            TermStore termStore = GetDefaultTermStore(web);

            if (termStore == null)
                throw new NullReferenceException("The default term store is not available.");

            if (string.IsNullOrEmpty(mmsTermSetName))
                throw new ArgumentNullException(nameof(mmsTermSetName), "The MMS term set is not specified.");

            // get the term group and term set
            TermGroup termGroup = termStore.Groups.GetByName(mmsGroupName);
            TermSet termSet = termGroup.TermSets.GetByName(mmsTermSetName);
            web.Context.Load(termStore);
            web.Context.Load(termSet);
            web.Context.ExecuteQueryRetry();

            WireUpTaxonomyField(web, field, termSet, multiValue);
        }

        /// <summary>
        /// Wires up MMS field to the specified term set.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="field">Field to be wired up</param>
        /// <param name="termSet">Taxonomy TermSet</param>
        /// <param name="multiValue">If true, create a multivalue field</param>
        public static void WireUpTaxonomyField(this Web web, Field field, TermSet termSet, bool multiValue = false)
        {
            WireUpTaxonomyFieldInternal(field, termSet, multiValue);
        }

        /// <summary>
        /// Wires up MMS field to the specified term.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="field">Field to be wired up</param>
        /// <param name="anchorTerm">Taxonomy Term</param>
        /// <param name="multiValue">If true, create a multivalue field</param>
        public static void WireUpTaxonomyField(this Web web, Field field, Term anchorTerm, bool multiValue = false)
        {
            WireUpTaxonomyFieldInternal(field, anchorTerm, multiValue);
        }

        /// <summary>
        /// Wires up MMS field to the specified term set.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="id">Field ID to be wired up</param>
        /// <param name="mmsGroupName">Taxonomy group</param>
        /// <param name="mmsTermSetName">Term set name</param>
        /// <param name="multiValue">If true, create a multivalue field</param>
        public static void WireUpTaxonomyField(this Web web, Guid id, string mmsGroupName, string mmsTermSetName, bool multiValue = false)
        {
            var field = web.Fields.GetById(id);
            web.Context.Load(field);
            web.WireUpTaxonomyField(field, mmsGroupName, mmsTermSetName, multiValue);
        }

        /// <summary>
        /// Wires up MMS field to the specified term set.
        /// </summary>
        /// <param name="list">List to be processed</param>
        /// <param name="field">Field to be wired up</param>
        /// <param name="termSet">Taxonomy TermSet</param>
        /// <param name="multiValue">Term set name</param>
        public static void WireUpTaxonomyField(this List list, Field field, TermSet termSet, bool multiValue = false)
        {
            WireUpTaxonomyFieldInternal(field, termSet, multiValue);
        }

        /// <summary>
        /// Wires up MMS field to the specified term.
        /// </summary>
        /// <param name="list">List to be processed</param>
        /// <param name="field">Field to be wired up</param>
        /// <param name="anchorTerm">Taxonomy Term</param>
        /// <param name="multiValue">Allow multiple selection</param>
        public static void WireUpTaxonomyField(this List list, Field field, Term anchorTerm, bool multiValue = false)
        {
            WireUpTaxonomyFieldInternal(field, anchorTerm, multiValue);
        }

        /// <summary>
        /// Wires up MMS field to the specified term set.
        /// </summary>
        /// <param name="list">List to be processed</param>
        /// <param name="field">Field to be wired up</param>
        /// <param name="mmsGroupName">Taxonomy group</param>
        /// <param name="mmsTermSetName">Term set name</param>
        /// <param name="multiValue">Allow multiple selection</param>
        public static void WireUpTaxonomyField(this List list, Field field, string mmsGroupName, string mmsTermSetName, bool multiValue = false)
        {
            var clientContext = list.Context as ClientContext;
            TermStore termStore = clientContext.Site.GetDefaultSiteCollectionTermStore();

            if (termStore == null)
                throw new NullReferenceException("The default term store is not available.");

            if (string.IsNullOrEmpty(mmsTermSetName))
                throw new ArgumentNullException("mmsTermSetName", "The MMS term set is not specified.");

            // get the term group and term set
            TermGroup termGroup = termStore.Groups.GetByName(mmsGroupName);
            TermSet termSet = termGroup.TermSets.GetByName(mmsTermSetName);
            clientContext.Load(termStore);
            clientContext.Load(termSet);
            clientContext.ExecuteQueryRetry();

            list.WireUpTaxonomyField(field, termSet, multiValue);
        }

        /// <summary>
        /// Wires up MMS field to the specified term set.
        /// </summary>
        /// <param name="list">List to be processed</param>
        /// <param name="id">Field ID to be wired up</param>
        /// <param name="mmsGroupName">Taxonomy group</param>
        /// <param name="mmsTermSetName">Term set name</param>
        /// <param name="multiValue">Allow multiple selection</param>
        public static void WireUpTaxonomyField(this List list, Guid id, string mmsGroupName, string mmsTermSetName, bool multiValue = false)
        {
            var clientContext = list.Context as ClientContext;
            var field = list.Fields.GetById(id);
            clientContext.Load(field);
            list.WireUpTaxonomyField(field, mmsGroupName, mmsTermSetName, multiValue);
        }

        /// <summary>
        /// Wires up MMS field to the specified term set or term.
        /// </summary>
        /// <param name="field">Field to be wired up</param>
        /// <param name="taxonomyItem">Taxonomy TermSet or Term</param>
        /// <param name="multiValue">Allow multiple selection</param>
        private static void WireUpTaxonomyFieldInternal(Field field, TaxonomyItem taxonomyItem, bool multiValue)
        {
            var clientContext = field.Context as ClientContext;

            taxonomyItem.ValidateNotNullOrEmpty("taxonomyItem");

            var anchorTerm = taxonomyItem as Term;

            if (anchorTerm != default(Term) && !anchorTerm.IsPropertyAvailable("TermSet"))
            {
                clientContext.Load(anchorTerm.TermSet);
                clientContext.ExecuteQueryRetry();
            }

            var termSet = taxonomyItem is Term ? anchorTerm.TermSet : taxonomyItem as TermSet;

            if (termSet == default(TermSet))
                throw new ArgumentException("Bound TaxonomyItem must be either a TermSet or a Term");

            termSet.EnsureProperties(ts => ts.TermStore);

            // set the SSP ID and Term Set ID on the taxonomy field
            var taxField = clientContext.CastTo<TaxonomyField>(field);
            taxField.SspId = termSet.TermStore.Id;
            taxField.TermSetId = termSet.Id;

            if (anchorTerm != default(Term))
            {
                taxField.AnchorId = anchorTerm.Id;
            }

            taxField.AllowMultipleValues = multiValue;
            taxField.Update();
            clientContext.ExecuteQueryRetry();
        }

        /// <summary>
        /// Returns the Id for a term if present in the TaxonomyHiddenList. Otherwise returns -1;
        /// </summary>
        /// <param name="web"></param>
        /// <param name="term"></param>
        /// <returns></returns>
        public static int GetWssIdForTerm(this Web web, Term term)
        {
            var clientContext = web.Context as ClientContext;
            var list = clientContext.Site.RootWeb.GetListByUrl("Lists/TaxonomyHiddenList");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml =
                $@"<View><Query><Where><Eq><FieldRef Name='IdForTerm' /><Value Type='Text'>{term.Id}</Value></Eq></Where></Query></View>";

            var items = list.GetItems(camlQuery);
            web.Context.Load(items);
            web.Context.ExecuteQueryRetry();

            if (items.Any())
            {
                return items[0].Id;
            }
            else
            {
                return -1;
            }
        }

        /// <summary>
        /// Sets the default value for a managed metadata field
        /// </summary>
        /// <param name="field">Field to be wired up</param>
        /// <param name="taxonomyItem">Taxonomy TermSet or Term</param>
        /// <param name="defaultValue">default value for the field</param>
        /// <param name="pushChangesToLists">push changes to lists</param>
        public static void SetTaxonomyFieldDefaultValue(this Field field, TaxonomyItem taxonomyItem, string defaultValue, bool pushChangesToLists = false)
        {
            if (string.IsNullOrEmpty(defaultValue))
            {
                throw new ArgumentException("defaultValue");
            }

            var clientContext = field.Context as ClientContext;

            taxonomyItem.ValidateNotNullOrEmpty("taxonomyItem");

            var anchorTerm = taxonomyItem as Term;

            if (anchorTerm != default(Term) && !anchorTerm.IsPropertyAvailable("TermSet"))
            {
                clientContext.Load(anchorTerm.TermSet);
                clientContext.ExecuteQueryRetry();
            }

            var termSet = taxonomyItem is Term ? anchorTerm.TermSet : taxonomyItem as TermSet;

            if (termSet == default(TermSet))
            {
                throw new ArgumentException("Bound TaxonomyItem must be either a TermSet or a Term");
            }

            // set the SSP ID and Term Set ID on the taxonomy field
            var taxField = clientContext.CastTo<TaxonomyField>(field);


            if (!termSet.IsPropertyAvailable("Terms"))
            {
                clientContext.Load(termSet.Terms);
                clientContext.ExecuteQueryRetry();
            }

            Term defaultValTerm = termSet.Terms.GetByName(defaultValue);
            if (defaultValTerm != null)
            {
                clientContext.Load(defaultValTerm);
                clientContext.ExecuteQueryRetry();

                TaxonomyFieldValue taxValue = new TaxonomyFieldValue();
                taxValue.WssId = -1;
                taxValue.TermGuid = defaultValTerm.Id.ToString();
                taxValue.Label = defaultValTerm.Name;
                //get validate string
                var validateValue = taxField.GetValidatedString(taxValue);
                field.Context.ExecuteQueryRetry();

                taxField.DefaultValue = validateValue.Value;
                if (!pushChangesToLists)
                {
                    taxField.Update();
                }
                else
                {
                    taxField.UpdateAndPushChanges(true);
                }
                clientContext.ExecuteQueryRetry();
            }
        }
        #endregion
    }
}
