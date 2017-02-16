using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client.InformationPolicy;
using OfficeDevPnP.Core.Entities;

namespace Microsoft.SharePoint.Client
{

    /// <summary>
    /// Class that deals with information management features
    /// </summary>
    public static partial class InformationManagementExtensions
    {

        /// <summary>
        /// Does this web have a site policy applied?
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if a policy has been applied, false otherwise</returns>
        public static bool HasSitePolicyApplied(this Web web)
        {
            var hasSitePolicyApplied = ProjectPolicy.DoesProjectHavePolicy(web.Context, web);
            web.Context.ExecuteQueryRetry();
            return hasSitePolicyApplied.Value;
        }

        /// <summary>
        /// Gets the site expiration date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the expiration date, DateTime.MinValue in case there was no policy applied</returns>
        public static DateTime GetSiteExpirationDate(this Web web)
        {
            if (web.HasSitePolicyApplied())
            {
                var expirationDate = ProjectPolicy.GetProjectExpirationDate(web.Context, web);
                web.Context.ExecuteQueryRetry();
                return expirationDate.Value;
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Gets the site closure date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the closure date, DateTime.MinValue in case there was no policy applied</returns>
        public static DateTime GetSiteCloseDate(this Web web)
        {
            if (web.HasSitePolicyApplied())
            {
                var closeDate = ProjectPolicy.GetProjectCloseDate(web.Context, web);
                web.Context.ExecuteQueryRetry();
                return closeDate.Value;
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Gets a list of the available site policies
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A list of <see cref="SitePolicyEntity"/> objects</returns>
        public static List<SitePolicyEntity> GetSitePolicies(this Web web)
        {
            var sitePolicies = ProjectPolicy.GetProjectPolicies(web.Context, web);
            web.Context.Load(sitePolicies);
            web.Context.ExecuteQueryRetry();

            var policies = new List<SitePolicyEntity>();

            if (sitePolicies != null && sitePolicies.Count > 0)
            {
                foreach (var policy in sitePolicies)
                {
                    policies.Add(new SitePolicyEntity
                    {
                        Name = policy.Name,
                        Description = policy.Description,
                        EmailBody = policy.EmailBody,
                        EmailBodyWithTeamMailbox = policy.EmailBodyWithTeamMailbox,
                        EmailSubject = policy.EmailSubject
                    });
                }
            }

            return policies;
        }

        /// <summary>
        /// Gets the site policy that currently is applied
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the applied policy</returns>
        public static SitePolicyEntity GetAppliedSitePolicy(this Web web)
        {
            if (web.HasSitePolicyApplied())
            {
                var policy = ProjectPolicy.GetCurrentlyAppliedProjectPolicyOnWeb(web.Context, web);
                web.Context.Load(policy,
                             p => p.Name,
                             p => p.Description,
                             p => p.EmailSubject,
                             p => p.EmailBody,
                             p => p.EmailBodyWithTeamMailbox);
                web.Context.ExecuteQueryRetry();
                return new SitePolicyEntity
                {
                    Name = policy.Name,
                    Description = policy.Description,
                    EmailBody = policy.EmailBody,
                    EmailBodyWithTeamMailbox = policy.EmailBodyWithTeamMailbox,
                    EmailSubject = policy.EmailSubject
                };
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the site policy with the given name
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Site policy to fetch</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the fetched policy</returns>
        public static SitePolicyEntity GetSitePolicyByName(this Web web, string sitePolicy)
        {
            var policies = web.GetSitePolicies();

            if (policies.Count > 0)
            {
                var policy = policies.FirstOrDefault(p => p.Name == sitePolicy);
                return policy;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Apply a policy to a site
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Policy to apply</param>
        /// <returns>True if applied, false otherwise</returns>
        public static bool ApplySitePolicy(this Web web, string sitePolicy)
        {
            var result = false;

            var sitePolicies = ProjectPolicy.GetProjectPolicies(web.Context, web);
            web.Context.Load(sitePolicies);
            web.Context.ExecuteQueryRetry();

            if (sitePolicies != null && sitePolicies.Count > 0)
            {
                var policyToApply = sitePolicies.FirstOrDefault(p => p.Name == sitePolicy);

                if (policyToApply != null)
                {
                    ProjectPolicy.ApplyProjectPolicy(web.Context, web, policyToApply);
                    web.Context.ExecuteQueryRetry();
                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Check if a site is closed
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if site is closed, false otherwise</returns>
        public static bool IsClosedBySitePolicy(this Web web)
        {
            var isClosed = ProjectPolicy.IsProjectClosed(web.Context, web);
            web.Context.ExecuteQueryRetry();
            return isClosed.Value;
        }

        /// <summary>
        /// Close a site, if it has a site policy applied and is currently not closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was closed, false otherwise</returns>
        public static bool SetClosedBySitePolicy(this Web web)
        {
            if (web.HasSitePolicyApplied() && !IsClosedBySitePolicy(web))
            {
                ProjectPolicy.CloseProject(web.Context, web);
                web.Context.ExecuteQueryRetry();
                return true;
            }
            return false;
        }

        /// <summary>
        /// Open a site, if it has a site policy applied and is currently closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was opened, false otherwise</returns>
        public static bool SetOpenBySitePolicy(this Web web)
        {
            if (web.HasSitePolicyApplied() && IsClosedBySitePolicy(web))
            {
                ProjectPolicy.OpenProject(web.Context, web);
                web.Context.ExecuteQueryRetry();
                return true;
            }
            return false;
        }
    }
}
