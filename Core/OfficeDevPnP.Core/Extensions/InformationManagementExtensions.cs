using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.InformationPolicy;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities.Async;

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
#if !ONPREMISES || SP2019
            return Task.Run(() => web.HasSitePolicyAppliedImplementation()).GetAwaiter().GetResult();
#else
            return web.HasSitePolicyAppliedImplementation();
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Does this web have a site policy applied?
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if a policy has been applied, false otherwise</returns>
        public static async Task<bool> HasSitePolicyAppliedAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.HasSitePolicyAppliedImplementation();
        }
#endif
        /// <summary>
        /// Does this web have a site policy applied?
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if a policy has been applied, false otherwise</returns>
#if !ONPREMISES || SP2019
        private static async Task<bool> HasSitePolicyAppliedImplementation(this Web web)
#else
        private static bool HasSitePolicyAppliedImplementation(this Web web)
#endif
        {
            var hasSitePolicyApplied = ProjectPolicy.DoesProjectHavePolicy(web.Context, web);
#if !ONPREMISES || SP2019
            await web.Context.ExecuteQueryRetryAsync();
#else
            web.Context.ExecuteQueryRetry();
#endif
            return hasSitePolicyApplied.Value;
        }

        /// <summary>
        /// Gets the site expiration date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the expiration date, DateTime.MinValue in case there was no policy applied</returns>
        public static DateTime GetSiteExpirationDate(this Web web)
        {
#if !ONPREMISES || SP2019
            return Task.Run(() => web.GetSiteExpirationDateImplementation()).GetAwaiter().GetResult();
#else
            return web.GetSiteExpirationDateImplementation();
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Gets the site expiration date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the expiration date, DateTime.MinValue in case there was no policy applied</returns>
        public static async Task<DateTime> GetSiteExpirationDateAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.GetSiteExpirationDateImplementation();
        }
#endif
        /// <summary>
        /// Gets the site expiration date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the expiration date, DateTime.MinValue in case there was no policy applied</returns>
#if !ONPREMISES || SP2019
        private static async Task<DateTime> GetSiteExpirationDateImplementation(this Web web)
#else
        private static DateTime GetSiteExpirationDateImplementation(this Web web)
#endif
        {
#if !ONPREMISES || SP2019
            if (await web.HasSitePolicyAppliedImplementation())
#else
            if (web.HasSitePolicyAppliedImplementation())
#endif
            {
                var expirationDate = ProjectPolicy.GetProjectExpirationDate(web.Context, web);
#if !ONPREMISES || SP2019
                await web.Context.ExecuteQueryRetryAsync();
#else
                web.Context.ExecuteQueryRetry();
#endif
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
#if !ONPREMISES || SP2019
            return Task.Run(() => web.GetSiteCloseDateImplementation()).GetAwaiter().GetResult();
#else
            return web.GetSiteCloseDateImplementation();
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Gets the site closure date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the closure date, DateTime.MinValue in case there was no policy applied</returns>
        public static async Task<DateTime> GetSiteCloseDateAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.GetSiteCloseDateImplementation();
        }
#endif
        /// <summary>
        /// Gets the site closure date
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>DateTime value holding the closure date, DateTime.MinValue in case there was no policy applied</returns>
#if !ONPREMISES || SP2019
        private static async Task<DateTime> GetSiteCloseDateImplementation(this Web web)
#else
        private static DateTime GetSiteCloseDateImplementation(this Web web)
#endif
        {
#if !ONPREMISES || SP2019
            if (await web.HasSitePolicyAppliedImplementation())
#else
            if (web.HasSitePolicyAppliedImplementation())
#endif
            {
                var closeDate = ProjectPolicy.GetProjectCloseDate(web.Context, web);
#if !ONPREMISES || SP2019
                await web.Context.ExecuteQueryRetryAsync();
#else
                web.Context.ExecuteQueryRetry();
#endif
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
#if !ONPREMISES || SP2019
            return Task.Run(() => web.GetSitePoliciesImplementation()).GetAwaiter().GetResult();
#else
            return web.GetSitePoliciesImplementation();
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Gets a list of the available site policies
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A list of <see cref="SitePolicyEntity"/> objects</returns>
        public static async Task<List<SitePolicyEntity>> GetSitePoliciesAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.GetSitePoliciesImplementation();
        }
#endif
        /// <summary>
        /// Gets a list of the available site policies
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A list of <see cref="SitePolicyEntity"/> objects</returns>
#if !ONPREMISES || SP2019
        private static async Task<List<SitePolicyEntity>> GetSitePoliciesImplementation(this Web web)
#else
        private static List<SitePolicyEntity> GetSitePoliciesImplementation(this Web web)
#endif
        {
            var sitePolicies = ProjectPolicy.GetProjectPolicies(web.Context, web);
            web.Context.Load(sitePolicies);
#if !ONPREMISES || SP2019
            await web.Context.ExecuteQueryRetryAsync();
#else
            web.Context.ExecuteQueryRetry();
#endif

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
#if !ONPREMISES || SP2019
            return Task.Run(() => web.GetAppliedSitePolicyImplementation()).GetAwaiter().GetResult();
#else
            return web.GetAppliedSitePolicyImplementation();
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Gets the site policy that currently is applied
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the applied policy</returns>
        public static async Task<SitePolicyEntity> GetAppliedSitePolicyAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.GetAppliedSitePolicyImplementation();
        }
#endif
        /// <summary>
        /// Gets the site policy that currently is applied
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the applied policy</returns>
#if !ONPREMISES || SP2019
        private static async Task<SitePolicyEntity> GetAppliedSitePolicyImplementation(this Web web)
#else
        private static SitePolicyEntity GetAppliedSitePolicyImplementation(this Web web)
#endif
        {
#if !ONPREMISES || SP2019
            if (await web.HasSitePolicyAppliedImplementation())
#else
            if (web.HasSitePolicyAppliedImplementation())
#endif
            {
                var policy = ProjectPolicy.GetCurrentlyAppliedProjectPolicyOnWeb(web.Context, web);
                web.Context.Load(policy,
                             p => p.Name,
                             p => p.Description,
                             p => p.EmailSubject,
                             p => p.EmailBody,
                             p => p.EmailBodyWithTeamMailbox);
#if !ONPREMISES || SP2019
                await web.Context.ExecuteQueryRetryAsync();
#else
                web.Context.ExecuteQueryRetry();
#endif
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
#if !ONPREMISES || SP2019
            return Task.Run(() => web.GetSitePolicyByNameImplementation(sitePolicy)).GetAwaiter().GetResult();
#else
            return web.GetSitePolicyByNameImplementation(sitePolicy);
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Gets the site policy with the given name
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Site policy to fetch</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the fetched policy</returns>
        public static async Task<SitePolicyEntity> GetSitePolicyByNameAsync(this Web web, string sitePolicy)
        {
            await new SynchronizationContextRemover();
            return await web.GetSitePolicyByNameImplementation(sitePolicy);
        }
#endif
        /// <summary>
        /// Gets the site policy with the given name
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Site policy to fetch</param>
        /// <returns>A <see cref="SitePolicyEntity"/> object holding the fetched policy</returns>
#if !ONPREMISES || SP2019
        private static async Task<SitePolicyEntity> GetSitePolicyByNameImplementation(this Web web, string sitePolicy)
        {
            var policies = await web.GetSitePoliciesAsync();
#else
        private static SitePolicyEntity GetSitePolicyByNameImplementation(this Web web, string sitePolicy)
        {
            var policies = web.GetSitePolicies();
#endif

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
#if !ONPREMISES || SP2019
            return Task.Run(() => web.ApplySitePolicyImplementation(sitePolicy)).GetAwaiter().GetResult();
#else
            return web.ApplySitePolicyImplementation(sitePolicy);
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Apply a policy to a site
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Policy to apply</param>
        /// <returns>True if applied, false otherwise</returns>
        public static async Task<bool> ApplySitePolicyAsync(this Web web, string sitePolicy)
        {
            await new SynchronizationContextRemover();
            return await web.ApplySitePolicyImplementation(sitePolicy);
        }
#endif
        /// <summary>
        /// Apply a policy to a site
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <param name="sitePolicy">Policy to apply</param>
        /// <returns>True if applied, false otherwise</returns>
#if !ONPREMISES || SP2019
        private static async Task<bool> ApplySitePolicyImplementation(this Web web, string sitePolicy)

#else
        private static bool ApplySitePolicyImplementation(this Web web, string sitePolicy)
#endif
        {
            var result = false;

            var sitePolicies = ProjectPolicy.GetProjectPolicies(web.Context, web);
            web.Context.Load(sitePolicies);
#if !ONPREMISES || SP2019
            await web.Context.ExecuteQueryRetryAsync();
#else
            web.Context.ExecuteQueryRetry();
#endif

            if (sitePolicies != null && sitePolicies.Count > 0)
            {
                var policyToApply = sitePolicies.FirstOrDefault(p => p.Name == sitePolicy);

                if (policyToApply != null)
                {
                    ProjectPolicy.ApplyProjectPolicy(web.Context, web, policyToApply);
#if !ONPREMISES || SP2019
                    await web.Context.ExecuteQueryRetryAsync();
#else
                    web.Context.ExecuteQueryRetry();
#endif
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
#if !ONPREMISES || SP2019
            return Task.Run(() => web.IsClosedBySitePolicyImplementation()).GetAwaiter().GetResult();
#else
            return web.IsClosedBySitePolicyImplementation();
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Check if a site is closed
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if site is closed, false otherwise</returns>
        public static async Task<bool> IsClosedBySitePolicyAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.IsClosedBySitePolicyImplementation();
        }
#endif
        /// <summary>
        /// Check if a site is closed
        /// </summary>
        /// <param name="web">Web to operate on</param>
        /// <returns>True if site is closed, false otherwise</returns>
#if !ONPREMISES || SP2019
        private static async Task<bool> IsClosedBySitePolicyImplementation(this Web web)
#else
        private static bool IsClosedBySitePolicyImplementation(this Web web)
#endif
        {
            var isClosed = ProjectPolicy.IsProjectClosed(web.Context, web);
#if !ONPREMISES || SP2019
            await web.Context.ExecuteQueryRetryAsync();
#else
            web.Context.ExecuteQueryRetry();
#endif
            return isClosed.Value;
        }

        /// <summary>
        /// Close a site, if it has a site policy applied and is currently not closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was closed, false otherwise</returns>
        public static bool SetClosedBySitePolicy(this Web web)
        {
#if !ONPREMISES || SP2019
            return Task.Run(() => web.SetClosedBySitePolicyImplementation()).GetAwaiter().GetResult();
#else
            return web.SetClosedBySitePolicyImplementation();
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Close a site, if it has a site policy applied and is currently not closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was closed, false otherwise</returns>
        public static async Task<bool> SetClosedBySitePolicyAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.SetClosedBySitePolicyImplementation();
        }
#endif
        /// <summary>
        /// Close a site, if it has a site policy applied and is currently not closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was closed, false otherwise</returns>
#if !ONPREMISES || SP2019
        private static async Task<bool> SetClosedBySitePolicyImplementation(this Web web)
#else
        private static bool SetClosedBySitePolicyImplementation(this Web web)
#endif
        {
#if !ONPREMISES || SP2019
            if (await web.HasSitePolicyAppliedImplementation() && !await web.IsClosedBySitePolicyImplementation())
#else
            if (web.HasSitePolicyAppliedImplementation() && !IsClosedBySitePolicyImplementation(web))
#endif
            {
                ProjectPolicy.CloseProject(web.Context, web);
#if !ONPREMISES || SP2019
                await web.Context.ExecuteQueryRetryAsync();
#else
                web.Context.ExecuteQueryRetry();
#endif
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
#if !ONPREMISES || SP2019
            return Task.Run(() => web.SetOpenBySitePolicyImplementation()).GetAwaiter().GetResult();
#else
            return web.SetOpenBySitePolicyImplementation();
#endif
        }
#if !ONPREMISES || SP2019
        /// <summary>
        /// Open a site, if it has a site policy applied and is currently closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was opened, false otherwise</returns>
        public static async Task<bool> SetOpenBySitePolicyAsync(this Web web)
        {
            await new SynchronizationContextRemover();
            return await web.SetOpenBySitePolicyImplementation();
        }
#endif
        /// <summary>
        /// Open a site, if it has a site policy applied and is currently closed
        /// </summary>
        /// <param name="web"></param>
        /// <returns>True if site was opened, false otherwise</returns>
#if !ONPREMISES || SP2019
        private static async Task<bool> SetOpenBySitePolicyImplementation(this Web web)
#else
        private static bool SetOpenBySitePolicyImplementation(this Web web)
#endif
        {
#if !ONPREMISES || SP2019
            if (await web.HasSitePolicyAppliedImplementation() && !await web.IsClosedBySitePolicyImplementation())
#else
            if (web.HasSitePolicyAppliedImplementation() && !IsClosedBySitePolicyImplementation(web))
#endif
            {
                ProjectPolicy.OpenProject(web.Context, web);
#if !ONPREMISES || SP2019
                await web.Context.ExecuteQueryRetryAsync();
#else
                web.Context.ExecuteQueryRetry();
#endif
                return true;
            }
            return false;
        }
    }
}
