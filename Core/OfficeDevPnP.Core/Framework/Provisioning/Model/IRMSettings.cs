using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Declares the Information Rights Management settings for the list or library.
    /// </summary>
    public partial class IRMSettings : BaseModel, IEquatable<IRMSettings>
    {
        #region Public Properties

        /// <summary>
        /// Defines whether the IRM settings have to be enabled or not.
        /// </summary>
        public Boolean Enabled { get; set; }

        /// <summary>
        /// Defines whether a viewer can print the downloaded document.
        /// </summary>
        public Boolean AllowPrint { get; set; }

        /// <summary>
        /// Defines whether a viewer can run a script on the downloaded document.
        /// </summary>
        public Boolean AllowScript { get; set; }

        /// <summary>
        /// Defines whether a viewer can write on a copy of the downloaded document.
        /// </summary>
        public Boolean AllowWriteCopy { get; set; }

        /// <summary>
        /// Defines whether to block Office Web Application Companion applications (WACs) from showing this document.
        /// </summary>
        public Boolean DisableDocumentBrowserView { get; set; }

        /// <summary>
        /// Defines the number of days after which the downloaded document will expire.
        /// </summary>
        public Int32 DocumentAccessExpireDays { get; set; }

        /// <summary>
        /// Defines the expire days for the Information Rights Management (IRM) protection of this document library will expire.
        /// </summary>
        public Int32 DocumentLibraryProtectionExpiresInDays { get; set; }

        /// <summary>
        /// Defines whether the downloaded document will expire.
        /// </summary>
        public Boolean EnableDocumentAccessExpire { get; set; }

        /// <summary>
        /// Defines whether to enable Office Web Application Companion applications (WACs) to publishing view.
        /// </summary>
        public Boolean EnableDocumentBrowserPublishingView { get; set; }

        /// <summary>
        /// Defines whether the permission of the downloaded document is applicable to a group.
        /// </summary>
        public Boolean EnableGroupProtection { get; set; }

        /// <summary>
        /// Defines whether a user must verify their credentials after some interval.
        /// </summary>
        public Boolean EnableLicenseCacheExpire { get; set; }

        /// <summary>
        /// Defines the group name (email address) that the permission is also applicable to.
        /// </summary>
        public String GroupName { get; set; }

        /// <summary>
        /// Defines the number of days that the application that opens the document caches the IRM license. When these elapse, the application will connect to the IRM server to validate the license.
        /// </summary>
        public Int32 LicenseCacheExpireDays { get; set; }

        /// <summary>
        /// Defines the permission policy description.
        /// </summary>
        public String PolicyDescription { get; set; }

        /// <summary>
        /// Defines the permission policy title.
        /// </summary>
        public String PolicyTitle { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}",
                    this.Enabled.GetHashCode(),
                    this.AllowPrint.GetHashCode(),
                    this.AllowScript.GetHashCode(),
                    this.AllowWriteCopy.GetHashCode(),
                    this.DisableDocumentBrowserView.GetHashCode(),
                    this.DocumentAccessExpireDays.GetHashCode(),
                    this.DocumentLibraryProtectionExpiresInDays.GetHashCode(),
                    this.EnableDocumentAccessExpire.GetHashCode(),
                    this.EnableDocumentBrowserPublishingView.GetHashCode(),
                    this.EnableGroupProtection.GetHashCode(),
                    this.EnableLicenseCacheExpire.GetHashCode(),
                    this.GroupName?.GetHashCode() ?? 0,
                    this.LicenseCacheExpireDays.GetHashCode(),
                    this.PolicyDescription?.GetHashCode() ?? 0,
                    this.PolicyTitle?.GetHashCode() ?? 0
                ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ObjectSecurity
        /// </summary>
        /// <param name="obj">Object that represents ObjectSecurity</param>
        /// <returns>true if the current object is equal to the ObjectSecurity</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is IRMSettings))
            {
                return (false);
            }
            return (Equals((IRMSettings)obj));
        }

        /// <summary>
        /// Compares IRMSettings object based on Enabled, AllowPrint, AllowScript, AllowWriteCopy, DisableDocumentBrowserView, DocumentAccessExpireDays, 
        /// DocumentLibraryProtectionExpiresInDays, EnableDocumentAccessExpire, EnableDocumentBrowserPublishingView, 
        /// EnableGroupProtection, EnableLicenseCacheExpire, GroupName, LicenseCacheExpireDays, PolicyDescription, and PolicyTitle
        /// </summary>
        /// <param name="other">IRMSettings object</param>
        /// <returns>true if the IRMSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(IRMSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                    this.Enabled == other.Enabled &&
                    this.AllowPrint == other.AllowPrint &&
                    this.AllowScript == other.AllowScript &&
                    this.AllowWriteCopy == other.AllowWriteCopy &&
                    this.DisableDocumentBrowserView == other.DisableDocumentBrowserView &&
                    this.DocumentAccessExpireDays == other.DocumentAccessExpireDays &&
                    this.DocumentLibraryProtectionExpiresInDays == other.DocumentLibraryProtectionExpiresInDays &&
                    this.EnableDocumentAccessExpire == other.EnableDocumentAccessExpire &&
                    this.EnableDocumentBrowserPublishingView == other.EnableDocumentBrowserPublishingView &&
                    this.EnableGroupProtection == other.EnableGroupProtection &&
                    this.EnableLicenseCacheExpire == other.EnableLicenseCacheExpire &&
                    this.GroupName == other.GroupName &&
                    this.LicenseCacheExpireDays == other.LicenseCacheExpireDays &&
                    this.PolicyDescription == other.PolicyDescription &&
                    this.PolicyTitle == other.PolicyTitle
                );
        }

        #endregion
    }
}
