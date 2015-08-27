using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Publishing configuration to provision
    /// </summary>
    public class Publishing : IEquatable<Publishing>
    {
        #region Private Members

        private DesignPackage _designPackage = null;
        private List<AvailableWebTemplate> _availableWebTemplates = new List<AvailableWebTemplate>();
        private List<PageLayout> _pageLayouts = new List<PageLayout>();

        #endregion

        #region Constructors

        public Publishing() { }

        public Publishing(AutoCheckRequirementsOptions autoCheckRequirements, DesignPackage designPackage = null, IEnumerable<AvailableWebTemplate> availableWebTemplates = null, IEnumerable<PageLayout> pageLayouts = null)
        {
            this.AutoCheckRequirements = autoCheckRequirements;

            if (designPackage != null)
            {
                this.DesignPackage = designPackage;
            }
            if (availableWebTemplates != null)
            {
                this._availableWebTemplates.AddRange(availableWebTemplates);
            }
            if (pageLayouts != null)
            {
                this._pageLayouts.AddRange(pageLayouts);
            }
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines a Design Package to import into the current Publishing site
        /// </summary>
        public DesignPackage DesignPackage
        {
            get { return this._designPackage; }
            set { this._designPackage = value; }
        }

        /// <summary>
        /// Defines the Available Web Templates for the current Publishing site
        /// </summary>
        public List<AvailableWebTemplate> AvailableWebTemplates
        {
            get { return this._availableWebTemplates; }
            private set { this._availableWebTemplates = value; }
        }

        /// <summary>
        /// Defines the Available Page Layouts for the current Publishing site
        /// </summary>
        public List<PageLayout> PageLayouts
        {
            get { return this._pageLayouts; }
            private set { this._pageLayouts = value; }
        }

        /// <summary>
        /// Defines how an engine should behave if the requirements for provisioning publishing capabilities are not satisfied by the target site 
        /// </summary>
        public AutoCheckRequirementsOptions AutoCheckRequirements { get; set; }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                this.AutoCheckRequirements.GetHashCode(),
                this.AvailableWebTemplates.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.DesignPackage.GetHashCode(),
                this.PageLayouts.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Publishing))
            {
                return (false);
            }
            return (Equals((Publishing)obj));
        }

        public bool Equals(Publishing other)
        {
            return (
                this.AutoCheckRequirements == other.AutoCheckRequirements &&
                this.AvailableWebTemplates.DeepEquals(other.AvailableWebTemplates) &&
                this.DesignPackage == other.DesignPackage &&
                this.PageLayouts.DeepEquals(other.PageLayouts)
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines how an engine should behave if the requirements for provisioning publishing capabilities are not satisfied by the target site 
    /// </summary>
    public enum AutoCheckRequirementsOptions
    {
        /// <summary>
        /// Instructs the engine to make the target site compliant with the requirements
        /// </summary>
        MakeCompliant,
        /// <summary>
        /// Instructs the engine to skip the Publishing section if the target site is not compliant with the requirements
        /// </summary>
        SkipIfNotCompliant,
        /// <summary>
        /// Instructs the engine to throw an exception/failure if the target site is not compliant with the requirements
        /// </summary>
        FailIfNotCompliant,
    }
}
