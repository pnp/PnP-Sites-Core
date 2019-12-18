using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// The Fun Settings for the Team
    /// </summary>
    public partial class TeamFunSettings : BaseModel, IEquatable<TeamFunSettings>
    {
        #region Private Members

        private string _giphyContentRating;

        #endregion

        #region Public Members

        /// <summary>
        /// Defines whether Giphys are consented or not
        /// </summary>
        public Boolean AllowGiphy { get; set; }

        /// <summary>
        /// Defines the Content Rating for Giphys
        /// </summary>
        public string GiphyContentRating
        {
            get
            {
                return (_giphyContentRating);
            }
            set
            {
                _giphyContentRating = value?.ToLower();
            }
        }

        /// <summary>
        /// Defines whether Stickers and Memes are consented or not
        /// </summary>
        public Boolean AllowStickersAndMemes { get; set; }

        /// <summary>
        /// Defines whether Custom Memes are consented or not
        /// </summary>
        public Boolean AllowCustomMemes { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                AllowGiphy.GetHashCode(),
                GiphyContentRating.GetHashCode(),
                AllowStickersAndMemes.GetHashCode(),
                AllowCustomMemes.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamFunSettings class
        /// </summary>
        /// <param name="obj">Object that represents TeamFunSettings</param>
        /// <returns>Checks whether object is TeamFunSettings class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamFunSettings))
            {
                return (false);
            }
            return (Equals((TeamFunSettings)obj));
        }

        /// <summary>
        /// Compares TeamFunSettings object based on AllowGiphy, GiphyContentRating, AllowStickersAndMemes, and AllowCustomMemes
        /// </summary>
        /// <param name="other">TeamFunSettings Class object</param>
        /// <returns>true if the TeamFunSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamFunSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AllowGiphy == other.AllowGiphy &&
                this.GiphyContentRating == other.GiphyContentRating &&
                this.AllowStickersAndMemes == other.AllowStickersAndMemes &&
                this.AllowCustomMemes == other.AllowCustomMemes
                );
        }

        #endregion
    }

        /// <summary>
        /// Defines the Content Rating for Giphys
        /// </summary>
        public static class TeamGiphyContentRating
    {
        /// <summary>
        /// Moderate Content Rating
        /// </summary>
        public const string Moderate = "moderate";
        /// <summary>
        /// Strict Content Rating
        /// </summary>
        public const string Strict = "strict";
    }
}
