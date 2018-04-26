using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a CanvasControl
    /// </summary>
    public partial class CanvasControl : BaseModel, IEquatable<CanvasControl>
    {
        #region Public Members

        /// <summary>
        /// Defines the custom properties for the client-side web part control.
        /// </summary>
        public Dictionary<String, String> ControlProperties { get; set; }

        /// <summary>
        /// Defines the Type of Client-side Web Part.
        /// </summary>
        public WebPartType Type { get; set; }

        /// <summary>
        /// Defines the Name of the client-side web part if the WebPartType attribute has a value of "Custom".
        /// </summary>
        public String CustomWebPartName { get; set; }

        /// <summary>
        /// Defines the JSON Control Data for Canvas Control of a Client-side Page.
        /// </summary>
        public String JsonControlData { get; set; }

        /// <summary>
        /// Defines the Instance Id for Canvas Control of a Client-side Page.
        /// </summary>
        public Guid ControlId { get; set; }

        /// <summary>
        /// Defines the order of the Canvas Control for a Client-side Page.
        /// </summary>
        public Int32 Order { get; set; }

        /// <summary>
        /// Defines the column of the section in which the Canvas Control will be inserted. Optional, default 0.
        /// </summary>
        public Int32 Column { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|",
                this.ControlProperties.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                Type.GetHashCode(),
                CustomWebPartName?.GetHashCode() ?? 0,
                JsonControlData?.GetHashCode() ?? 0,
                ControlId.GetHashCode(),
                Order.GetHashCode(),
                Column.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with CanvasControl class
        /// </summary>
        /// <param name="obj">Object that represents CanvasControl</param>
        /// <returns>Checks whether object is CanvasControl class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is CanvasControl))
            {
                return (false);
            }
            return (Equals((CanvasControl)obj));
        }

        /// <summary>
        /// Compares CanvasControl object based on Controls, Order, and Type
        /// </summary>
        /// <param name="other">CanvasControl Class object</param>
        /// <returns>true if the CanvasControl object is equal to the current object; otherwise, false.</returns>
        public bool Equals(CanvasControl other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.ControlProperties.DeepEquals(other.ControlProperties) &&
                this.Type == other.Type &&
                this.CustomWebPartName == other.CustomWebPartName &&
                this.JsonControlData == other.JsonControlData &&
                this.ControlId == other.ControlId &&
                this.Order == other.Order &&
                this.Column == other.Column
                );
        }

        #endregion
    }

    public enum WebPartType
    {
        Custom,
        Text,
        ContentRollup,
        BingMap,
        ContentEmbed,
        DocumentEmbed,
        Image,
        ImageGallery,
        LinkPreview,
        NewsFeed,
        NewsReel,
        PowerBIReportEmbed,
        QuickChart,
        SiteActivity,
        VideoEmbed,
        YammerEmbed,
        Events,
        GroupCalendar,
        Hero,
        List,
        PageTitle,
        People,
        QuickLinks,
        CustomMessageRegion,
        Divider,
        MicrosoftForms,
        Spacer,
        ClientWebPart
    }
}
