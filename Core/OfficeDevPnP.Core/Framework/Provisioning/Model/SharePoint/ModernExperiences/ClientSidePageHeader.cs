using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Represents the Header of a Client Side page
    /// </summary>
    public partial class ClientSidePageHeader : BaseModel, IEquatable<ClientSidePageHeader>
    {
        #region Public Members

        /// <summary>
        /// Defines the type of the header for the client side page
        /// </summary>
        public ClientSidePageHeaderType Type { get; set; }

        /// <summary>
        /// Defines the server-relative URL of the image for the header of the current client side page
        /// </summary>
        public String ServerRelativeImageUrl { get; set; }

        /// <summary>
        /// Defines the x-translate of the image for the header of the current client side page.
        /// </summary>
        public Double? TranslateX { get; set; }

        /// <summary>
        /// Defines the y-translate of the image for the header of the current client side page.
        /// </summary>
        public Double? TranslateY { get; set; }

        /// <summary>
        /// Defines the type of layout used inside the header of the current client side page
        /// </summary>
        public ClientSidePageHeaderLayoutType LayoutType { get; set; }

        /// <summary>
        /// Defines the text alignment of the text in the header of the current client side page
        /// </summary>
        public ClientSidePageHeaderTextAlignment TextAlignment { get; set; }

        /// <summary>
        /// Defines whether to show the topic header in the title region of the current client side page
        /// </summary>
        public Boolean ShowTopicHeader { get; set; }

        /// <summary>
        /// Defines whether to show the page publication date in the title region of the current client side page
        /// </summary>
        public Boolean ShowPublishDate { get; set; }

        /// <summary>
        /// Defines the topic header text to show if ShowTopicHeader is set to true of the current client side page
        /// </summary>
        public String TopicHeader { get; set; }

        /// <summary>
        /// Defines the alternative text for the header image of the current client side page
        /// </summary>
        public String AlternativeText { get; set; }

        /// <summary>
        /// Defines the page author(s) to be displayed of the current client side page
        /// </summary>
        public String Authors { get; set; }

        /// <summary>
        /// Defines the page author by line of the current client side page
        /// </summary>
        public String AuthorByLine { get; set; }

        /// <summary>
        /// Defines the ID of the page author by line of the current client side page
        /// </summary>
        public Int32 AuthorByLineId { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|",
                this.Type.GetHashCode(),
                this.ServerRelativeImageUrl?.GetHashCode() ?? 0,
                this.TranslateX.GetHashCode(),
                this.TranslateY.GetHashCode(),
                this.LayoutType.GetHashCode(),
                this.TextAlignment.GetHashCode(),
                this.ShowTopicHeader.GetHashCode(),
                this.ShowPublishDate.GetHashCode(),
                this.TopicHeader?.GetHashCode() ?? 0,
                this.AlternativeText?.GetHashCode() ?? 0,
                this.Authors?.GetHashCode() ?? 0,
                this.AuthorByLine?.GetHashCode() ?? 0,
                this.AuthorByLineId.GetHashCode()
           ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ClientSidePageHeader class
        /// </summary>
        /// <param name="obj">Object that represents ClientSidePageHeader</param>
        /// <returns>Checks whether object is ClientSidePageHeader class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ClientSidePageHeader))
            {
                return (false);
            }
            return (Equals((ClientSidePageHeader)obj));
        }

        /// <summary>
        /// Compares ClientSidePageHeader object based on Type, ServerRelativeImageUrl, TranslateX, TranslateY, 
        /// Layout, TextAlignment, TopicHeader, AlternativeText, Authors, AuthorByLine, and AuthorByLineId
        /// </summary>
        /// <param name="other">ClientSidePageHeader Class object</param>
        /// <returns>true if the ClientSidePageHeader object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ClientSidePageHeader other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Type == other.Type &&
                this.ServerRelativeImageUrl == other.ServerRelativeImageUrl &&
                this.TranslateX == other.TranslateX &&
                this.TranslateY == other.TranslateY &&
                this.LayoutType == other.LayoutType &&
                this.TextAlignment == other.TextAlignment &&
                this.ShowTopicHeader == other.ShowTopicHeader &&
                this.ShowPublishDate == other.ShowPublishDate &&
                this.TopicHeader == other.TopicHeader &&
                this.AlternativeText == other.AlternativeText &&
                this.Authors == other.Authors &&
                this.AuthorByLine == other.AuthorByLine &&
                this.AuthorByLineId == other.AuthorByLineId
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the possible values for the client side page header type
    /// </summary>
    public enum ClientSidePageHeaderType
    {
        /// <summary>
        /// The client side page does not hava any header
        /// </summary>
        None,
        /// <summary>
        /// Default client side page header
        /// </summary>
        Default,
        /// <summary>
        /// The client side page has a custom header
        /// </summary>
        Custom,
    }

    /// <summary>
    /// Defines the type of layout used inside the header of the current client side page
    /// </summary>
    public enum ClientSidePageHeaderLayoutType
    {
        /// <summary>
        /// Full Width Image
        /// </summary>
        FullWidthImage,
        /// <summary>
        /// No Image
        /// </summary>
        NoImage,
        /// <summary>
        /// Color Block
        /// </summary>
        ColorBlock,
        /// <summary>
        /// Cut In Shape
        /// </summary>
        CutInShape,
    }

    /// <summary>
    /// Defines the text alignment of the text in the header of the current client side page
    /// </summary>
    public enum ClientSidePageHeaderTextAlignment
    {
        /// <summary>
        /// Align Left
        /// </summary>
        Left,
        /// <summary>
        /// Align Center
        /// </summary>
        Center,
    }
}
