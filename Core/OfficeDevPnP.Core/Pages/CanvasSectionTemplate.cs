namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// The type of canvas being used
    /// </summary>
    public enum CanvasSectionTemplate
    {
        /// <summary>
        /// One column
        /// </summary>
        OneColumn = 0,
        /// <summary>
        /// One column, full browser width. This one only works for communication sites in combination with image or hero webparts
        /// </summary>
        OneColumnFullWidth = 1,
        /// <summary>
        /// Two columns of the same size
        /// </summary>
        TwoColumn = 2,
        /// <summary>
        /// Three columns of the same size
        /// </summary>
        ThreeColumn = 3,
        /// <summary>
        /// Two columns, left one is 2/3, right one 1/3
        /// </summary>
        TwoColumnLeft = 4,
        /// <summary>
        /// Two columns, left one is 1/3, right one 2/3
        /// </summary>
        TwoColumnRight = 5,

    }
#endif
}
