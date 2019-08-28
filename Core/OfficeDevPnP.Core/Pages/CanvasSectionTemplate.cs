namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
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
#if !SP2019
        /// <summary>
        /// One column + one vertical section column
        /// </summary>
        OneColumnVerticalSection = 6,
        /// <summary>
        /// Two columns of the same size + one vertical section column
        /// </summary>
        TwoColumnVerticalSection = 7,
        /// <summary>
        /// Three columns of the size + one vertical section column
        /// </summary>
        ThreeColumnVerticalSection = 8,
        /// <summary>
        /// Two columns, left 2/3, right 1/3 + one vertical section column
        /// </summary>
        TwoColumnLeftVerticalSection = 9,
        /// <summary>
        /// Two columns, left 1/3, right 2/3 + one vertical section column
        /// </summary>
        TwoColumnRightVerticalSection = 10
#endif
    }
#endif
}
