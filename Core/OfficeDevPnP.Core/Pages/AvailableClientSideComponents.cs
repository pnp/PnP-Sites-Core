namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Class holding a collection of client side webparts (retrieved via the _api/web/GetClientSideWebParts REST call)
    /// </summary>
    public class AvailableClientSideComponents
    {
        public ClientSideComponent[] value { get; set; }
    }
#endif
}
