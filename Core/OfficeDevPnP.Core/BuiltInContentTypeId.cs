using System.Collections.Generic;

namespace OfficeDevPnP.Core
{
    /// <summary>
    /// A class that returns strings that represent identifiers (IDs) for built-in content types.
    /// </summary>
    public static class BuiltInContentTypeId
    {
        public const string AdminTask = "0x010802";
        public const string Announcement = "0x0104";
        public const string BasicPage = "0x010109";
        public const string BlogComment = "0x0111";
        public const string BlogPost = "0x0110";
        public const string CallTracking = "0x0100807FBAC5EB8A4653B8D24775195B5463";
        public const string Contact = "0x0106";
        public const string Discussion = "0x012002";
        public const string DisplayTemplateJS = "0x0101002039C03B61C64EC4A04F5361F3851068";
        public const string Document = "0x0101";

        /// <summary>
        /// Contains the content identifier (ID) for the DocumentSet content type. To get content type from a list, use BestMatchContentTypeId().
        /// </summary>
        public const string DocumentSet = "0x0120D520";

        public const string DocumentWorkflowItem = "0x010107";
        public const string DomainGroup = "0x010C";
        public const string DublinCoreName = "0x01010B";
        public const string Event = "0x0102";
        public const string FarEastContact = "0x0116";
        public const string Folder = "0x0120";
        public const string GbwCirculationCTName = "0x01000F389E14C9CE4CE486270B9D4713A5D6";
        public const string GbwOfficeNoticeCTName = "0x01007CE30DD1206047728BAFD1C39A850120";
        public const string HealthReport = "0x0100F95DB3A97E8046B58C6A54FB31F2BD46";
        public const string HealthRuleDefinition = "0x01003A8AA7A4F53046158C5ABD98036A01D5";
        public const string Holiday = "0x01009BE2AB5291BF4C1A986910BD278E4F18";
        public const string IMEDictionaryItem = "0x010018F21907ED4E401CB4F14422ABC65304";
        public const string Issue = "0x0103";

        /// <summary>
        /// Contains the content identifier (ID) for the Item content type.
        /// </summary>
        public const string Item = "0x01";

        public const string Link = "0x0105";
        public const string LinkToDocument = "0x01010A";
        public const string MasterPage = "0x010105";
        public const string Message = "0x0107";
        public const string ODCDocument = "0x010100629D00608F814DD6AC8A86903AEE72AA";
        public const string Person = "0x010A";
        public const string Picture = "0x010102";
        public const string Resource = "0x01004C9F4486FBF54864A7B0A33D02AD19B1";
        public const string ResourceGroup = "0x0100CA13F2F8D61541B180952DFB25E3E8E4";
        public const string ResourceReservation = "0x0102004F51EFDEA49C49668EF9C6744C8CF87D";
        public const string RootOfList = "0x012001";
        public const string Schedule = "0x0102007DBDC1392EAF4EBBBF99E41D8922B264";
        public const string ScheduleAndResourceReservation = "0x01020072BB2A38F0DB49C3A96CF4FA85529956";
        public const string SharePointGroup = "0x010B";
        public const string SummaryTask = "0x012004";
        public const string System = "0x";
        public const string Task = "0x0108";
        public const string Timecard = "0x0100C30DDA8EDB2E434EA22D793D9EE42058";
        public const string UDCDocument = "0x010100B4CBD48E029A4AD8B62CB0E41868F2B0";
        public const string UntypedDocument = "0x010104";
        public const string WebPartPage = "0x01010901";
        public const string WhatsNew = "0x0100A2CA87FF01B442AD93F37CD7DD0943EB";
        public const string Whereabouts = "0x0100FBEEE6F0C500489B99CDA6BB16C398F7";
        public const string WikiDocument = "0x010108";
        public const string WorkflowHistory = "0x0109";
        public const string WorkflowTask = "0x010801";
        public const string Workflow2013Task = "0x0108003365C4474CAE8C42BCE396314E88E51F";
        public const string XMLDocument = "0x010101";
        public const string XSLStyle = "0x010100734778F2B7DF462491FC91844AE431CF";

        /// <summary>
        /// Contains the content identifier (ID) for content types used in the publishing infrastructure.
        /// </summary>
        public const string ReusableHTML = "0x01002CF74A4DAE39480396EEA7A4BA2BE5FB";
        public const string PublishedLink = "0x01004613D6562E4C41A7B9DADDAC1689E00D";
        public const string ReusableText = "0x01004D5A79BAFA4A4576B79C56FF3D0D662D";
        public const string PageOutputCache = "0x010087D89D279834C94E98E5E1B4A913C67E";
        public const string DeviceChannel = "0x01009AF87C5C1DF34CA38277FEABCB5018F6";
        public const string DeviceChannelMappings = "0x010100FDA260FD09A244B183A666F2AE2475A6";
        public const string SystemPageLayout = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC1";
        public const string PageLayout = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE811";
        public const string HtmlPageLayout = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE8110003D357F861E29844953D5CAA1D4D8A3B";
        public const string SystemMasterPage = "0x0101000F1C8B9E0EB4BE489F09807B2C53288F";
        public const string ASPNETMasterPage = "0x0101000F1C8B9E0EB4BE489F09807B2C53288F0054AD6EF48B9F7B45A142F8173F171BD1";
        public const string HtmlMasterPage = "0x0101000F1C8B9E0EB4BE489F09807B2C53288F0054AD6EF48B9F7B45A142F8173F171BD10003D357F861E29844953D5CAA1D4D8A3A";
        public const string SystemPage = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2";
        public const string Page = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF39";
        public const string ArticlePage = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900242457EFB8B24247815D688C526CD44D";
        public const string EnterpriseWikiPage = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF39004C1F8B46085B4D22B1CDC3DE08CFFB9C";
        public const string ProjectPage = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF39004C1F8B46085B4D22B1CDC3DE08CFFB9C0055EF50AAFF2E4BADA437E4BAE09A30F8";
        public const string WelcomePage = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF390064DEA0F50FC8C147B0B6EA0636C4A7D4";
        public const string ErrorPage = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900796F542FC5E446758C697981E370458C";
        public const string RedirectPage = "0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900FD0E870BA06948879DBD5F9813CD8799";

        /// <summary>
        /// Contains the content identifier (ID) for content types used in the modern page infrastructure
        /// </summary>
        public const string ModernArticlePage = "0x0101009D1CB255DA76424F860D91F20E6C4118";
        public const string RepostPage = "0x0101009D1CB255DA76424F860D91F20E6C4118002A50BFCFB7614729B56886FADA02339B";


        private static Dictionary<string, bool> s_dict = (Dictionary<string, bool>) null;

        public static bool Contains(string id)
        {
            if (s_dict == null)
            {
                s_dict = new Dictionary<string, bool>();
                s_dict.Add(AdminTask, true);
                s_dict.Add(Announcement, true);
                s_dict.Add(BasicPage, true);
                s_dict.Add(BlogComment, true);
                s_dict.Add(CallTracking, true);
                s_dict.Add(Contact, true);
                s_dict.Add(Discussion, true);
                s_dict.Add(DisplayTemplateJS, true);
                s_dict.Add(Document, true);
                s_dict.Add(DocumentSet, true);
                s_dict.Add(DocumentWorkflowItem, true);
                s_dict.Add(DomainGroup, true);
                s_dict.Add(DublinCoreName, true);
                s_dict.Add(Event, true);
                s_dict.Add(FarEastContact, true);
                s_dict.Add(Folder, true);
                s_dict.Add(GbwCirculationCTName, true);
                s_dict.Add(GbwOfficeNoticeCTName, true);
                s_dict.Add(HealthReport, true);
                s_dict.Add(HealthRuleDefinition, true);
                s_dict.Add(Holiday, true);
                s_dict.Add(IMEDictionaryItem, true);
                s_dict.Add(Issue, true);
                s_dict.Add(Item, true);
                s_dict.Add(Link, true);
                s_dict.Add(LinkToDocument, true);
                s_dict.Add(MasterPage, true);
                s_dict.Add(Message, true);
                s_dict.Add(ODCDocument, true);
                s_dict.Add(Person, true);
                s_dict.Add(Picture, true);
                s_dict.Add(Resource, true);
                s_dict.Add(ResourceGroup, true);
                s_dict.Add(ResourceReservation, true);
                s_dict.Add(RootOfList, true);
                s_dict.Add(Schedule, true);
                s_dict.Add(ScheduleAndResourceReservation, true);
                s_dict.Add(SharePointGroup, true);
                s_dict.Add(SummaryTask, true);
                s_dict.Add(System, true);
                s_dict.Add(Task, true);
                s_dict.Add(Timecard, true);
                s_dict.Add(UDCDocument, true);
                s_dict.Add(UntypedDocument, true);
                s_dict.Add(WebPartPage, true);
                s_dict.Add(WhatsNew, true);
                s_dict.Add(Whereabouts, true);
                s_dict.Add(WikiDocument, true);
                s_dict.Add(WorkflowHistory, true);
                s_dict.Add(WorkflowTask, true);
                s_dict.Add(XMLDocument, true);
                s_dict.Add(XSLStyle, true);
                s_dict.Add(ReusableHTML, true);
                s_dict.Add(PublishedLink, true);
                s_dict.Add(ReusableText, true);
                s_dict.Add(PageOutputCache, true);
                s_dict.Add(DeviceChannel, true);
                s_dict.Add(DeviceChannelMappings, true);
                s_dict.Add(SystemPageLayout, true);
                s_dict.Add(PageLayout, true);
                s_dict.Add(HtmlPageLayout, true);
                s_dict.Add(SystemMasterPage, true);
                s_dict.Add(ASPNETMasterPage, true);
                s_dict.Add(HtmlMasterPage, true);
                s_dict.Add(SystemPage, true);
                s_dict.Add(Page, true);
                s_dict.Add(ArticlePage, true);
                s_dict.Add(EnterpriseWikiPage, true);
                s_dict.Add(ProjectPage, true);
                s_dict.Add(WelcomePage, true);
                s_dict.Add(ErrorPage, true);
                s_dict.Add(RedirectPage, true);
                s_dict.Add(ModernArticlePage, true);
            }
            bool flag = false;
            s_dict.TryGetValue(id, out flag);
            return flag;
        }
    }
}
