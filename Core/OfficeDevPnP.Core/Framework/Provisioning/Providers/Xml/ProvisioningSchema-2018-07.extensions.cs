namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201807
{

    public partial class CanvasSection
    {
        private int zoneEmphasisField;

        private bool zoneEmphasisFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int ZoneEmphasis
        {
            get
            {
                return this.zoneEmphasisField;
            }
            set
            {
                this.zoneEmphasisField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool ZoneEmphasisFieldSpecified
        {
            get
            {
                return this.zoneEmphasisFieldSpecified;
            }
            set
            {
                this.zoneEmphasisFieldSpecified = value;
            }
        }
    }

}
