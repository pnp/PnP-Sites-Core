using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    #region Classes to support json (de-)serialization of control/webpart data
    #region Control data

    /// <summary>
    /// Base class representing the json control data that will describe a control versus the zones and sections on a page
    /// </summary>
    public class ClientSideCanvasPosition
    {
        [JsonProperty(PropertyName = "zoneIndex")]
        public float ZoneIndex { get; set; }
        [JsonProperty(PropertyName = "sectionIndex")]
        public int SectionIndex { get; set; }
        [JsonProperty(PropertyName = "sectionFactor")]
        public int SectionFactor { get; set; }
    }

    /// <summary>
    /// Class representing the json control data that will describe a control versus the zones and sections on a page
    /// </summary>
    public class ClientSideCanvasControlPosition: ClientSideCanvasPosition
    {
        [JsonProperty(PropertyName = "controlIndex")]
        public int ControlIndex { get; set; }
    }

    /// <summary>
    /// Base class representing the json control data that will be included in each client side control (de-)serialization (data-sp-controldata attribute)
    /// </summary>
    public class ClientSideCanvasData
    {
        [JsonProperty(PropertyName = "position")]
        public ClientSideCanvasPosition Position { get; set; }
    }

    /// <summary>
    /// Base class representing the json control data that will be included in each client side control (de-)serialization (data-sp-controldata attribute)
    /// </summary>
    public class ClientSideCanvasControlData
    {
        [JsonProperty(PropertyName = "controlType")]
        public int ControlType { get; set; }

        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "position")]
        public ClientSideCanvasControlPosition Position { get; set; }
    }

    /// <summary>
    /// Control data for controls of type 4 (= text control)
    /// </summary>
    public class ClientSideTextControlData : ClientSideCanvasControlData
    {
        [JsonProperty(PropertyName = "editorType")]
        public string EditorType { get; set; }
    }

    /// <summary>
    /// Control data for controls of type 3 (= client side web parts)
    /// </summary>
    public class ClientSideWebPartControlData : ClientSideCanvasControlData
    {
        [JsonProperty(PropertyName = "webPartId")]
        public string WebPartId { get; set; }
    }
    #endregion

    #region WebPart data
    /// <summary>
    /// Json web part data that will be included in each client side web part (de-)serialization (data-sp-webpartdata)
    /// </summary>
    public class ClientSideWebPartData
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "instanceId")]
        public string InstanceId { get; set; }

        [JsonProperty(PropertyName = "title")]
        public string Title { get; set; }

        [JsonProperty(PropertyName = "description")]
        public string Description { get; set; }

        [JsonProperty(PropertyName = "dataVersion")]
        public string DataVersion { get; set; }

        [JsonProperty(PropertyName = "properties")]
        public string Properties { get; set; }
    }
    #endregion
    #endregion
#endif
}
