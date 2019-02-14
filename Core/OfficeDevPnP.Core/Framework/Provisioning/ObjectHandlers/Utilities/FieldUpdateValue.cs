namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
    public class FieldUpdateValue
    {
        public string Key { get; set; }
        public object Value { get; set; }
        public string FieldTypeString { get; set; }

        public FieldUpdateValue(string key, object value)
        {
            Key = key;
            Value = value;
        }

        public FieldUpdateValue(string key, object value, string fieldTypeString)
        {
            Key = key;
            Value = value;
            FieldTypeString = fieldTypeString;
        }
    }
}