# Refreshing the object model #
After updating the schema file use below command to refresh the object model:

```Cmd
xsd -c ProvisioningSchema-2016-05.xsd /n:OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605
```

Remember to update the auto-generated code by commenting (removing) the attribute System.Xml.Serialization.XmlTypeAttribute 
for the following types:
* DataValue
* BaseFieldValue
* FieldDefault
* WebPartPageWebPart
* WikiPageWebPart
