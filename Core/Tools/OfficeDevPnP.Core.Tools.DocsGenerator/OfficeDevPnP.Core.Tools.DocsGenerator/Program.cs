using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Reflection;

namespace OfficeDevPnP.Core.Tools.DocsGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            GenerateMDFromPnPSchema();
            // GenerateMDFromTokenDefinitions();
        }

        static void GenerateMDFromPnPSchema()
        {
            XDocument xsd = XDocument.Load(@"..\..\..\..\..\OfficeDevPnP.Core\Framework\Provisioning\Providers\Xml\ProvisioningSchema-2020-02.xsd");
            XslCompiledTransform xslt = new XslCompiledTransform();
            xslt.Load(@"..\..\XSD2MD.xslt");

            XsltArgumentList xsltArgs = new XsltArgumentList();
            xsltArgs.AddParam("now", String.Empty, DateTime.Now.ToShortDateString());

            using (FileStream fs = new FileStream(@"..\..\..\..\..\ProvisioningSchema-2020-02.md", FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                xslt.Transform(xsd.CreateNavigator(), xsltArgs, fs);
            }
        }

        //static void GenerateMDFromTokenDefinitions()
        //{
        //    var path = @"..\..\..\..\..\ProvisioningEngineTokens.md";
        //    var assembly = Assembly.GetAssembly(typeof(OfficeDevPnP.Core.ALM.AppManager));
        //    var analyzer = new ParameterAnalyzer(assembly);
        //    var parameters = analyzer.Analyze();

        //    var builder = new StringBuilder();
        //    builder.Append($"Office 365 Developer PnP Core Component Provisioning Engine Tokens{Environment.NewLine}");
        //    builder.Append($"=================================================================={Environment.NewLine}{Environment.NewLine}");
        //    builder.Append($"### Summary ###{Environment.NewLine}");
        //    builder.Append($"The SharePoint PnP Core Provisioning Engine supports certain tokens which will be replaced by corresponding values during provisioning. These tokens can be used to make the template site collection independent for instance.{Environment.NewLine}{Environment.NewLine}");
        //    builder.Append($"Below all the supported tokens are listed:{Environment.NewLine}{Environment.NewLine}");
        //    builder.Append($"Token|Description|Example|Returns{Environment.NewLine}");
        //    builder.Append($":-----|:----------|:------|:------{Environment.NewLine}");
        //    foreach (var parameter in parameters.OrderBy(t => t.Token))
        //    {
        //        builder.Append($"{parameter.Token}|{parameter.Description}|{parameter.Example}|{parameter.Returns}{Environment.NewLine}");
        //    }
        //    System.IO.File.WriteAllText(path, builder.ToString());
        //}
    }
}
