using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeDevPnP.Core.Tools.DocsGenerator
{
    public class ParameterAnalyzer
    {
        Assembly assembly;
        public ParameterAnalyzer(Assembly assemblyToAnalyze)
        {
            assembly = assemblyToAnalyze;
        }

        public List<EngineParameter> Analyze()
        {
            List<EngineParameter> parameters = new List<EngineParameter>();

            foreach (var type in assembly.GetTypes())
            {
                var attributes = type.GetCustomAttributes<OfficeDevPnP.Core.Attributes.TokenDefinitionDescriptionAttribute>();
                if (attributes.Any())
                {
                    foreach (var attribute in attributes)
                    {
                        var parameter = new EngineParameter();
                        parameter.Description = attribute.Description;
                        parameter.Token = attribute.Token;
                        parameter.Example = attribute.Example;
                        parameter.Returns = attribute.Returns;

                        parameters.Add(parameter);
                    }
                }
            }

            return parameters;
        }
    }
}
