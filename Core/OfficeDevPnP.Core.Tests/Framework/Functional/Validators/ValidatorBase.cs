using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{

    /// <summary>
    /// Base object validator class
    /// </summary>
    public class ValidatorBase
    {
        #region Validation methods
        /// <summary>
        /// Validate two collection objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sourceElement"></param>
        /// <param name="targetElement"></param>
        /// <param name="props"></param>
        /// <returns></returns>
        public static bool ValidateObjects<T>(T sourceElement, T targetElement, List<string> property) where T : class
        {
            IEnumerable sElements = (IEnumerable)sourceElement;
            IEnumerable tElements = (IEnumerable)targetElement;
            int sCount = 0;
            int tCount = 0;

            foreach (string p in property)
            {
                foreach (object sElem in sElements)
                {
                    sCount++;
                    object sValue = sElem.GetType().GetProperty(p).GetValue(sElem);

                    foreach (object tElem in tElements)
                    {
                        object tValue = tElem.GetType().GetProperty(p).GetValue(tElem);

                        if (Convert.ToString(sValue) == Convert.ToString(tValue))
                        {
                            tCount++;
                            break;
                        }
                    }
                }
            }

            if (sCount != tCount)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sourceParser"></param>
        /// <param name="targetParser"></param>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="property"></param>
        /// <returns></returns>
        public static bool ValidateObjectSchemaXML<T>(TokenParser sourceParser, TokenParser targetParser, IEnumerable<T> source, IEnumerable<T> target, string property) where T : class
        {
            int scount = 0;
            int tcount = 0;

            foreach (var sField in source)
            {
                object sSchemaXml = sField.GetType().GetProperty("SchemaXml").GetValue(sField);
                XElement sourceElement = XElement.Parse(sourceParser.ParseString(sSchemaXml.ToString(), "~sitecollection", "~site"));
                var sValue = sourceElement.Attribute(property).Value;
                scount++;

                foreach (var tField in target)
                {
                    object tSchemaXml = sField.GetType().GetProperty("SchemaXml").GetValue(sField);
                    XElement targetElement = XElement.Parse(targetParser.ParseString(tSchemaXml.ToString(), "~sitecollection", "~site"));
                    var tValue = targetElement.Attribute(property).Value;

                    if (sValue == tValue)
                    {
                        tcount++;
                        break;
                    }
                }
            }

            if (scount != tcount)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="context"></param>
        /// <param name="security"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public static bool ValidateSecurity(ClientContext context, ObjectSecurity security, SecurableObject item)
        {
            int dataRowRoleAssignmentCount = security.RoleAssignments.Count;
            int roleCount = 0;

            IEnumerable roles = context.LoadQuery(item.RoleAssignments.Include(roleAsg => roleAsg.Member,
                roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name)));
            context.ExecuteQuery();

            foreach (var s in security.RoleAssignments)
            {
                foreach (Microsoft.SharePoint.Client.RoleAssignment r in roles)
                {
                    if (r.Member.LoginName.Contains(s.Principal) && r.RoleDefinitionBindings.Where(i => i.Name == s.RoleDefinition).FirstOrDefault() != null)
                    {
                        roleCount++;
                    }
                }
            }

            if (dataRowRoleAssignmentCount != roleCount)
            {
                return false;
            }

            return true;
        }
        #endregion

    }
}
