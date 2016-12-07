using Microsoft.SharePoint.Client;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class ListViewIdToken : TokenDefinition
    {
        private string _listTitle = null;
        private string _viewTitle = null;

        public ListViewIdToken(Web web, string listTitle, string viewTitle)
            : base(web, string.Format("{{viewid:{0},{1}}}",
                Regex.Escape(listTitle),
                Regex.Escape(viewTitle)))
        {
            _listTitle = listTitle;
            _viewTitle = viewTitle;
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    var list = this.Web.Lists.GetByTitle(_listTitle);
                    var view = list.Views.GetByTitle(_viewTitle);
                    context.Load(view, v => v.Id);
                    context.ExecuteQueryRetry();
                    CacheValue = view.Id.ToString();
                }
            }
            return CacheValue;
        }
    }
}