using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// Holds yammer operation methods
    /// </summary>
    public static class YammerUtility
    {
        /// <summary>
        /// Returns Yammer Group if group exists. If the group does not exist, returns null.
        /// </summary>
        /// <param name="groupName">Group name to search for</param>
        /// <param name="accessToken">accessToken will have all the required permissions to update or retrieve data to Yammer on behalf of the user</param>
        /// <returns>Returns Yammer Group is group exists. If group does not exists, returns null.</returns>
        public static YammerGroup GetYammerGroupByName(string groupName, string accessToken)
        {
            YammerGroup yamGroup = null;
            var groups = GetYammerGroups(accessToken);

            foreach (var item in groups)
            {
                if (item.full_name.Equals(groupName, StringComparison.CurrentCultureIgnoreCase))
                {
                    yamGroup = item;
                }
            }
            return yamGroup;
        }

        /// <summary>
        /// Returns Yammer Group if group exists. If the group does not exist, returns null.
        /// </summary>
        /// <param name="groupId">Group Id to search for</param>
        /// <param name="accessToken">accessToken will have all the required permissions to update or retrieve data to Yammer on behalf of the user</param>
        /// <returns>Returns Yammer Group is group exists. If group does not exists, returns null.</returns>
        public static YammerGroup GetYammerGroupById(int groupId, string accessToken)
        {
            YammerGroup yamGroup = null;
            var groups = GetYammerGroups(accessToken);

            foreach (var item in groups)
            {
                if (item.id == groupId)
                {
                    yamGroup = item;
                }
            }
            return yamGroup;
        }

        /// <summary>
        /// Returns Yammer groups based on the access token. All groups are returned for registered apps.
        /// </summary>
        /// <param name="accessToken">accessToken will have all the required permissions to update or retrieve data to Yammer on behalf of the user</param>
        /// <returns>All groups in the network</returns>
        public static List<YammerGroup> GetYammerGroups(string accessToken)
        {
            // Get user
            YammerUser user = GetYammerUser(accessToken);

            //get the users groups to check for the group
            var response = GetYammerJson($"https://www.yammer.com/api/v1/groups/for_user/{user.id}.json", accessToken);
            List<YammerGroup> groups = JsonUtility.Deserialize<List<YammerGroup>>(response);

            // Updated network information to the group data
            foreach (var item in groups)
            {
                item.network_id = user.network_id;
                item.network_name = user.network_name;
            }

            return groups;
        }

        /// <summary>
        /// Returns yammer user if user exists. If the user doesn't exist, returns null
        /// </summary>
        /// <param name="accessToken">accessToken will have all the required permissions to update or retrieve data to Yammer on behalf of the user</param>
        /// <returns></returns>
        public static YammerUser GetYammerUser(string accessToken)
        {
            //get service account token
            var response = GetYammerJson($"https://www.yammer.com/api/v1/users/current.json?access_token={accessToken}", accessToken);
            return JsonUtility.Deserialize<YammerUser>(response);
        }

        /// <summary>
        /// Can be used to create Yammer group to the Yammer network
        /// </summary>
        /// <param name="groupName">Creates yammer group with given name</param>
        /// <param name="isPrivate">Sets yammer groups as private if the value is true. Otherwise sets as public group</param>
        /// <param name="accessToken">accessToken will have all the required permissions to update or retrieve data to Yammer on behalf of the user</param>
        /// <returns>Returns YammerGroup created</returns>
        public static YammerGroup CreateYammerGroup(string groupName, bool isPrivate, string accessToken)
        {
            //Get reference existing group if exists
            YammerGroup yamGroup = GetYammerGroupByName(groupName, accessToken);
            if (yamGroup == null)
            {
                //Create yammer group
                string url = $"https://www.yammer.com/api/v1/groups.json?name={groupName}&private={isPrivate.ToString().ToLower()}";
                PostYammerJson(url, accessToken);
                yamGroup = GetYammerGroupByName(groupName, accessToken);
            }
            return yamGroup;
        }

        /// <summary>
        /// Creates web part entity with the Yammer group structure on it
        /// </summary>
        /// <param name="yammerNetworkName">Yammer group network name</param>
        /// <param name="yammerGroupId">Yammer group id</param>
        /// <param name="showHeader">Shows header of webpart based on the value</param>
        /// <param name="showFooter">Shows footer of webpart based on value</param>
        /// <returns>Returns created WebPartEntity</returns>
        public static WebPartEntity GetYammerGroupDiscussionPart(string yammerNetworkName, int yammerGroupId, bool showHeader, bool showFooter)
        {
            WebPartEntity wpYammer = new WebPartEntity();
            wpYammer.WebPartXml = CreateYammerGroupDiscussionPartXml(yammerNetworkName, yammerGroupId, showHeader, showFooter);
            wpYammer.WebPartIndex = 0;
            wpYammer.WebPartTitle = "Yammer";
            return wpYammer;
        }

        /// <summary>
        /// Creates web part entity with the Yammer OpenGraph structure on it for specific URL
        /// </summary>
        /// <param name="yammerNetworkName">Yammer network name</param>
        /// <param name="url">yammer URL</param>
        /// <param name="showHeader">Shows header based on value</param>
        /// <param name="showFooter">Shows footer based on value</param>
        /// <param name="postTitle">creates webpart with given post title</param>
        /// <param name="postImageUrl">creates webpart with given post title</param>
        /// <param name="defaultGroupId">creates webpart with givne group id</param>
        /// <returns>Returns created WebPartEntity</returns>
        public static WebPartEntity GetYammerOpenGraphDiscussionPart(string yammerNetworkName, string url, bool showHeader, bool showFooter, string postTitle = "", string postImageUrl = "", string defaultGroupId = "")
        {
            WebPartEntity wpYammer = new WebPartEntity();
            wpYammer.WebPartXml = CreateYammerOpenGraphDiscussionPartXml(yammerNetworkName, url, showHeader, showFooter, postTitle, postImageUrl, true, defaultGroupId);
            wpYammer.WebPartIndex = 0;
            wpYammer.WebPartTitle = "Yammer";
            return wpYammer;
        }

        /// <summary>
        /// Constructs the webpart XML for yammer group needed to inject as Yammer web part to SharePoint page
        /// </summary>
        /// <param name="yammerNetworkName">Yammer network name</param>
        /// <param name="yammerGroupId">Group ID</param>
        /// <param name="showHeader">Shows header based on value</param>
        /// <param name="showFooter">Shows footer based on value</param>
        /// <param name="useSSO">uses given Single Sign-on information to create webpart</param>
        /// <returns>The constructed web part XML</returns>
        public static string CreateYammerDiscussionPartXml(string yammerNetworkName, int yammerGroupId, bool showHeader, bool showFooter, bool useSSO = true)
        {
            return CreateYammerGroupDiscussionPartXml(yammerNetworkName, yammerGroupId, showHeader, showFooter, useSSO);
        }

        /// <summary>
        /// Constructs the webpart XML for yammer group needed to inject as Yammer web part to SharePoint page
        /// </summary>
        /// <param name="yammerNetworkName">Yammer network name</param>
        /// <param name="yammerGroupId">Group ID</param>
        /// <param name="showHeader">Shows header based on value</param>
        /// <param name="showFooter">Shows footer based on value</param>
        /// <param name="useSSO">uses given Single Sign-on information to create webpart</param>
        /// <returns>The constructed web part XML</returns>
        public static string CreateYammerGroupDiscussionPartXml(string yammerNetworkName, int yammerGroupId, bool showHeader, bool showFooter, bool useSSO = true)
        {
            StringBuilder wp = new StringBuilder(100);
            wp.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            wp.Append("<webParts>");
            wp.Append("	<webPart xmlns='http://schemas.microsoft.com/WebPart/v3'>");
            wp.Append("		<metaData>");
#if ONPREMISES
            wp.Append("			<type name='Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' />");
#else
            wp.Append("			<type name='Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' />");
#endif
            wp.Append("			<importErrorMessage>Cannot import this Web Part.</importErrorMessage>");
            wp.Append("		</metaData>");
            wp.Append("		<data>");
            wp.Append("			<properties>");
            wp.Append("				<property name='Title' type='string'>$Resources:core,ScriptEditorWebPartTitle;</property>");
            wp.Append("				<property name='Description' type='string'>$Resources:core,ScriptEditorWebPartDescription;</property>");
            wp.Append("				<property name='ChromeType' type='chrometype'>None</property>");
            wp.Append("				<property name='Content' type='string'>");
            wp.Append("				<![CDATA[");
            wp.Append("				    <div id='embedded-feed' style='height: 500px;'></div>");
            wp.Append("				    <script type='text/javascript' src='https://assets.yammer.com/assets/platform_embed.js'></script>");
            wp.Append("				    <script type='text/javascript'>  yam.connect.embedFeed({ container: '#embedded-feed', network: '" + yammerNetworkName
                                                        + @"', feedType: 'group', feedId: '" + yammerGroupId
                                                        + @"', config: { use_sso: " + useSSO.ToString().ToLower()
                                                                + @", header: " + showHeader.ToString().ToLower()
                                                                + @", footer: " + showFooter.ToString().ToLower()
                                                                + " }}); </script>");
            wp.Append("				]]>");
            wp.Append("				</property>");
            wp.Append("			</properties>");
            wp.Append("		</data>");
            wp.Append("	</webPart>");
            wp.Append("</webParts>");

            return wp.ToString();
        }

        /// <summary>
        /// Constructs web part definition for Open Graph discussion web part definition
        /// </summary>
        /// <param name="yammerNetworkName">Yammer network name</param>
        /// <param name="url">Yammer URL</param>
        /// <param name="showHeader">Shows header based on value</param>
        /// <param name="showFooter">Shows footer based on value</param>
        /// <param name="postImageUrl">creates webpart with given post title</param>
        /// <param name="useSso">uses given Single Sign-on information to create webpart</param>
        /// <param name="postTitle">creates webpart with given post title</param>
        /// <param name="groupId">yammer group id</param>
        /// <returns>The constructed Open Graph Discussion Part XML</returns>
        public static string CreateYammerOpenGraphDiscussionPartXml(string yammerNetworkName, string url, bool showHeader,
                                                                    bool showFooter, string postTitle = "", string postImageUrl = "",
                                                                    bool useSso = true, string groupId = "")
        {
            StringBuilder wp = new StringBuilder(100);
            wp.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            wp.Append("<webParts>");
            wp.Append("	<webPart xmlns='http://schemas.microsoft.com/WebPart/v3'>");
            wp.Append("		<metaData>");
#if ONPREMISES
            wp.Append("			<type name='Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' />");
#else
            wp.Append("			<type name='Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' />");
#endif
            wp.Append("			<importErrorMessage>Cannot import this Web Part.</importErrorMessage>");
            wp.Append("		</metaData>");
            wp.Append("		<data>");
            wp.Append("			<properties>");
            wp.Append("				<property name='Title' type='string'>$Resources:core,ScriptEditorWebPartTitle;</property>");
            wp.Append("				<property name='Description' type='string'>$Resources:core,ScriptEditorWebPartDescription;</property>");
            wp.Append("				<property name='ChromeType' type='chrometype'>None</property>");
            wp.Append("				<property name='Content' type='string'>");
            wp.Append("				<![CDATA[");
            wp.Append("				    <div id='embedded-feed' style='height: 500px;'></div>");
            wp.Append("				    <script type='text/javascript' src='https://assets.yammer.com/assets/platform_embed.js'></script>");
            wp.Append("				    <script>");
            wp.Append("				        yam.connect.embedFeed({");
            wp.Append("				          container: '#embedded-feed'");
            wp.Append("				                , feedType: 'open-graph'");
            wp.Append("				                , feedId: ''");
            wp.Append("				                , config: {");
            wp.Append("				                     use_sso: " + useSso.ToString().ToLower());
            wp.Append("				                     , header: " + showHeader.ToString().ToLower());
            wp.Append("				                     , footer: " + showFooter.ToString().ToLower());
            wp.Append("				                     , showOpenGraphPreview: false");
            wp.Append("				                     , defaultToCanonical: false");
            wp.Append("				                     , hideNetworkName: false");
            wp.Append("				                     , promptText: 'Start a conversation'");
            if (!string.IsNullOrEmpty(groupId))
            {
                wp.Append("				                 , defaultGroupId: '" + groupId + "'");
            }
            wp.Append("				                }");
            wp.Append("				                , objectProperties: {");
            wp.Append("				                  url: '" + url + "'");
            wp.Append("				                  , type: 'page'");
            wp.Append("				                  , title: '" + postTitle + "'");
            wp.Append("				                  , image: '" + postImageUrl + "'");
            wp.Append("				                }");
            wp.Append("				            });");
            wp.Append("				        </script>");
            wp.Append("				]]>");
            wp.Append("				</property>");
            wp.Append("			</properties>");
            wp.Append("		</data>");
            wp.Append("	</webPart>");
            wp.Append("</webParts>");

            return wp.ToString();
        }

        private static string GetYammerJson(string url, string accessToken)
        {
            //make the request
            string json = null;
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Headers.Add("Authorization", "Bearer" + " " + accessToken);
            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                Encoding encode = Encoding.GetEncoding("utf-8");
                StreamReader reader = new StreamReader(response.GetResponseStream(), encode);
                json = reader.ReadToEnd();
            }
            return json;
        }

        private static string PostYammerJson(string url, string accessToken)
        {
            //make the request
            string json = null;
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Method = "POST";
            request.ContentLength = 0;
            request.Headers.Add("Authorization", "Bearer" + " " + accessToken);
            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                Encoding encode = Encoding.GetEncoding("utf-8");
                StreamReader reader = new StreamReader(response.GetResponseStream(), encode);
                json = reader.ReadToEnd();
            }
            return json;
        }
    }
}
