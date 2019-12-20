using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Static class that provides basic functionalities to deliver webhook notifications
    /// </summary>
    public static class WebhookSender
    {
        /// <summary>
        /// Public method to send a Webhook notification
        /// </summary>
        /// <param name="webhook">The reference webhook to notify</param>
        /// <param name="httpClient">The HttpClient instance to use to send the notification</param>
        /// <param name="kind">The Kind of webhook</param>
        /// <param name="parser">The parser to use for parsing parameters, optional</param>
        /// <param name="objectHandler">The objectHandler name, optional</param>
        /// <param name="exception">The exception, optional</param>
        /// <param name="scope">The PnP Provisioning Scope, if any, optional</param>
        public static void InvokeWebhook(ProvisioningWebhookBase webhook,
            HttpClient httpClient,
            ProvisioningTemplateWebhookKind kind, TokenParser parser = null,
            string objectHandler = null, Exception exception = null, 
            PnPMonitoredScope scope = null)
        {
            var requestParameters = new Dictionary<String, String>();

            if (exception != null)
            {
                // For GET requests we limit the size of the exception to avoid issues
                requestParameters["__exception"] =
                    webhook.Method == ProvisioningTemplateWebhookMethod.GET ?
                    exception.Message : exception.ToString();
            }

            SimpleTokenParser internalParser = new SimpleTokenParser();
            foreach (var webhookparam in webhook.Parameters)
            {
                requestParameters.Add(webhookparam.Key, parser != null ? parser.ParseString(webhookparam.Value) : webhookparam.Value);
                internalParser.AddToken(new WebhookParameter(webhookparam.Key, requestParameters[webhookparam.Key]));
            }
            var url = parser != null ? parser.ParseString(webhook.Url) : webhook.Url; // parse for template scoped parameters
            url = internalParser.ParseString(url); // parse for webhook scoped parameters

            switch (webhook.Method)
            {
                case ProvisioningTemplateWebhookMethod.GET:
                    {
                        url += $"&__webhookKind={kind.ToString()}"; // add the webhook kind to the REST request URL

                        foreach (var k in requestParameters.Keys)
                        {
                            url += $"&{HttpUtility.UrlEncode(k)}={HttpUtility.UrlEncode(requestParameters[k])}";
                        }

                        if ((kind == ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted
                            || kind == ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningCompleted
                            || kind == ProvisioningTemplateWebhookKind.ExceptionOccurred) && objectHandler != null)
                        {
                            url += $"&__handler={HttpUtility.UrlEncode(objectHandler)}"; // add the handler name to the REST request URL
                        }
                        try
                        {
                            if (webhook.Async)
                            {
                                Task.Factory.StartNew(async () =>
                                {
                                    await httpClient.GetAsync(url);
                                });
                            }
                            else
                            {
                                httpClient.GetAsync(url).GetAwaiter().GetResult();
                            }
                        }
                        catch (HttpRequestException ex)
                        {
                            scope?.LogError(ex, "Error calling provisioning template webhook");
                        }
                        break;
                    }
                case ProvisioningTemplateWebhookMethod.POST:
                    {
                        requestParameters.Add("__webhookKind", kind.ToString()); // add the webhook kind to the parameters of the request body

                        if ((kind == ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningCompleted
                            || kind == ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted
                            || kind == ProvisioningTemplateWebhookKind.ExceptionOccurred) && objectHandler != null)
                        {
                            requestParameters.Add("__handler", objectHandler); // add the handler name to the parameters of the request body
                        }
                        try
                        {
                            if (webhook.Async)
                            {
                                Task.Factory.StartNew(async () =>
                                {
                                    switch (webhook.BodyFormat)
                                    {
                                        case ProvisioningTemplateWebhookBodyFormat.Json:
                                            await httpClient.PostAsJsonAsync(url, requestParameters);
                                            break;
                                        case ProvisioningTemplateWebhookBodyFormat.Xml:
                                            await httpClient.PostAsXmlAsync(url, requestParameters);
                                            break;
                                        case ProvisioningTemplateWebhookBodyFormat.FormUrlEncoded:
                                            var content = new FormUrlEncodedContent(requestParameters);
                                            await httpClient.PostAsync(url, content);
                                            break;
                                    }
                                });
                            }
                            else
                            {
                                switch (webhook.BodyFormat)
                                {
                                    case ProvisioningTemplateWebhookBodyFormat.Json:
                                        httpClient.PostAsJsonAsync(url, requestParameters).GetAwaiter().GetResult();
                                        break;
                                    case ProvisioningTemplateWebhookBodyFormat.Xml:
                                        httpClient.PostAsXmlAsync(url, requestParameters).GetAwaiter().GetResult();
                                        break;
                                    case ProvisioningTemplateWebhookBodyFormat.FormUrlEncoded:
                                        var content = new FormUrlEncodedContent(requestParameters);
                                        httpClient.PostAsync(url, content).GetAwaiter().GetResult();
                                        break;
                                }
                            }
                        }
                        catch (HttpRequestException ex)
                        {
                            scope?.LogError(ex, "Error calling provisioning template webhook");
                        }
                        break;
                    }
            }
        }
    }
}
