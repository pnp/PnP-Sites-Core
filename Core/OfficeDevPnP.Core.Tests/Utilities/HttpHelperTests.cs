using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Tests.Utilities
{
    [TestClass]
    public class HttpHelperTests
    {
        static string TestAPIUrl;

        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            HttpHelperTests.TestAPIUrl = ConfigurationManager.AppSettings["HttpHelperFunctionAppUrl"];
        }

        [TestMethod]
        public void MakeGetRequestForStringTest()
        {

            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            var result = HttpHelper.MakeGetRequestForString(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakeGetRequestForString");
            Assert.AreEqual(result, "Here is the string response!");
        }

        [TestMethod]
        public void MakeGetRequestForStreamTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            var mem = HttpHelper.MakeGetRequestForStream(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakeGetRequestForStream", 
                "text/plain");
            using (var sr = new StreamReader(mem))
            {
                var result = sr.ReadToEnd();
                Assert.AreEqual(result, "Here is the Stream response!");
            }
        }

        [TestMethod]
        public void MakeGetRequestForStreamWithResponseHeadersTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            System.Net.Http.Headers.HttpResponseHeaders responseHeaders;
            var mem = HttpHelper.MakeGetRequestForStreamWithResponseHeaders(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakeGetRequestForStreamWithResponseHeaders", 
                "text/plain",
                out responseHeaders);
            using (var sr = new StreamReader(mem))
            {
                var result = sr.ReadToEnd();
                Assert.AreEqual(result, "Here is the Stream response!");
            }

            Assert.IsNotNull(responseHeaders);
        }

        [TestMethod]
        public void MakePostRequestTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            HttpHelper.MakePostRequest(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakePostRequest");
        }

        [TestMethod]
        public void MakePostRequestForStringTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            var result = HttpHelper.MakePostRequestForString(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakePostRequestForString");
            Assert.AreEqual(result, "Here is the string response!");
        }

        [TestMethod]
        public void MakePostRequestForHeadersTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            var result = HttpHelper.MakePostRequestForHeaders(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakePostRequestForHeaders");
            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void MakePutRequestTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            HttpHelper.MakePutRequest(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakePutRequest");
        }

        [TestMethod]
        public void MakePutRequestForStringTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            var result = HttpHelper.MakePutRequestForString(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakePutRequestForString");
            Assert.AreEqual(result, "Here is the string response!");
        }

        [TestMethod]
        public void MakePatchRequestForStringTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            var result = HttpHelper.MakePatchRequestForString(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakePatchRequestForString");
            Assert.AreEqual(result, "I've got your request!");
        }

        [TestMethod]
        public void MakeDeleteRequestTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            HttpHelper.MakeDeleteRequest(
                $"{HttpHelperTests.TestAPIUrl}&requestType=MakeDeleteRequest");
        }

        [TestMethod]
        public void MakeGetRequestForStringWithSPContextTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            using (var clientContext = TestCommon.CreatePnPClientContext(5, 1000))
            {
                var web = clientContext.Web;
                web.EnsureProperties(w => w.Url, w => w.Title);

                var requestHeaders = new Dictionary<string, string>();
                requestHeaders.Add("X-RequestDigest", (clientContext as ClientContext).GetRequestDigest().GetAwaiter().GetResult());

                var result = HttpHelper.MakeGetRequestForString(
                    $"{web.Url}/_api/web",
                    accept: "application/json",
                    requestHeaders: requestHeaders,
                    spContext: clientContext);

                var webJson = JsonConvert.DeserializeAnonymousType(result, new { Title = "" });
                Assert.AreEqual(webJson.Title, web.Title);
            }
        }

        [TestMethod]
        public void MakePostRequestForStringWithSPContextTest()
        {
            if (string.IsNullOrEmpty(HttpHelperTests.TestAPIUrl)) Assert.Inconclusive("No value set for HttpHelperFunctionAppUrl property in the config file");

            using (var clientContext = TestCommon.CreatePnPClientContext(5, 1000))
            {
                var web = clientContext.Web;
                web.EnsureProperties(w => w.Url, w => w.Title);

                var requestHeaders = new Dictionary<string, string>();
                requestHeaders.Add("X-RequestDigest", (clientContext as ClientContext).GetRequestDigest().GetAwaiter().GetResult());

                var result = HttpHelper.MakePostRequestForString(
                    $"{web.Url}/_api/web",
                    accept: "application/json",
                    requestHeaders: requestHeaders,
                    spContext: clientContext);

                var webJson = JsonConvert.DeserializeAnonymousType(result, new { Title = "" });
                Assert.AreEqual(webJson.Title, web.Title);
            }
        }
    }
}
