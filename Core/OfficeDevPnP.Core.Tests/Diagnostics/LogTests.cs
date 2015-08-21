using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Tests.Diagnostics
{
    [TestClass]
    public class LogTests
    {
        [TestMethod]
        public void LogTest1()
        {
            Log.Info("Test Source", "Information test message");
            Log.Debug("Test Source", "Debug test message");

            Log.LogLevel = LogLevel.Information;

            Log.Error("Test Source", "Information test message 2");

        }
    }
}
