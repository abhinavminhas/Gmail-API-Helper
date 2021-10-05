using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GmailAPIHelper.NET.Tests
{
    [TestClass]
    public class GmailTests : TestBase
    {

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test1_GetLatestMessage()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName, GmailHelper.TokenPathType.WORKING_DIRECTORY)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'TEST']in:inbox is:read", markRead: true);
            TestContext.WriteLine(message);
        }
    }
}