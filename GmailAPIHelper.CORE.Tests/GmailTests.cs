using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GmailAPIHelper.CORE.Tests
{
    [TestClass]
    public class GmailTests : TestBase
    {

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test1_GetLatestMessage()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'TEST']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
            TestContext.WriteLine(message);
        }
    }
}