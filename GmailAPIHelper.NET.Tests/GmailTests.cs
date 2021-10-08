using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace GmailAPIHelper.NET.Tests
{
    [TestClass]
    public class GmailTests : TestBase
    {

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test1_GetLatestMessage()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'TEST']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
            TestContext.WriteLine(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_PlainText()
        {
            var body = System.IO.File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: "EMAIL WITH PLAIN TEXT", body: body);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_HtmlText()
        {
            var body = System.IO.File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\HTMLEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.HTML, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: "EMAIL WITH HTML TEXT", body: body);
        }

        [TestMethod]
        [DoNotParallelize]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MoveMessageToTrash()
        {
            //Test Data
            var subject = Guid.NewGuid().ToString();
            var body = System.IO.File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'" + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);
        }
    }
}