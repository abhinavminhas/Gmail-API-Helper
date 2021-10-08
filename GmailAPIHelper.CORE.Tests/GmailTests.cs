using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace GmailAPIHelper.CORE.Tests
{
    [TestClass]
    public class GmailTests : TestBase
    {

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'TEST']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
            TestContext.WriteLine(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_SendMessage_PlainText()
        {
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: "EMAIL WITH PLAIN TEXT", body: body);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_SendMessage_HtmlText()
        {
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\HTMLEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/HTMLEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/HTMLEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.HTML, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: "EMAIL WITH HTML TEXT", body: body);
        }

        [TestMethod]
        [DoNotParallelize]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MoveMessageToTrash()
        {
            //Test Data
            var subject = Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'"+ subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);
        }

    }
}