using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace GmailAPIHelper.NET.Tests
{
    [TestClass]
    public class GmailTests : TestBase
    {
        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_EmailContentType()
        {
            Assert.AreEqual("PLAIN", GmailHelper.EmailContentType.PLAIN.ToString());
            Assert.AreEqual(1, (int)GmailHelper.EmailContentType.PLAIN);
            Assert.AreEqual("HTML", GmailHelper.EmailContentType.HTML.ToString());
            Assert.AreEqual(2, (int)GmailHelper.EmailContentType.HTML);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_TokenPathType()
        {
            Assert.AreEqual("HOME", GmailHelper.TokenPathType.HOME.ToString());
            Assert.AreEqual(1, (int)GmailHelper.TokenPathType.HOME);
            Assert.AreEqual("WORKING_DIRECTORY", GmailHelper.TokenPathType.WORKING_DIRECTORY.ToString());
            Assert.AreEqual(2, (int)GmailHelper.TokenPathType.WORKING_DIRECTORY);
            Assert.AreEqual("CUSTOM", GmailHelper.TokenPathType.CUSTOM.ToString());
            Assert.AreEqual(3, (int)GmailHelper.TokenPathType.CUSTOM);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_PlainText()
        {
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH PLAIN TEXT (TEXT/PLAIN)']in:inbox is:read", markRead: true);
            Assert.AreEqual(body, message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_HtmlText()
        {
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\HTMLEmail.txt");
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH HTML TEXT (TEXT/HTML)']in:inbox is:read", markRead: true);
            Assert.AreEqual(body, message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_Multipart_NoText()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH NO TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_Multipart_PlainText()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[subject:'READ EMAIL WITH PLAIN TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_NoMatchingEmail()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:read", markRead: true);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_PlainText()
        {
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: "EMAIL WITH PLAIN TEXT", body: body);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_HtmlText()
        {
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\HTMLEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.HTML, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: "EMAIL WITH HTML TEXT", body: body);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_InvalidToEmail()
        {
            var invalidEmailTypes = new string[] { "testgmail.com", "test@gmailcom" , "1test@gmail.com",
                "test@gmail.com,testgmail.com", "test@gmail.com,test@gmailcom", "test@gmail.com,1test@gmail.com" };
            foreach (var invalidEmailType in invalidEmailTypes)
            {
                try
                {
                    GmailHelper.GetGmailService(ApplicatioName).SendMessage(GmailHelper.EmailContentType.PLAIN, invalidEmailType);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (Exception ex)
                {
                    var invalidEmail = invalidEmailType.Contains(",") ? invalidEmailType.Split(',')[1] : invalidEmailType;
                    Assert.AreEqual(string.Format("Not a valid 'To' email address. Email: '{0}'", invalidEmail), ex.Message);
                }
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_InvalidCcEmail()
        {
            var invalidEmailTypes = new string[] { "testgmail.com", "test@gmailcom" , "1test@gmail.com",
                "test@gmail.com,testgmail.com", "test@gmail.com,test@gmailcom", "test@gmail.com,1test@gmail.com" };
            foreach (var invalidEmailType in invalidEmailTypes)
            {
                try
                {
                    GmailHelper.GetGmailService(ApplicatioName).SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: invalidEmailType);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (Exception ex)
                {
                    var invalidEmail = invalidEmailType.Contains(",") ? invalidEmailType.Split(',')[1] : invalidEmailType;
                    Assert.AreEqual(string.Format("Not a valid 'Cc' email address. Email: '{0}'", invalidEmail), ex.Message);
                }
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_InvalidBccEmail()
        {
            var invalidEmailTypes = new string[] { "testgmail.com", "test@gmailcom" , "1test@gmail.com",
                "test@gmail.com,testgmail.com", "test@gmail.com,test@gmailcom", "test@gmail.com,1test@gmail.com" };
            foreach (var invalidEmailType in invalidEmailTypes)
            {
                try
                {
                    GmailHelper.GetGmailService(ApplicatioName).SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", bcc: invalidEmailType);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (Exception ex)
                {
                    var invalidEmail = invalidEmailType.Contains(",") ? invalidEmailType.Split(',')[1] : invalidEmailType;
                    Assert.AreEqual(string.Format("Not a valid 'Bcc' email address. Email: '{0}'", invalidEmail), ex.Message);
                }
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MoveMessageToTrash()
        {
            //Test Data
            var subject = Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'" + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MoveMessageToTrash_NoMatchingEmail()
        {
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
            Assert.IsFalse(isMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MoveMessagesToTrash()
        {
            //Test Data
            var subject = new string[2];
            for (int i = 0; i < 2; i++)
            {
                subject[i] = "TEST MOVE MESSAGES TO TRASH " + Guid.NewGuid().ToString();
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicatioName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject[i], body: body);
            }

            //Test Run
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com][subject:'TEST MOVE MESSAGES TO TRASH']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMovedToTrash);
        }
    }
}