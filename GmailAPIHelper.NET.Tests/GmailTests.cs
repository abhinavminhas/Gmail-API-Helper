using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
        public void Test_GmailService_Dispose()
        {
            //Dispose (service argument)
            var service = GmailHelper.GetGmailService(ApplicationName);
            Assert.IsNotNull(service);
            GmailHelper.DisposeGmailService(service);
            try
            {
                service.GetMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
                Assert.Fail("No Object Disposed Exception Thrown");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (ObjectDisposedException ex)
            {
                Assert.IsTrue(ex.Message.Contains("Cannot access a disposed object."));
            }

            //Dispose (service extension)
            service = GmailHelper.GetGmailService(ApplicationName);
            Assert.IsNotNull(service);
            service.DisposeGmailService();
            try
            {
                service.GetMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
                Assert.Fail("No Object Disposed Exception Thrown");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (ObjectDisposedException ex)
            {
                Assert.IsTrue(ex.Message.Contains("Cannot access a disposed object."));
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetMessage()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetMessage_NoMatchingEmail()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetMessage(query: EmailDoesNotExistsSearchQuery);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetMessages()
        {
            var messages = GmailHelper.GetGmailService(ApplicationName)
                .GetMessages(query: "[from:test.auto.helper@gmail.com][subject:'EMAIL']in:inbox is:read", markRead: true);
            Assert.AreEqual(5, messages.Count);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetMessages_NoMatchingEmail()
        {
            var messages = GmailHelper.GetGmailService(ApplicationName)
                .GetMessages(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, messages.Count);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_PlainText()
        {
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH PLAIN TEXT (TEXT/PLAIN)']in:inbox is:read", markRead: true);
            Assert.AreEqual(body, message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_HtmlText()
        {
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\HTMLEmail.txt");
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH HTML TEXT (TEXT/HTML)']in:inbox is:read", markRead: true);
            Assert.AreEqual(body, message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_Multipart_NoText()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH NO TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_Multipart_PlainText()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: "[subject:'READ EMAIL WITH PLAIN TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_Multipart_HtmlText()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: "[subject:'READ EMAIL WITH HTML TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_MultipleMatchingEmails()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_NoMatchingEmail()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: EmailDoesNotExistsSearchQuery, markRead: true);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_PlainText()
        {
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: "EMAIL WITH PLAIN TEXT", body: body);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_SendMessage_HtmlText()
        {
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\HTMLEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
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
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, invalidEmailType);
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
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: invalidEmailType);
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
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", bcc: invalidEmailType);
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
            var subject = "MOVE DOTNETFRAMEWORK MESSAGE TO TRASH " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'MOVE DOTNETFRAMEWORK MESSAGE TO TRASH " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MoveMessagesToTrash_NoMatchingEmail()
        {
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MoveMessageToTrash_NoMatchingEmail()
        {
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessageToTrash(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MoveMessagesToTrash()
        {
            //Test Data
            var subject = "MOVE DOTNETFRAMEWORK MESSAGES TO TRASH " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
            {
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com][subject:'MOVE DOTNETFRAMEWORK MESSAGES TO TRASH '" + subject + "]in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UntrashMessage()
        {
            //Test Data
            var subject = "UNTRASH DOTNETFRAMEWORK MESSAGE " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'UNTRASH DOTNETFRAMEWORK MESSAGE  " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);

            //Test Run
            var isUntrashed = GmailHelper.GetGmailService(ApplicationName)
                .UntrashMessage(query: "[from:test.auto.helper@gmail.com][subject:'UNTRASH DOTNETFRAMEWORK MESSAGE  " + subject + "']in:trash is:unread");
            Assert.IsTrue(isUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UntrashMessage_NoMatchingEmail()
        {
            var isUntrashed = GmailHelper.GetGmailService(ApplicationName)
                .UntrashMessage(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UntrashMessages()
        {
            //Test Data
            var subject = "UNTRASH DOTNETFRAMEWORK MESSAGES " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
            {
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com][subject:'UNTRASH DOTNETFRAMEWORK MESSAGES " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMovedToTrash);

            //Test Run
            var countOfMessagesUntrashed = GmailHelper.GetGmailService(ApplicationName)
                .UntrashMessages(query: "[from:test.auto.helper@gmail.com][subject:'UNTRASH DOTNETFRAMEWORK MESSAGES " + subject + "']in:trash is:unread");
            Assert.AreEqual(2, countOfMessagesUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UntrashMessages_NoMatchingEmail()
        {
            var countOfMessagesUntrashed = GmailHelper.GetGmailService(ApplicationName)
                .UntrashMessages(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ReportSpamMessage()
        {
            //Test Data
            var subject = "REPORT DOTNETFRAMEWORK SPAM MESSAGE " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isSpamReported = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessage(query: "[from:test.auto.helper@gmail.com][subject:'REPORT DOTNETFRAMEWORK SPAM MESSAGE " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isSpamReported);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ReportSpamMessage_NoMatchingEmail()
        {
            var isSpamReported = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessage(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isSpamReported);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ReportSpamMessages()
        {
            //Test Data
            var subject = "REPORT DOTNETFRAMEWORK SPAM MESSAGES " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
            {
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesReportedAsSpam = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessages(query: "[from:test.auto.helper@gmail.com][subject:'REPORT DOTNETFRAMEWORK SPAM MESSAGES " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesReportedAsSpam);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ReportSpamMessages_NoMatchingEmail()
        {
            var countOfMessagesReportedAsSpam = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessages(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesReportedAsSpam);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UnspamMessage()
        {
            //Test Data
            var subject = "UNSPAM DOTNETFRAMEWORK MESSAGE " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            var isSpamReported = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessage(query: "[from:test.auto.helper@gmail.com][subject:'UNSPAM DOTNETFRAMEWORK MESSAGE " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isSpamReported);

            //Test Run
            var isUnspamed = GmailHelper.GetGmailService(ApplicationName)
                .UnspamMessage(query: "[from:test.auto.helper@gmail.com][subject:'UNSPAM DOTNETFRAMEWORK MESSAGE " + subject + "']in:spam is:unread");
            Assert.IsTrue(isUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UnspamMessage_NoMatchingEmail()
        {
            var isUnspamed = GmailHelper.GetGmailService(ApplicationName)
                .UnspamMessage(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UnspamMessages()
        {
            //Test Data
            var subject = "UNSPAM DOTNETFRAMEWORK MESSAGES " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
            {
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }
            var countOfMessagesReportedAsSpam = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessages(query: "[from:test.auto.helper@gmail.com][subject:'UNSPAM DOTNETFRAMEWORK MESSAGES " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesReportedAsSpam);

            //Test Run
            var countOfMessagesUnspamed = GmailHelper.GetGmailService(ApplicationName)
                .UnspamMessages(query: "[from:test.auto.helper@gmail.com][subject:'UNSPAM DOTNETFRAMEWORK MESSAGES " + subject + "']in:spam is:unread");
            Assert.AreEqual(2, countOfMessagesUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UnspamMessages_NoMatchingEmail()
        {
            var countOfMessagesUnspamed = GmailHelper.GetGmailService(ApplicationName)
                .UnspamMessages(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MarkMessageAsRead()
        {
            //Test Data
            var subject = "MARK DOTNETFRAMEWORK MESSAGE AS READ " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isMarkedRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsRead(query: "[from:test.auto.helper@gmail.com][subject:'MARK DOTNETFRAMEWORK MESSAGE AS READ  " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMarkedRead);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MarkMessageAsRead_NoMatchingEmail()
        {
            var isMarkedRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsRead(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isMarkedRead);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MarkMessagesAsRead()
        {
            //Test Data
            var subject = "MARK DOTNETFRAMEWORK MESSAGES AS READ " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
            {
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesMarkedAsRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsRead(query: "[from:test.auto.helper@gmail.com][subject:'MARK DOTNETFRAMEWORK MESSAGES AS READ " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMarkedAsRead);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MarkMessagesAsRead_NoMatchingEmail()
        {
            var countOfMessagesMarkedAsRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsRead(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesMarkedAsRead);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MarkMessageAsUnread()
        {
            //Test Data
            var subject = "MARK DOTNETFRAMEWORK MESSAGE AS UNREAD " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            var isMarkedRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsRead(query: "[from:test.auto.helper@gmail.com][subject:'MARK DOTNETFRAMEWORK MESSAGE AS UNREAD  " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMarkedRead);

            //Test Run
            var isMarkedUnread = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsUnread(query: "[from:test.auto.helper@gmail.com][subject:'MARK DOTNETFRAMEWORK MESSAGE AS UNREAD  " + subject + "']in:inbox is:read");
            Assert.IsTrue(isMarkedUnread);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MarkMessageAsUnread_NoMatchingEmail()
        {
            var isMarkedUnread = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsUnread(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isMarkedUnread);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MarkMessagesAsUnread()
        {
            //Test Data
            var subject = "MARK DOTNETFRAMEWORK MESSAGES AS UNREAD " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
            {
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }
            var countOfMessagesMarkedAsRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsRead(query: "[from:test.auto.helper@gmail.com][subject:'MARK DOTNETFRAMEWORK MESSAGES AS UNREAD " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMarkedAsRead);

            //Test Run
            var countOfMessagesMarkedAsUnread = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsUnread(query: "[from:test.auto.helper@gmail.com][subject:'MARK DOTNETFRAMEWORK MESSAGES AS UNREAD " + subject + "']in:inbox is:read");
            Assert.AreEqual(2, countOfMessagesMarkedAsUnread);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MarkMessagesAsUnread_NoMatchingEmail()
        {
            var countOfMessagesMarkedAsUnread = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsUnread(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesMarkedAsUnread);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessage()
        {
            //Test Data
            var subject = "MODIFY DOTNETFRAMEWORK MESSAGE " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGE " + subject + "']in:inbox", labelsToAdd: new List<string>() { "IMPORTANT", "SPAM", });
            Assert.IsTrue(isModified);
            isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGE " + subject + "']in:spam", labelsToRemove: new List<string>() { "IMPORTANT", "UNREAD" });
            Assert.IsTrue(isModified);
            isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGE " + subject + "']in:spam", labelsToAdd: new List<string>() { "INBOX", "STARRED", "UNREAD", }, labelsToRemove: new List<string>() { "SPAM" });
            Assert.IsTrue(isModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessage_NoLabelsSupplied()
        {
            try
            {
                GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: EmailDoesNotExistsSearchQuery);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (NullReferenceException ex) { Assert.AreEqual("Either 'Labels To Add' or 'Labels to Remove' required.", ex.Message); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessage_NoMatchingEmail()
        {
            var isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: EmailDoesNotExistsSearchQuery, labelsToAdd: new List<string>() { "STARRED", "IMPORTANT", }, labelsToRemove: new List<string>() { "UNREAD" });
            Assert.IsFalse(isModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessages()
        {
            //Test Data
            var subject = "MODIFY DOTNETFRAMEWORK MESSAGES " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
            {
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGES " + subject + "']in:inbox", labelsToAdd: new List<string>() { "IMPORTANT", "SPAM", });
            Assert.AreEqual(2, countOfMessagesModified);
            countOfMessagesModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGES " + subject + "']in:spam", labelsToRemove: new List<string>() { "IMPORTANT", "UNREAD" });
            Assert.AreEqual(2, countOfMessagesModified);
            countOfMessagesModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGES " + subject + "']in:spam", labelsToAdd: new List<string>() { "INBOX", "STARRED", "UNREAD", }, labelsToRemove: new List<string>() { "SPAM" });
            Assert.AreEqual(2, countOfMessagesModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessages_NoLabelsSupplied()
        {
            try
            {
                GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: EmailDoesNotExistsSearchQuery);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (NullReferenceException ex) { Assert.AreEqual("Either 'Labels To Add' or 'Labels to Remove' required.", ex.Message); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessages_NoMatchingEmail()
        {
            var countOfMessagesModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: EmailDoesNotExistsSearchQuery, labelsToAdd: new List<string>() { "STARRED", "IMPORTANT", });
            Assert.AreEqual(0, countOfMessagesModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetMessageLabels()
        {
            //Test Data
            var subject = "GET DOTNETFRAMEWORK MESSAGE LABELS " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            var isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'GET DOTNETFRAMEWORK MESSAGE LABELS " + subject + "']in:inbox", labelsToAdd: new List<string>() { "IMPORTANT", "STARRED", });
            Assert.IsTrue(isModified);

            //Test Run
            var labels = GmailHelper.GetGmailService(ApplicationName)
                .GetMessageLabels(query: "GET DOTNETFRAMEWORK MESSAGE LABELS " + subject + "']in:inbox");
            Assert.AreEqual(5, labels.Count);
            var importantLabel = labels.FirstOrDefault(x => x.Name.Equals("IMPORTANT"));
            Assert.IsNotNull(importantLabel);
            Assert.AreEqual("IMPORTANT", importantLabel.Id);
            Assert.AreEqual("labelHide", importantLabel.LabelListVisibility);
            Assert.AreEqual("hide", importantLabel.MessageListVisibility);
            Assert.AreEqual("IMPORTANT", importantLabel.Name);
            Assert.AreEqual("system", importantLabel.Type);
            var starredLabel = labels.FirstOrDefault(x => x.Name.Equals("STARRED"));
            Assert.IsNotNull(starredLabel);
            Assert.AreEqual("STARRED", starredLabel.Id);
            Assert.AreEqual(null, starredLabel.LabelListVisibility);
            Assert.AreEqual(null, starredLabel.MessageListVisibility);
            Assert.AreEqual("STARRED", starredLabel.Name);
            Assert.AreEqual("system", starredLabel.Type);
            var unreadLabel = labels.FirstOrDefault(x => x.Name.Equals("UNREAD"));
            Assert.IsNotNull(unreadLabel);
            Assert.AreEqual("UNREAD", unreadLabel.Id);
            Assert.AreEqual(null, unreadLabel.LabelListVisibility);
            Assert.AreEqual(null, unreadLabel.MessageListVisibility);
            Assert.AreEqual("UNREAD", unreadLabel.Name);
            Assert.AreEqual("system", unreadLabel.Type);
            var inboxLabel = labels.FirstOrDefault(x => x.Name.Equals("INBOX"));
            Assert.IsNotNull(inboxLabel);
            Assert.AreEqual("INBOX", inboxLabel.Id);
            Assert.AreEqual(null, inboxLabel.LabelListVisibility);
            Assert.AreEqual(null, inboxLabel.MessageListVisibility);
            Assert.AreEqual("INBOX", inboxLabel.Name);
            Assert.AreEqual("system", inboxLabel.Type);
            var sentLabel = labels.FirstOrDefault(x => x.Name.Equals("SENT"));
            Assert.IsNotNull(sentLabel);
            Assert.AreEqual("SENT", sentLabel.Id);
            Assert.AreEqual(null, sentLabel.LabelListVisibility);
            Assert.AreEqual(null, sentLabel.MessageListVisibility);
            Assert.AreEqual("SENT", sentLabel.Name);
            Assert.AreEqual("system", sentLabel.Type);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetMessageLabels_NoMatchingEmail()
        {
            var labels = GmailHelper.GetGmailService(ApplicationName)
                .GetMessageLabels(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, labels.Count);
        }

        [TestMethod]
        [TestCategory("TestCleanup")]
        public void Inbox_CleanUp()
        {
            GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com]in:inbox is:unread");
            GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com]in:spam is:unread");
            GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[subject:'MARK DOTNETCORE MESSAGE AS READ']in:inbox is:read");
            GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[subject:'MARK DOTNETFRAMEWORK MESSAGE AS READ']in:inbox is:read");
            GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[subject:'MARK DOTNETCORE MESSAGES AS READ']in:inbox is:read");
            GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[subject:'MARK DOTNETFRAMEWORK MESSAGES AS READ']in:inbox is:read");
        }
    }
}