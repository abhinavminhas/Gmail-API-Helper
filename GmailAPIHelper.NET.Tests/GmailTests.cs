﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
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
        public void Test_GmailService_Dispose()
        {
            //Dispose (service argument)
            var service = GmailHelper.GetGmailService(ApplicatioName);
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
            service = GmailHelper.GetGmailService(ApplicatioName);
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
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetMessages()
        {
            var messages = GmailHelper.GetGmailService(ApplicatioName)
                .GetMessages(query: "[from:test.auto.helper@gmail.com][subject:'EMAIL']in:inbox is:read", markRead: true);
            Assert.AreEqual(5, messages.Count);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetMessage_NoMatchingEmail()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:read", markRead: true);
            Assert.IsNull(message);
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
        public void Test_GetLatestMessage_Multipart_HtmlText()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[subject:'READ EMAIL WITH HTML TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_GetLatestMessage_MultipleMatchingEmails()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
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
            var subject = "MOVE DOTNETFRAMEWORK MESSAGE TO TRASH " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'MOVE DOTNETFRAMEWORK MESSAGE TO TRASH " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_MoveMessagesToTrash_NoMatchingEmail()
        {
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
            Assert.AreEqual(0, countOfMessagesMovedToTrash);
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
            var subject = "MOVE DOTNETFRAMEWORK MESSAGES TO TRASH " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
            {
                var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
                GmailHelper.GetGmailService(ApplicatioName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
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
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'UNTRASH DOTNETFRAMEWORK MESSAGE  " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);

            //Test Run
            var isUntrashed = GmailHelper.GetGmailService(ApplicatioName)
                .UntrashMessage(query: "[from:test.auto.helper@gmail.com][subject:'UNTRASH DOTNETFRAMEWORK MESSAGE  " + subject + "']in:trash is:unread");
            Assert.IsTrue(isUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UntrashMessage_NoMatchingEmail()
        {
            var isUntrashed = GmailHelper.GetGmailService(ApplicatioName)
                .UntrashMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
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
                GmailHelper.GetGmailService(ApplicatioName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com][subject:'UNTRASH DOTNETFRAMEWORK MESSAGES " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMovedToTrash);

            //Test Run
            var countOfMessagesUntrashed = GmailHelper.GetGmailService(ApplicatioName)
                .UntrashMessages(query: "[from:test.auto.helper@gmail.com][subject:'UNTRASH DOTNETFRAMEWORK MESSAGES " + subject + "']in:trash is:unread");
            Assert.AreEqual(2, countOfMessagesUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UntrashMessages_NoMatchingEmail()
        {
            var countOfMessagesUntrashed = GmailHelper.GetGmailService(ApplicatioName)
                .UntrashMessages(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
            Assert.AreEqual(0, countOfMessagesUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ReportSpamMessage()
        {
            //Test Data
            var subject = "REPORT DOTNETFRAMEWORK SPAM MESSAGE " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isSpamReported = GmailHelper.GetGmailService(ApplicatioName)
                .ReportSpamMessage(query: "[from:test.auto.helper@gmail.com][subject:'REPORT DOTNETFRAMEWORK SPAM MESSAGE " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isSpamReported);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ReportSpamMessage_NoMatchingEmail()
        {
            var isSpamReported = GmailHelper.GetGmailService(ApplicatioName)
                .ReportSpamMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
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
                GmailHelper.GetGmailService(ApplicatioName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesReportedAsSpam = GmailHelper.GetGmailService(ApplicatioName)
                .ReportSpamMessages(query: "[from:test.auto.helper@gmail.com][subject:'REPORT DOTNETFRAMEWORK SPAM MESSAGES " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesReportedAsSpam);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ReportSpamMessages_NoMatchingEmail()
        {
            var countOfMessagesReportedAsSpam = GmailHelper.GetGmailService(ApplicatioName)
                .ReportSpamMessages(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
            Assert.AreEqual(0, countOfMessagesReportedAsSpam);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UnspamMessage()
        {
            //Test Data
            var subject = "UNSPAM DOTNETFRAMEWORK MESSAGE " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            var isSpamReported = GmailHelper.GetGmailService(ApplicatioName)
                .ReportSpamMessage(query: "[from:test.auto.helper@gmail.com][subject:'UNSPAM DOTNETFRAMEWORK MESSAGE " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isSpamReported);

            //Test Run
            var isUnspamed = GmailHelper.GetGmailService(ApplicatioName)
                .UnspamMessage(query: "[from:test.auto.helper@gmail.com][subject:'UNSPAM DOTNETFRAMEWORK MESSAGE " + subject + "']in:spam is:unread");
            Assert.IsTrue(isUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_UnspamMessage_NoMatchingEmail()
        {
            var isUnspamed = GmailHelper.GetGmailService(ApplicatioName)
                .UnspamMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
            Assert.IsFalse(isUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessage()
        {
            //Test Data
            var subject = "MODIFY DOTNETFRAMEWORK MESSAGE " + Guid.NewGuid().ToString();
            var body = File.ReadAllText(Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt");
            GmailHelper.GetGmailService(ApplicatioName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);

            //Test Run
            var isModified = GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGE " + subject + "']in:inbox", labelsToAdd: new List<string>() { "IMPORTANT", "SPAM", });
            Assert.IsTrue(isModified);
            isModified = GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGE " + subject + "']in:spam", labelsToRemove: new List<string>() { "IMPORTANT", "UNREAD" });
            Assert.IsTrue(isModified);
            isModified = GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGE " + subject + "']in:spam", labelsToAdd: new List<string>() { "INBOX", "STARRED", "UNREAD", }, labelsToRemove: new List<string>() { "SPAM" });
            Assert.IsTrue(isModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessage_NoLabelsSupplied()
        {
            try
            {
                GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (NullReferenceException ex) { Assert.AreEqual("Either 'Labels To Add' or 'Labels to Remove' required.", ex.Message); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessage_NoMatchingEmail()
        {
            var isModified = GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread", labelsToAdd: new List<string>() { "STARRED", "IMPORTANT", }, labelsToRemove: new List<string>() { "UNREAD" });
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
                GmailHelper.GetGmailService(ApplicatioName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesModified = GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessages(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGES " + subject + "']in:inbox", labelsToAdd: new List<string>() { "IMPORTANT", "SPAM", });
            Assert.AreEqual(2, countOfMessagesModified);
            countOfMessagesModified = GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessages(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGES " + subject + "']in:spam", labelsToRemove: new List<string>() { "IMPORTANT", "UNREAD" });
            Assert.AreEqual(2, countOfMessagesModified);
            countOfMessagesModified = GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessages(query: "[from:test.auto.helper@gmail.com][subject:'MODIFY DOTNETFRAMEWORK MESSAGES " + subject + "']in:spam", labelsToAdd: new List<string>() { "INBOX", "STARRED", "UNREAD", }, labelsToRemove: new List<string>() { "SPAM" });
            Assert.AreEqual(2, countOfMessagesModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessages_NoLabelsSupplied()
        {
            try
            {
                GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessages(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (NullReferenceException ex) { Assert.AreEqual("Either 'Labels To Add' or 'Labels to Remove' required.", ex.Message); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETFRAMEWORK")]
        public void Test_ModifyMessages_NoMatchingEmail()
        {
            var countOfMessagesModified = GmailHelper.GetGmailService(ApplicatioName)
                .ModifyMessages(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox", labelsToAdd: new List<string>() { "STARRED", "IMPORTANT", });
            Assert.AreEqual(0, countOfMessagesModified);
        }

        [TestMethod]
        [TestCategory("TestCleanup")]
        public void Inbox_CleanUp()
        {
            GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com]in:inbox is:unread");
            GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com]in:spam is:unread");
        }
    }
}