using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace GmailAPIHelper.CORE.Tests
{
    [TestClass]
    public class GmailTests : TestBase
    {
        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_EmailContentType()
        {
            Assert.AreEqual("PLAIN", GmailHelper.EmailContentType.PLAIN.ToString());
            Assert.AreEqual(1, (int)GmailHelper.EmailContentType.PLAIN);
            Assert.AreEqual("HTML", GmailHelper.EmailContentType.HTML.ToString());
            Assert.AreEqual(2, (int)GmailHelper.EmailContentType.HTML);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
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
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetGmailService_TokenPath_Home()
        {
            var sourcePath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                sourcePath = Environment.CurrentDirectory + "\\" + "token.json" + "\\Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                sourcePath = Environment.CurrentDirectory + "/" + "token.json" + "/Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                sourcePath = Environment.CurrentDirectory + "/" + "token.json" + "/Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
            var destPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                destPath = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%") + "\\" + "token.json";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                destPath = Environment.GetEnvironmentVariable("HOME") + "/" + "token.json";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                destPath = Environment.GetEnvironmentVariable("HOME") + "/" + "token.json";
            Directory.CreateDirectory(destPath);
            File.Copy(sourcePath, destPath + "/Google.Apis.Auth.OAuth2.Responses.TokenResponse-user", overwrite: true);
            var message = GmailHelper.GetGmailService(ApplicatioName, GmailHelper.TokenPathType.HOME);
            Directory.Delete(destPath, recursive: true);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetGmailService_TokenPath_Custom()
        {
            var credPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                credPath = Environment.CurrentDirectory + "\\" + "token.json";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                credPath = Environment.CurrentDirectory + "/" + "token.json";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                credPath = Environment.CurrentDirectory + "/" + "token.json";
            var message = GmailHelper.GetGmailService(ApplicatioName, GmailHelper.TokenPathType.CUSTOM, credPath);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetGmailService_TokenPath_Custom_EmptyPath()
        {
            try
            {
                GmailHelper.GetGmailService(ApplicatioName, GmailHelper.TokenPathType.CUSTOM);
                Assert.Fail("No Argument Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (ArgumentException) { }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessage()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessages()
        {
            var messages = GmailHelper.GetGmailService(ApplicatioName)
                .GetMessages(query: "[from:test.auto.helper@gmail.com][subject:'EMAIL']in:inbox is:read", markRead: true);
            Assert.AreEqual(5, messages.Count);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessage_NoMatchingEmail()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:read", markRead: true);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_PlainText()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH PLAIN TEXT (TEXT/PLAIN)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_HtmlText()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH HTML TEXT (TEXT/HTML)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_Multipart_NoText()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL WITH NO TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_Multipart_PlainText()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[subject:'READ EMAIL WITH PLAIN TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_Multipart_HtmlText()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[subject:'READ EMAIL WITH HTML TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_MultipleMatchingEmails()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_NoMatchingEmail()
        {
            var message = GmailHelper.GetGmailService(ApplicatioName)
                .GetLatestMessage(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:read", markRead: true);
            Assert.IsNull(message);
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
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
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
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
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

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MoveMessageToTrash_NoMatchingEmail()
        {
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessageToTrash(query: "[from:test.auto.helper@gmail.com][subject:'Email does not exists']in:inbox is:unread");
            Assert.IsFalse(isMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MoveMessagesToTrash()
        {
            //Test Data
            var subject = "MOVE DOTNETCORE MESSAGES TO TRASH " + Guid.NewGuid().ToString();
            for (int i = 0; i < 2; i++)
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
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, "test.auto.helper@gmail.com", cc: "test.auto.helper@gmail.com", bcc: "test.auto.helper@gmail.com", subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com][subject:'MOVE DOTNETCORE MESSAGES TO TRASH " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMovedToTrash);
        }

        [TestMethod]
        [TestCategory("TestCleanup")]
        public void Inbox_CleanUp()
        {
            GmailHelper.GetGmailService(ApplicatioName)
                .MoveMessagesToTrash(query: "[from:test.auto.helper@gmail.com]in:inbox is:unread");
        }
    }
}