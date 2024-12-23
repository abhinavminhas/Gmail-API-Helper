using Google;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

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
        public void Test_LabelListVisibility()
        {
            Assert.AreEqual("LABEL_SHOW", GmailHelper.LabelListVisibility.LABEL_SHOW.ToString());
            Assert.AreEqual(1, (int)GmailHelper.LabelListVisibility.LABEL_SHOW);
            Assert.AreEqual("LABEL_SHOW_IF_UNREAD", GmailHelper.LabelListVisibility.LABEL_SHOW_IF_UNREAD.ToString());
            Assert.AreEqual(2, (int)GmailHelper.LabelListVisibility.LABEL_SHOW_IF_UNREAD);
            Assert.AreEqual("LABEL_HIDE", GmailHelper.LabelListVisibility.LABEL_HIDE.ToString());
            Assert.AreEqual(3, (int)GmailHelper.LabelListVisibility.LABEL_HIDE);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MessageListVisibility()
        {
            Assert.AreEqual("SHOW", GmailHelper.MessageListVisibility.SHOW.ToString());
            Assert.AreEqual(1, (int)GmailHelper.MessageListVisibility.SHOW);
            Assert.AreEqual("HIDE", GmailHelper.MessageListVisibility.HIDE.ToString());
            Assert.AreEqual(2, (int)GmailHelper.MessageListVisibility.HIDE);
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
            var service = GmailHelper.GetGmailService(ApplicationName, GmailHelper.TokenPathType.HOME);
            Assert.IsTrue(service.GetType() == typeof(Google.Apis.Gmail.v1.GmailService));
            Directory.Delete(destPath, recursive: true);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetGmailService_TokenPath_WorkingDirectory()
        {
            var service = GmailHelper.GetGmailService(ApplicationName, GmailHelper.TokenPathType.WORKING_DIRECTORY);
            Assert.IsTrue(service.GetType() == typeof(Google.Apis.Gmail.v1.GmailService));
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetGmailService_TokenPath_Custom()
        {
            var tokenPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                tokenPath = Environment.CurrentDirectory + "\\" + "token.json";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                tokenPath = Environment.CurrentDirectory + "/" + "token.json";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                tokenPath = Environment.CurrentDirectory + "/" + "token.json";
            var service = GmailHelper.GetGmailService(ApplicationName, GmailHelper.TokenPathType.CUSTOM, tokenPath);
            Assert.IsTrue(service.GetType() == typeof(Google.Apis.Gmail.v1.GmailService));
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetGmailService_TokenPath_Custom_EmptyPath()
        {
            try
            {
                GmailHelper.GetGmailService(ApplicationName, GmailHelper.TokenPathType.CUSTOM);
                Assert.Fail("No Argument Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (ArgumentException) { }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetGmailService_CredentialsPath_Custom()
        {
            var tokenPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                tokenPath = Environment.CurrentDirectory + "\\" + "token.json";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                tokenPath = Environment.CurrentDirectory + "/" + "token.json";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                tokenPath = Environment.CurrentDirectory + "/" + "token.json";
            var credentialsPath = Environment.CurrentDirectory + "/" + "credentials.json";
            // Token Path & Credentials Path
            var service = GmailHelper.GetGmailService(ApplicationName, GmailHelper.TokenPathType.CUSTOM, tokenPath, credentialsPath);
            Assert.IsTrue(service.GetType() == typeof(Google.Apis.Gmail.v1.GmailService));
            // Only Token Path
            service = GmailHelper.GetGmailService(ApplicationName, GmailHelper.TokenPathType.CUSTOM, tokenPath);
            Assert.IsTrue(service.GetType() == typeof(Google.Apis.Gmail.v1.GmailService));
            // Only Credentials Path
            service = GmailHelper.GetGmailService(ApplicationName, GmailHelper.TokenPathType.WORKING_DIRECTORY, credentialsPath: credentialsPath);
            Assert.IsTrue(service.GetType() == typeof(Google.Apis.Gmail.v1.GmailService));
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GmailService_Dispose()
        {
            //Dispose (service as argument)
            var service = GmailHelper.GetGmailService(ApplicationName);
            Assert.IsNotNull(service);
            GmailHelper.DisposeGmailService(service);
            try
            {
                service.GetMessage(query: $"[from:{TestEmailId}][subject:'READ EMAIL']in:inbox is:read", markRead: true);
                Assert.Fail("No Object Disposed Exception Thrown");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (ObjectDisposedException ex)
            {
                Assert.IsTrue(ex.Message.Contains("Cannot access a disposed object."));
            }

            //Dispose (service as extension)
            service = GmailHelper.GetGmailService(ApplicationName);
            Assert.IsNotNull(service);
            service.DisposeGmailService();
            try
            {
                service.GetMessage(query: $"[from:{TestEmailId}][subject:'READ EMAIL']in:inbox is:read", markRead: true);
                Assert.Fail("No Object Disposed Exception Thrown");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (ObjectDisposedException ex)
            {
                Assert.IsTrue(ex.Message.Contains("Cannot access a disposed object."));
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessage()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetMessage(query: $"[from:{TestEmailId}][subject:'READ EMAIL']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessage_NoMatchingEmail()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetMessage(query: EmailDoesNotExistsSearchQuery);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessages()
        {
            var messages = GmailHelper.GetGmailService(ApplicationName)
                .GetMessages(query: $"[from:{TestEmailId}][subject:'EMAIL']in:inbox is:read", markRead: true);
            Assert.AreEqual(9, messages.Count);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessages_NoMatchingEmail()
        {
            var messages = GmailHelper.GetGmailService(ApplicationName)
                .GetMessages(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, messages.Count);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_PlainText()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: $"[from:{TestEmailId}][subject:'READ EMAIL WITH PLAIN TEXT (TEXT/PLAIN)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_HtmlText()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: $"[from:{TestEmailId}][subject:'READ EMAIL WITH HTML TEXT (TEXT/HTML)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_Multipart_NoText()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: $"[from:{TestEmailId}][subject:'READ EMAIL WITH NO TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_Multipart_PlainText()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: "[subject:'READ EMAIL WITH PLAIN TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_Multipart_HtmlText()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: "[subject:'READ EMAIL WITH HTML TEXT (MULTIPART/ALTERNATIVE)']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_MultipleMatchingEmails()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: $"[from:{TestEmailId}][subject:'READ EMAIL']in:inbox is:read", markRead: true);
            Assert.IsNotNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetLatestMessage_NoMatchingEmail()
        {
            var message = GmailHelper.GetGmailService(ApplicationName)
                .GetLatestMessage(query: EmailDoesNotExistsSearchQuery, markRead: true);
            Assert.IsNull(message);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessageAttachments()
        {
            var destPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Attach-Files";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Attach-Files";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Attach-Files";
            Directory.CreateDirectory(destPath);
            //EMAIL WITH ATTACHMENTS AND NO BODY
            var countOfMessageAttachmentsDownloaded = GmailHelper.GetGmailService(ApplicationName)
                .GetMessageAttachments(query: $"[from:{TestEmailId}][subject:'EMAIL WITH ATTACHMENTS AND NO BODY']in:inbox is:read", directoryPath: destPath);
            Assert.AreEqual(10, countOfMessageAttachmentsDownloaded);
            //EMAIL WITH ATTACHMENTS AND PLAIN TEXT BODY
            countOfMessageAttachmentsDownloaded = GmailHelper.GetGmailService(ApplicationName)
                .GetMessageAttachments(query: $"[from:{TestEmailId}][subject:'EMAIL WITH ATTACHMENTS AND PLAIN TEXT BODY']in:inbox is:read", directoryPath: destPath);
            Assert.AreEqual(10, countOfMessageAttachmentsDownloaded);
            //EMAIL WITH ATTACHMENTS AND HTML BODY
            countOfMessageAttachmentsDownloaded = GmailHelper.GetGmailService(ApplicationName)
                .GetMessageAttachments(query: $"[from:{TestEmailId}][subject:'EMAIL WITH ATTACHMENTS AND HTML BODY']in:inbox is:read", directoryPath: destPath);
            Assert.AreEqual(10, countOfMessageAttachmentsDownloaded);
            Directory.Delete(destPath, recursive: true);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessageAttachments_PathNotFound()
        {
            try
            {
                var destPath = "/home/user/attachments";
                GmailHelper.GetGmailService(ApplicationName)
                    .GetMessageAttachments(query: $"[from:{TestEmailId}][subject:'EMAIL WITH ATTACHMENTS AND NO BODY']in:inbox is:read", directoryPath: destPath);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (DirectoryNotFoundException ex) { Assert.AreEqual("Path - '/home/user/attachments' Not Found.", ex.Message); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessageAttachments_NoMatchingEmail()
        {
            var destPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Attachments";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Attachments";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Attachments";
            Directory.CreateDirectory(destPath);
            var countOfMessageAttachmentsDownloaded = GmailHelper.GetGmailService(ApplicationName)
                .GetMessageAttachments(query: EmailDoesNotExistsSearchQuery, directoryPath: destPath);
            Assert.AreEqual(0, countOfMessageAttachmentsDownloaded);
            Directory.Delete(destPath, recursive: true);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessagesAttachments()
        {
            var destPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Multi-Attach-Files";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Multi-Attach-Files";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Multi-Attach-Files";
            Directory.CreateDirectory(destPath);
            var messagesAttachmentsDownloaded = GmailHelper.GetGmailService(ApplicationName)
                .GetMessagesAttachments(query: $"[from:{TestEmailId}][subject:'EMAIL WITH ATTACHMENTS']in:inbox is:read", directoryPath: destPath);
            Assert.AreEqual(3, messagesAttachmentsDownloaded.Count);
            foreach (var messagesAttachments in messagesAttachmentsDownloaded)
                Assert.AreEqual(10, messagesAttachments.Value);
            Directory.Delete(destPath, recursive: true);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessagesAttachments_PathNotFound()
        {
            try
            {
                var destPath = "/home/user/attachments";
                GmailHelper.GetGmailService(ApplicationName)
                    .GetMessagesAttachments(query: $"[from:{TestEmailId}][subject:'EMAIL WITH ATTACHMENTS AND NO BODY']in:inbox is:read", directoryPath: destPath);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (DirectoryNotFoundException ex) { Assert.AreEqual("Path - '/home/user/attachments' Not Found.", ex.Message); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessagesAttachments_NoMatchingEmail()
        {
            var destPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Multi-Attachments";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Multi-Attachments";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                destPath = Environment.CurrentDirectory + "/" + "DotNet-Multi-Attachments";
            Directory.CreateDirectory(destPath);
            var messagesAttachmentsDownloaded = GmailHelper.GetGmailService(ApplicationName)
                .GetMessagesAttachments(query: EmailDoesNotExistsSearchQuery, directoryPath: destPath);
            Assert.AreEqual(0, messagesAttachmentsDownloaded.Count);
            Directory.Delete(destPath, recursive: true);
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
            Assert.IsFalse(Regex.IsMatch(body, "<(.|\n)*?>", RegexOptions.None, TimeSpan.FromMilliseconds(100)));
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: "EMAIL WITH PLAIN TEXT", body: body);
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
            Assert.IsTrue(Regex.IsMatch(body, "<(.|\n)*?>", RegexOptions.None, TimeSpan.FromMilliseconds(100)));
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.HTML, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: "EMAIL WITH HTML TEXT", body: body);
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
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, invalidEmailType);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (FormatException ex)
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
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: invalidEmailType);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (FormatException ex)
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
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, bcc: invalidEmailType);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (FormatException ex)
                {
                    var invalidEmail = invalidEmailType.Contains(",") ? invalidEmailType.Split(',')[1] : invalidEmailType;
                    Assert.AreEqual(string.Format("Not a valid 'Bcc' email address. Email: '{0}'", invalidEmail), ex.Message);
                }
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_SendMessage_Attachments_PlainText()
        {
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            Assert.IsFalse(Regex.IsMatch(body, "<(.|\n)*?>", RegexOptions.None, TimeSpan.FromMilliseconds(100)));
            var attachmentPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                attachmentPath = Environment.CurrentDirectory + "\\TestFiles\\Attachments\\";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                attachmentPath = Environment.CurrentDirectory + "/TestFiles/Attachments/";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                attachmentPath = Environment.CurrentDirectory + "/TestFiles/Attachments/";
            var attachments = new List<string>
            {
               attachmentPath + "Attachment.bmp",
               attachmentPath + "Attachment.docx",
               attachmentPath + "Attachment.gif",
               attachmentPath + "Attachment.jpg",
               attachmentPath + "Attachment.pdf",
               attachmentPath + "Attachment.png",
               attachmentPath + "Attachment.pptx",
               attachmentPath + "Attachment.tif",
               attachmentPath + "Attachment.txt",
               attachmentPath + "Attachment.wmv"
            };
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, attachments: attachments, cc: TestEmailId, bcc: TestEmailId, subject: "SEND DOTNETCORE EMAIL WITH PLAIN TEXT BODY AND ATTACHMENTS", body: body);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_SendMessage_Attachments_HtmlText()
        {
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\HTMLEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/HTMLEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/HTMLEmail.txt";
            var body = File.ReadAllText(path);
            Assert.IsTrue(Regex.IsMatch(body, "<(.|\n)*?>", RegexOptions.None, TimeSpan.FromMilliseconds(100)));
            var attachmentPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                attachmentPath = Environment.CurrentDirectory + "\\TestFiles\\Attachments\\";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                attachmentPath = Environment.CurrentDirectory + "/TestFiles/Attachments/";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                attachmentPath = Environment.CurrentDirectory + "/TestFiles/Attachments/";
            var attachments = new List<string>
            {
               attachmentPath + "Attachment.bmp",
               attachmentPath + "Attachment.docx",
               attachmentPath + "Attachment.gif",
               attachmentPath + "Attachment.jpg",
               attachmentPath + "Attachment.pdf",
               attachmentPath + "Attachment.png",
               attachmentPath + "Attachment.pptx",
               attachmentPath + "Attachment.tif",
               attachmentPath + "Attachment.txt",
               attachmentPath + "Attachment.wmv"
            };
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.HTML, TestEmailId, attachments: attachments, cc: TestEmailId, bcc: TestEmailId, subject: "SEND DOTNETCORE EMAIL WITH HTML TEXT BODY AND ATTACHMENTS", body: body);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_SendMessage_Attachments_InvalidToEmail()
        {
            var invalidEmailTypes = new string[] { "testgmail.com", "test@gmailcom" , "1test@gmail.com",
                "test@gmail.com,testgmail.com", "test@gmail.com,test@gmailcom", "test@gmail.com,1test@gmail.com" };
            foreach (var invalidEmailType in invalidEmailTypes)
            {
                try
                {
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, invalidEmailType, attachments: null);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (FormatException ex)
                {
                    var invalidEmail = invalidEmailType.Contains(",") ? invalidEmailType.Split(',')[1] : invalidEmailType;
                    Assert.AreEqual(string.Format("Not a valid 'To' email address. Email: '{0}'", invalidEmail), ex.Message);
                }
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_SendMessage_Attachments_InvalidCcEmail()
        {
            var invalidEmailTypes = new string[] { "testgmail.com", "test@gmailcom" , "1test@gmail.com",
                "test@gmail.com,testgmail.com", "test@gmail.com,test@gmailcom", "test@gmail.com,1test@gmail.com" };
            foreach (var invalidEmailType in invalidEmailTypes)
            {
                try
                {
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, attachments: null, cc: invalidEmailType);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (FormatException ex)
                {
                    var invalidEmail = invalidEmailType.Contains(",") ? invalidEmailType.Split(',')[1] : invalidEmailType;
                    Assert.AreEqual(string.Format("Not a valid 'Cc' email address. Email: '{0}'", invalidEmail), ex.Message);
                }
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_SendMessage_Attachments_InvalidBccEmail()
        {
            var invalidEmailTypes = new string[] { "testgmail.com", "test@gmailcom" , "1test@gmail.com",
                "test@gmail.com,testgmail.com", "test@gmail.com,test@gmailcom", "test@gmail.com,1test@gmail.com" };
            foreach (var invalidEmailType in invalidEmailTypes)
            {
                try
                {
                    GmailHelper.GetGmailService(ApplicationName).SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, attachments: null, bcc: invalidEmailType);
                    Assert.Fail(string.Format("No Invalid Email Exception Thrown. Email Id - '{0}'.", invalidEmailType));
                }
                catch (AssertFailedException ex) { throw ex; }
                catch (FormatException ex)
                {
                    var invalidEmail = invalidEmailType.Contains(",") ? invalidEmailType.Split(',')[1] : invalidEmailType;
                    Assert.AreEqual(string.Format("Not a valid 'Bcc' email address. Email: '{0}'", invalidEmail), ex.Message);
                }
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_SendMessage_Attachments_AttachmentFileNotFound()
        {
            var attachmentPath = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                attachmentPath = Environment.CurrentDirectory + "\\TestFiles\\Attachments\\";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                attachmentPath = Environment.CurrentDirectory + "/TestFiles/Attachments/";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                attachmentPath = Environment.CurrentDirectory + "/TestFiles/Attachments/";
            var attachments = new List<string>
            {
               attachmentPath + "Attachment.bmp",
               attachmentPath + "Attachment.docx",
               attachmentPath + "Attachment.gif",
               attachmentPath + "Attachment.jpg",
               attachmentPath + "Attachment.pdf",
               attachmentPath + "Attachment.png",
               attachmentPath + "Attachment.pptx",
               attachmentPath + "Attachment.tif",
               attachmentPath + "Attachment.txt",
               attachmentPath
            };
            try
            {
                GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, attachments: attachments, cc: TestEmailId, bcc: TestEmailId, subject: "ATTACHMENT FILE NOT FOUND", body: "Hello");
                Assert.Fail("No 'FileNotFoundException' Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (FileNotFoundException ex)
            {
                Assert.AreEqual(string.Format("Attachment file '{0}' not found.", attachmentPath), ex.Message);
            }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MoveMessageToTrash()
        {
            //Test Data
            var subject = "MOVE DOTNETCORE MESSAGE TO TRASH " + Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);

            //Test Run
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessageToTrash(query: $"[from:{TestEmailId}][subject:'MOVE DOTNETCORE MESSAGE TO TRASH " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MoveMessageToTrash_NoMatchingEmail()
        {
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessageToTrash(query: EmailDoesNotExistsSearchQuery);
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
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: $"[from:{TestEmailId}][subject:'MOVE DOTNETCORE MESSAGES TO TRASH " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MoveMessagesToTrash_NoMatchingEmail()
        {
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesMovedToTrash);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UntrashMessage()
        {
            //Test Data
            var subject = "UNTRASH DOTNETCORE MESSAGE " + Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            var isMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessageToTrash(query: $"[from:{TestEmailId}][subject:'UNTRASH DOTNETCORE MESSAGE  " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMovedToTrash);

            //Test Run
            var isUntrashed = GmailHelper.GetGmailService(ApplicationName)
                .UntrashMessage(query: $"[from:{TestEmailId}][subject:'UNTRASH DOTNETCORE MESSAGE  " + subject + "']in:trash is:unread");
            Assert.IsTrue(isUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UntrashMessage_NoMatchingEmail()
        {
            var isUntrashed = GmailHelper.GetGmailService(ApplicationName)
                .UntrashMessage(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UntrashMessages()
        {
            //Test Data
            var subject = "UNTRASH DOTNETCORE MESSAGES " + Guid.NewGuid().ToString();
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
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            }
            var countOfMessagesMovedToTrash = GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: $"[from:{TestEmailId}][subject:'UNTRASH DOTNETCORE MESSAGES " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMovedToTrash);

            //Test Run
            var countOfMessagesUntrashed = GmailHelper.GetGmailService(ApplicationName)
                .UntrashMessages(query: $"[from:{TestEmailId}][subject:'UNTRASH DOTNETCORE MESSAGES " + subject + "']in:trash is:unread");
            Assert.AreEqual(2, countOfMessagesUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UntrashMessages_NoMatchingEmail()
        {
            var countOfMessagesUntrashed = GmailHelper.GetGmailService(ApplicationName)
                .UntrashMessages(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesUntrashed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ReportSpamMessage()
        {
            //Test Data
            var subject = "REPORT DOTNETCORE SPAM MESSAGE " + Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);

            //Test Run
            var isSpamReported = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessage(query: $"[from:{TestEmailId}][subject:'REPORT DOTNETCORE SPAM MESSAGE " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isSpamReported);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ReportSpamMessage_NoMatchingEmail()
        {
            var isSpamReported = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessage(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isSpamReported);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ReportSpamMessages()
        {
            //Test Data
            var subject = "REPORT DOTNETCORE SPAM MESSAGES " + Guid.NewGuid().ToString();
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
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesReportedAsSpam = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessages(query: $"[from:{TestEmailId}][subject:'REPORT DOTNETCORE SPAM MESSAGES " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesReportedAsSpam);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ReportSpamMessages_NoMatchingEmail()
        {
            var countOfMessagesReportedAsSpam = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessages(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesReportedAsSpam);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UnspamMessage()
        {
            //Test Data
            var subject = "UNSPAM DOTNETCORE MESSAGE " + Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            var isSpamReported = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessage(query: $"[from:{TestEmailId}][subject:'UNSPAM DOTNETCORE MESSAGE " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isSpamReported);

            //Test Run
            var isUnspamed = GmailHelper.GetGmailService(ApplicationName)
                .UnspamMessage(query: $"[from:{TestEmailId}][subject:'UNSPAM DOTNETCORE MESSAGE " + subject + "']in:spam is:unread");
            Assert.IsTrue(isUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UnspamMessage_NoMatchingEmail()
        {
            var isUnspamed = GmailHelper.GetGmailService(ApplicationName)
                .UnspamMessage(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UnspamMessages()
        {
            //Test Data
            var subject = "UNSPAM DOTNETCORE MESSAGES " + Guid.NewGuid().ToString();
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
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            }
            var countOfMessagesReportedAsSpam = GmailHelper.GetGmailService(ApplicationName)
                .ReportSpamMessages(query: $"[from:{TestEmailId}][subject:'UNSPAM DOTNETCORE MESSAGES " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesReportedAsSpam);

            //Test Run
            var countOfMessagesUnspamed = GmailHelper.GetGmailService(ApplicationName)
                .UnspamMessages(query: $"[from:{TestEmailId}][subject:'UNSPAM DOTNETCORE MESSAGES " + subject + "']in:spam is:unread");
            Assert.AreEqual(2, countOfMessagesUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UnspamMessages_NoMatchingEmail()
        {
            var countOfMessagesUnspamed = GmailHelper.GetGmailService(ApplicationName)
                .UnspamMessages(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesUnspamed);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MarkMessageAsRead()
        {
            //Test Data
            var subject = "MARK DOTNETCORE MESSAGE AS READ " + Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);

            //Test Run
            var isMarkedRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsRead(query: $"[from:{TestEmailId}][subject:'MARK DOTNETCORE MESSAGE AS READ  " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMarkedRead);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MarkMessageAsRead_NoMatchingEmail()
        {
            var isMarkedRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsRead(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isMarkedRead);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MarkMessagesAsRead()
        {
            //Test Data
            var subject = "MARK DOTNETCORE MESSAGES AS READ " + Guid.NewGuid().ToString();
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
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesMarkedAsRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsRead(query: $"[from:{TestEmailId}][subject:'MARK DOTNETCORE MESSAGES AS READ " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMarkedAsRead);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MarkMessagesAsRead_NoMatchingEmail()
        {
            var countOfMessagesMarkedAsRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsRead(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesMarkedAsRead);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MarkMessageAsUnread()
        {
            //Test Data
            var subject = "MARK DOTNETCORE MESSAGE AS UNREAD " + Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            var isMarkedRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsRead(query: $"[from:{TestEmailId}][subject:'MARK DOTNETCORE MESSAGE AS UNREAD  " + subject + "']in:inbox is:unread");
            Assert.IsTrue(isMarkedRead);

            //Test Run
            var isMarkedUnread = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsUnread(query: $"[from:{TestEmailId}][subject:'MARK DOTNETCORE MESSAGE AS UNREAD  " + subject + "']in:inbox is:read");
            Assert.IsTrue(isMarkedUnread);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MarkMessageAsUnread_NoMatchingEmail()
        {
            var isMarkedUnread = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessageAsUnread(query: EmailDoesNotExistsSearchQuery);
            Assert.IsFalse(isMarkedUnread);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MarkMessagesAsUnread()
        {
            //Test Data
            var subject = "MARK DOTNETCORE MESSAGES AS UNREAD " + Guid.NewGuid().ToString();
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
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            }
            var countOfMessagesMarkedAsRead = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsRead(query: $"[from:{TestEmailId}][subject:'MARK DOTNETCORE MESSAGES AS UNREAD " + subject + "']in:inbox is:unread");
            Assert.AreEqual(2, countOfMessagesMarkedAsRead);

            //Test Run
            var countOfMessagesMarkedAsUnread = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsUnread(query: $"[from:{TestEmailId}][subject:'MARK DOTNETCORE MESSAGES AS UNREAD " + subject + "']in:inbox is:read");
            Assert.AreEqual(2, countOfMessagesMarkedAsUnread);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_MarkMessagesAsUnread_NoMatchingEmail()
        {
            var countOfMessagesMarkedAsUnread = GmailHelper.GetGmailService(ApplicationName)
                .MarkMessagesAsUnread(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, countOfMessagesMarkedAsUnread);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ModifyMessage()
        {
            //Test Data
            var subject = "MODIFY DOTNETCORE MESSAGE " + Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);

            //Test Run
            var isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: $"[from:{TestEmailId}][subject:'MODIFY DOTNETCORE MESSAGE " + subject + "']in:inbox", labelsToAdd: new List<string>() { "IMPORTANT", "SPAM", });
            Assert.IsTrue(isModified);
            isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: $"[from:{TestEmailId}][subject:'MODIFY DOTNETCORE MESSAGE " + subject + "']in:spam", labelsToRemove: new List<string>() { "IMPORTANT", "UNREAD" });
            Assert.IsTrue(isModified);
            isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: $"[from:{TestEmailId}][subject:'MODIFY DOTNETCORE MESSAGE " + subject + "']in:spam", labelsToAdd: new List<string>() { "INBOX", "STARRED", "UNREAD", }, labelsToRemove: new List<string>() { "SPAM" });
            Assert.IsTrue(isModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ModifyMessage_NoLabelsSupplied()
        {
            try
            {
                GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: EmailDoesNotExistsSearchQuery);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (ArgumentException ex) { Assert.AreEqual("Parameters 'labelsToAdd' / 'labelsToRemove' required.", ex.Message); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ModifyMessage_NoMatchingEmail()
        {
            var isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: EmailDoesNotExistsSearchQuery, labelsToAdd: new List<string>() { "STARRED", "IMPORTANT", }, labelsToRemove: new List<string>() { "UNREAD" });
            Assert.IsFalse(isModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ModifyMessages()
        {
            //Test Data
            var subject = "MODIFY DOTNETCORE MESSAGES " + Guid.NewGuid().ToString();
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
                GmailHelper.GetGmailService(ApplicationName)
                    .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            }

            //Test Run
            var countOfMessagesModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: $"[from:{TestEmailId}][subject:'MODIFY DOTNETCORE MESSAGES " + subject + "']in:inbox", labelsToAdd: new List<string>() { "IMPORTANT", "SPAM", });
            Assert.AreEqual(2, countOfMessagesModified);
            countOfMessagesModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: $"[from:{TestEmailId}][subject:'MODIFY DOTNETCORE MESSAGES " + subject + "']in:spam", labelsToRemove: new List<string>() { "IMPORTANT", "UNREAD" });
            Assert.AreEqual(2, countOfMessagesModified);
            countOfMessagesModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: $"[from:{TestEmailId}][subject:'MODIFY DOTNETCORE MESSAGES " + subject + "']in:spam", labelsToAdd: new List<string>() { "INBOX", "STARRED", "UNREAD", }, labelsToRemove: new List<string>() { "SPAM" });
            Assert.AreEqual(2, countOfMessagesModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ModifyMessages_NoLabelsSupplied()
        {
            try
            {
                GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: EmailDoesNotExistsSearchQuery);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (ArgumentException ex) { Assert.AreEqual("Parameters 'labelsToAdd' / 'labelsToRemove' required.", ex.Message); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ModifyMessages_NoMatchingEmail()
        {
            var countOfMessagesModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessages(query: EmailDoesNotExistsSearchQuery, labelsToAdd: new List<string>() { "STARRED", "IMPORTANT", });
            Assert.AreEqual(0, countOfMessagesModified);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessageLabels()
        {
            //Test Data
            var subject = "GET DOTNETCORE MESSAGE LABELS " + Guid.NewGuid().ToString();
            var path = "";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                path = Environment.CurrentDirectory + "\\TestFiles\\PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                path = Environment.CurrentDirectory + "/TestFiles/PlainEmail.txt";
            var body = File.ReadAllText(path);
            GmailHelper.GetGmailService(ApplicationName)
                .SendMessage(GmailHelper.EmailContentType.PLAIN, TestEmailId, cc: TestEmailId, bcc: TestEmailId, subject: subject, body: body);
            var isModified = GmailHelper.GetGmailService(ApplicationName)
                .ModifyMessage(query: $"[from:{TestEmailId}][subject:'GET DOTNETCORE MESSAGE LABELS " + subject + "']in:inbox", labelsToAdd: new List<string>() { "IMPORTANT", "STARRED", });
            Assert.IsTrue(isModified);

            //Test Run
            var labels = GmailHelper.GetGmailService(ApplicationName)
                .GetMessageLabels(query: "GET DOTNETCORE MESSAGE LABELS " + subject + "']in:inbox");
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
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_GetMessageLabels_NoMatchingEmail()
        {
            var labels = GmailHelper.GetGmailService(ApplicationName)
                .GetMessageLabels(query: EmailDoesNotExistsSearchQuery);
            Assert.AreEqual(0, labels.Count);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_CreateUserLabel()
        {
            //Test Run
            //Argument set 1
            var labelName1 = "L-CORE-1 " + Guid.NewGuid().ToString();
            var userLabel = GmailHelper.GetGmailService(ApplicationName)
                .CreateUserLabel(labelName1);
            Assert.IsFalse(string.IsNullOrEmpty(userLabel.Id));
            Assert.AreEqual("labelShow", userLabel.LabelListVisibility);
            Assert.AreEqual("show", userLabel.MessageListVisibility);
            Assert.AreEqual(labelName1, userLabel.Name);
            Assert.AreEqual(null, userLabel.Type);
            Assert.AreEqual("#666666", userLabel.Color.BackgroundColor);
            Assert.AreEqual("#ffffff", userLabel.Color.TextColor);
            //Argument set 2
            var labelName2 = "L-CORE-2 " + Guid.NewGuid().ToString();
            userLabel = GmailHelper.GetGmailService(ApplicationName)
                .CreateUserLabel(labelName2, labelListVisibility: GmailHelper.LabelListVisibility.LABEL_HIDE, messageListVisibility: GmailHelper.MessageListVisibility.HIDE);
            Assert.IsFalse(string.IsNullOrEmpty(userLabel.Id));
            Assert.AreEqual("labelHide", userLabel.LabelListVisibility);
            Assert.AreEqual("hide", userLabel.MessageListVisibility);
            Assert.AreEqual(labelName2, userLabel.Name);
            Assert.AreEqual(null, userLabel.Type);
            Assert.AreEqual("#666666", userLabel.Color.BackgroundColor);
            Assert.AreEqual("#ffffff", userLabel.Color.TextColor);
            //Argument set 3
            var labelName3 = "L-CORE-3 " + Guid.NewGuid().ToString();
            userLabel = GmailHelper.GetGmailService(ApplicationName)
                .CreateUserLabel(labelName3, labelListVisibility: GmailHelper.LabelListVisibility.LABEL_SHOW_IF_UNREAD);
            Assert.IsFalse(string.IsNullOrEmpty(userLabel.Id));
            Assert.AreEqual("labelShowIfUnread", userLabel.LabelListVisibility);
            Assert.AreEqual("show", userLabel.MessageListVisibility);
            Assert.AreEqual(labelName3, userLabel.Name);
            Assert.AreEqual(null, userLabel.Type);
            Assert.AreEqual("#666666", userLabel.Color.BackgroundColor);
            Assert.AreEqual("#ffffff", userLabel.Color.TextColor);
            //Argument set 4
            var labelName4  = "L-CORE-4 " + Guid.NewGuid().ToString();
            userLabel = GmailHelper.GetGmailService(ApplicationName)
                .CreateUserLabel(labelName4, labelBackgroundColor: "#ff7537", labelTextColor: "#1c4587");
            Assert.IsFalse(string.IsNullOrEmpty(userLabel.Id));
            Assert.AreEqual("labelShow", userLabel.LabelListVisibility);
            Assert.AreEqual("show", userLabel.MessageListVisibility);
            Assert.AreEqual(labelName4, userLabel.Name);
            Assert.AreEqual(null, userLabel.Type);
            Assert.AreEqual("#ff7537", userLabel.Color.BackgroundColor);
            Assert.AreEqual("#1c4587", userLabel.Color.TextColor);

            //Test Cleanup
            var isDeleted = GmailHelper.GetGmailService(ApplicationName).DeleteLabel(labelName1);
            Assert.IsTrue(isDeleted);
            isDeleted = GmailHelper.GetGmailService(ApplicationName).DeleteLabel(labelName2);
            Assert.IsTrue(isDeleted);
            isDeleted = GmailHelper.GetGmailService(ApplicationName).DeleteLabel(labelName3);
            Assert.IsTrue(isDeleted);
            isDeleted = GmailHelper.GetGmailService(ApplicationName).DeleteLabel(labelName4);
            Assert.IsTrue(isDeleted);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UpdateUserLabel()
        {
            //Test Data
            var originalLabelName = "L-CORE-" + Guid.NewGuid().ToString();
            var userLabel = GmailHelper.GetGmailService(ApplicationName)
                .CreateUserLabel(originalLabelName);
            Assert.IsFalse(string.IsNullOrEmpty(userLabel.Id));
            Assert.AreEqual("labelShow", userLabel.LabelListVisibility);
            Assert.AreEqual("show", userLabel.MessageListVisibility);
            Assert.AreEqual(originalLabelName, userLabel.Name);
            Assert.AreEqual(null, userLabel.Type);
            Assert.AreEqual("#666666", userLabel.Color.BackgroundColor);
            Assert.AreEqual("#ffffff", userLabel.Color.TextColor);

            //Test Run
            //Argument set 1
            var newLabelName1 = originalLabelName;
            var updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(originalLabelName, newLabelName1);
            Assert.IsNotNull(updatedLabel);
            Assert.IsFalse(string.IsNullOrEmpty(updatedLabel.Id));
            Assert.AreEqual("labelShow", updatedLabel.LabelListVisibility);
            Assert.AreEqual("show", updatedLabel.MessageListVisibility);
            Assert.AreEqual(newLabelName1, updatedLabel.Name);
            Assert.AreEqual("user", updatedLabel.Type);
            Assert.AreEqual("#666666", updatedLabel.Color.BackgroundColor);
            Assert.AreEqual("#ffffff", updatedLabel.Color.TextColor);
            //Argument set 2
            var newLabelName2 = "L-CORE-2 " + Guid.NewGuid().ToString();
            updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(newLabelName1, newLabelName2, labelListVisibility: GmailHelper.LabelListVisibility.LABEL_HIDE, messageListVisibility: GmailHelper.MessageListVisibility.HIDE);
            Assert.IsNotNull(updatedLabel);
            Assert.IsFalse(string.IsNullOrEmpty(updatedLabel.Id));
            Assert.AreEqual("labelHide", updatedLabel.LabelListVisibility);
            Assert.AreEqual("hide", updatedLabel.MessageListVisibility);
            Assert.AreEqual(newLabelName2, updatedLabel.Name);
            Assert.AreEqual("user", updatedLabel.Type);
            Assert.AreEqual("#666666", updatedLabel.Color.BackgroundColor);
            Assert.AreEqual("#ffffff", updatedLabel.Color.TextColor);
            //Argument set 3
            var newLabelName3 = "L-CORE-3 " + Guid.NewGuid().ToString();
            updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(newLabelName2, newLabelName3, labelListVisibility: GmailHelper.LabelListVisibility.LABEL_SHOW_IF_UNREAD);
            Assert.IsNotNull(updatedLabel);
            Assert.IsFalse(string.IsNullOrEmpty(updatedLabel.Id));
            Assert.AreEqual("labelShowIfUnread", updatedLabel.LabelListVisibility);
            Assert.AreEqual("show", updatedLabel.MessageListVisibility);
            Assert.AreEqual(newLabelName3, updatedLabel.Name);
            Assert.AreEqual("user", updatedLabel.Type);
            Assert.AreEqual("#666666", updatedLabel.Color.BackgroundColor);
            Assert.AreEqual("#ffffff", updatedLabel.Color.TextColor);
            //Argument set 4
            var newLabelName4 = "L-CORE-4 " + Guid.NewGuid().ToString();
            updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(newLabelName3, newLabelName4, labelBackgroundColor: "#ff7537", labelTextColor: "#1c4587");
            Assert.IsNotNull(updatedLabel);
            Assert.IsFalse(string.IsNullOrEmpty(updatedLabel.Id));
            Assert.AreEqual("labelShow", updatedLabel.LabelListVisibility);
            Assert.AreEqual("show", updatedLabel.MessageListVisibility);
            Assert.AreEqual(newLabelName4, updatedLabel.Name);
            Assert.AreEqual("user", updatedLabel.Type);
            Assert.AreEqual("#ff7537", updatedLabel.Color.BackgroundColor);
            Assert.AreEqual("#1c4587", updatedLabel.Color.TextColor);

            //Test Cleanup
            var isDeleted = GmailHelper.GetGmailService(ApplicationName)
                .DeleteLabel(newLabelName4);
            Assert.IsTrue(isDeleted);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UpdateUserLabel_UpdateSystemLabel()
        {
            var originalLabelName = "INBOX";
            var newLabelName = originalLabelName;
            var updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(originalLabelName, newLabelName);
            Assert.IsNull(updatedLabel);
            originalLabelName = "DRAFT";
            newLabelName = originalLabelName;
            updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(originalLabelName, newLabelName);
            Assert.IsNull(updatedLabel);
            originalLabelName = "SENT";
            newLabelName = originalLabelName;
            updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(originalLabelName, newLabelName);
            Assert.IsNull(updatedLabel);
            originalLabelName = "SPAM";
            newLabelName = originalLabelName;
            updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(originalLabelName, newLabelName);
            Assert.IsNull(updatedLabel);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_UpdateUserLabel_NoMatchingLabel()
        {
            var originalLabelName = "LABEL-DOES-NOT-EXISTS";
            var newLabelName = "RENAME-LABEL";
            var updatedLabel = GmailHelper.GetGmailService(ApplicationName)
                .UpdateUserLabel(originalLabelName, newLabelName);
            Assert.IsNull(updatedLabel);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_DeleteLabel()
        {
            //Test Data
            var labelName = "L-CORE-" + Guid.NewGuid().ToString();
            var userLabel = GmailHelper.GetGmailService(ApplicationName)
                .CreateUserLabel(labelName);
            Assert.IsFalse(string.IsNullOrEmpty(userLabel.Id));
            Assert.AreEqual("labelShow", userLabel.LabelListVisibility);
            Assert.AreEqual("show", userLabel.MessageListVisibility);
            Assert.AreEqual(labelName, userLabel.Name);
            Assert.AreEqual(null, userLabel.Type);
            Assert.AreEqual("#666666", userLabel.Color.BackgroundColor);
            Assert.AreEqual("#ffffff", userLabel.Color.TextColor);

            //Test Run
            var isDeleted = GmailHelper.GetGmailService(ApplicationName)
                .DeleteLabel(labelName);
            Assert.IsTrue(isDeleted);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_DeleteLabel_DeleteSystemLabel()
        {
            var labelName = "INBOX";
            try
            {
                GmailHelper.GetGmailService(ApplicationName).DeleteLabel(labelName);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (GoogleApiException ex) { Assert.IsTrue(ex.Message.Contains("BadRequest. Invalid delete request")); }
            labelName = "DRAFT";
            try
            {
                GmailHelper.GetGmailService(ApplicationName).DeleteLabel(labelName);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (GoogleApiException ex) { Assert.IsTrue(ex.Message.Contains("BadRequest. Invalid delete request")); }
            labelName = "SENT";
            try
            {
                GmailHelper.GetGmailService(ApplicationName).DeleteLabel(labelName);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (GoogleApiException ex) { Assert.IsTrue(ex.Message.Contains("BadRequest. Invalid delete request")); }
            labelName = "SPAM";
            try
            {
                GmailHelper.GetGmailService(ApplicationName).DeleteLabel(labelName);
                Assert.Fail("No Exception Thrown.");
            }
            catch (AssertFailedException ex) { throw ex; }
            catch (GoogleApiException ex) { Assert.IsTrue(ex.Message.Contains("BadRequest. Invalid delete request")); }
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_DeleteLabel_NoMatchingLabel()
        {
            var labelName = "LABEL-DOES-NOT-EXISTS";
            var isDeleted = GmailHelper.GetGmailService(ApplicationName)
                .DeleteLabel(labelName);
            Assert.IsFalse(isDeleted);
        }

        [TestMethod]
        [TestCategory("GMAIL-TESTS-DOTNETCORE")]
        public void Test_ListUserLabels()
        {
            var userLabels = GmailHelper.GetGmailService(ApplicationName).ListUserLabels();
            Assert.IsFalse(userLabels.Count == 0);
            var systemLabels = userLabels.Where(x => x.Type.Equals("system"));
            Assert.IsTrue(systemLabels.Count() >= 14);
            var inboxLabel = userLabels.FirstOrDefault(x => x.Name.Equals("INBOX"));
            Assert.IsNotNull(inboxLabel);
            Assert.AreEqual("INBOX", inboxLabel.Id);
            Assert.AreEqual(null, inboxLabel.LabelListVisibility);
            Assert.AreEqual(null, inboxLabel.MessageListVisibility);
            Assert.AreEqual("INBOX", inboxLabel.Name);
            Assert.AreEqual("system", inboxLabel.Type);
        }

        [TestMethod]
        [TestCategory("TestCleanup")]
        public void Inbox_CleanUp()
        {
            var mesagesMoved = 0;
            mesagesMoved += GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: $"[from:{TestEmailId}]in:inbox is:unread");
            mesagesMoved += GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: $"[from:{TestEmailId}]in:spam is:unread");
            mesagesMoved += GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[subject:'MARK DOTNETCORE MESSAGE AS READ']in:inbox is:read");
            mesagesMoved += GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[subject:'MARK DOTNETFRAMEWORK MESSAGE AS READ']in:inbox is:read");
            mesagesMoved += GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[subject:'MARK DOTNETCORE MESSAGES AS READ']in:inbox is:read");
            mesagesMoved += GmailHelper.GetGmailService(ApplicationName)
                .MoveMessagesToTrash(query: "[subject:'MARK DOTNETFRAMEWORK MESSAGES AS READ']in:inbox is:read");
            Assert.IsTrue(mesagesMoved >= 0);
        }
    }
}