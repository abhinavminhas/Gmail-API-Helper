﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using MessagePart = Google.Apis.Gmail.v1.Data.MessagePart;

namespace GmailAPIHelper
{
    /// <summary>
    /// Gmail Helper.
    /// </summary>
    public static class GmailHelper
    {
        private const string _tokenFile = "token.json";
        private const string _labelUnread = "UNREAD";
        private const string _labelInbox = "INBOX";
        private const string _labelSpam = "SPAM";

        /// <summary>
        /// 'Token Path Type' enum.
        /// 'HOME' - Home Directory.
        /// 'WORKING_DIRECTORY' - Working Directory.
        /// 'CUSTOM' - Custom Path.
        /// </summary>
        public enum TokenPathType
        {
            HOME = 1,
            WORKING_DIRECTORY = 2,
            CUSTOM = 3
        }

        /// <summary>
        /// 'Email Content Type' enum.
        /// 'PLAIN' - 'text/plain'.
        /// 'HTML' - 'text/html'.
        /// </summary>
        public enum EmailContentType
        {
            PLAIN = 1,
            HTML = 2
        }

        /// <summary>
        /// 'Label List Visibility' enum.
        /// 'LABEL_SHOW' - 'Show the label in the label list'.
        /// 'LABEL_SHOW_IF_UNREAD' - 'Show the label if there are any unread messages with that label'.
        /// 'LABEL_HIDE' - 'Do not show the label in the label list'.
        /// </summary>
        public enum LabelListVisibility
        {
            LABEL_SHOW = 1,
            LABEL_SHOW_IF_UNREAD = 2,
            LABEL_HIDE = 3
        }

        /// <summary>
        /// 'Message List Visibility' enum.
        /// 'SHOW' - 'Show the label in the message list'.
        /// 'HIDE' - 'Do not show the label in the message list'.
        /// </summary>
        public enum MessageListVisibility
        {
            SHOW = 1,
            HIDE = 2
        }

        /// <summary>
        /// Sets the token file path to be used.
        /// </summary>
        /// <param name="tokenPathType">'TokenPathType' enum value. 'HOME' for users home directory, 'WORKING_DIRECTORY' for working directory, 'CUSTOM' for any other custom path to be used.</param>
        /// <param name="tokenPath">Token file path value in case of 'TokenPathType - CUSTOM' value.</param>
        /// <returns>Token file path.</returns>
        /// <exception cref="NotImplementedException">Throws - 'NotImplementedException' for OS Platforms other than Windows/Linux/OSX.</exception>
        private static string SetTokenPath(TokenPathType tokenPathType, string tokenPath = "")
        {
            string filePath = "";
            if (tokenPathType == TokenPathType.HOME)
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    filePath = Path.Combine(Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%"), _tokenFile);
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                    filePath = Path.Combine(Environment.GetEnvironmentVariable("HOME"), _tokenFile);
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                    filePath = Path.Combine(Environment.GetEnvironmentVariable("HOME"), _tokenFile);
                else
                    throw new NotImplementedException("OS Platform: Not 'Windows/Linux/OSX' Platform.");
            }
            else if (tokenPathType == TokenPathType.WORKING_DIRECTORY)
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    filePath = Path.Combine(Environment.CurrentDirectory, _tokenFile);
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                    filePath = Path.Combine(Environment.CurrentDirectory, _tokenFile);
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                    filePath = Path.Combine(Environment.CurrentDirectory, _tokenFile);
                else
                    throw new NotImplementedException("OS Platform: Not 'Windows/Linux/OSX' Platform.");
            }
            else if (tokenPathType == TokenPathType.CUSTOM)
                filePath = tokenPath;
            return filePath;
        }

        /// <summary>
        /// Sets the 'credentials.json' file path to be used.
        /// </summary>
        /// <param name="credentialsPath">'credentials.json' file path.</param>
        /// <returns>'credentials.json' file path.</returns>
        private static string SetCredentialsPath(string credentialsPath)
        {
            string file = "credentials.json";
            if (string.IsNullOrEmpty(credentialsPath))
                return file;
            else
                return credentialsPath;
        }

        /// <summary>
        /// Connects to the 'Gmail Service'.
        /// </summary>
        /// <param name="applicationName">'Application Name' value created in Gmail API Console.</param>
        /// <param name="tokenPathType">'TokenPathType' enum value. 'HOME' for users home directory, 'WORKING_DIRECTORY' for working directory, 'CUSTOM' for any other custom path to be used.
        /// Default value - 'WORKING_DIRECTORY'.</param>
        /// <param name="tokenPath">'token.json' path to save generated token from gmail authentication/authorization. 
        /// Always asks in case of change in gmail authentication or valid token file missing in the given path. Default path is blank, required for 'TokenPathType - CUSTOM'.</param>
        /// <param name="credentialsPath">'credentials.json' file path. Default path is blank in which case uses working directory for 'credentials.json' file lookup.
        /// <returns>Gmail Service.</returns>
        public static GmailService GetGmailService(string applicationName, TokenPathType tokenPathType = TokenPathType.WORKING_DIRECTORY, string tokenPath = "", string credentialsPath = "")
        {
            var scopes = new List<string>
            {
                GmailService.Scope.GmailMetadata,
                GmailService.Scope.GmailReadonly,
                GmailService.Scope.GmailModify,
                GmailService.Scope.GmailLabels,
                GmailService.Scope.GmailSend
            };
            UserCredential credential;
            var credentials = SetCredentialsPath(credentialsPath);
            using (var stream = new FileStream(credentials, FileMode.Open, FileAccess.Read))
            {
                //The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
                string credPath = SetTokenPath(tokenPathType, tokenPath);
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }
            //Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            });
            return service;
        }

        /// <summary>
        /// Disposes Gmail service.
        /// </summary>
        /// <param name="gmailService">'Gmail' service instance value.</param>
        public static void DisposeGmailService(this GmailService gmailService)
        {
            gmailService.Dispose();
        }

        /// <summary>
        /// Gets Gmail latest complete message with metadata for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="markRead">Boolean value to mark retrieved latest message as read. Default - 'false'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Email message matching the search criteria.</returns>
        public static Message GetMessage(this GmailService gmailService, string query, bool markRead = false, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var requiredLatestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (requiredLatestMessage != null)
                {
                    var messageRequest = service.Users.Messages.Get(userId, requiredLatestMessage.Id);
                    messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Full;
                    requiredLatestMessage = messageRequest.Execute();
                    if (markRead)
                    {
                        var labelToRemove = new List<string> { _labelUnread };
                        service.RemoveLabels(requiredLatestMessage.Id, labelToRemove, userId: userId);
                    }
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return requiredLatestMessage;
            }
            else
            {
                if (disposeGmailService)
                    service.DisposeGmailService();
                return null;
            }
        }

        /// <summary>
        /// Gets Gmail messages with metadata for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="markRead">Boolean value to mark retrieved messages as read. Default - 'false'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>List of email messages matching the search criteria.</returns>
        public static List<Message> GetMessages(this GmailService gmailService, string query, bool markRead = false, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Full;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
                if (markRead)
                {
                    var labelToRemove = new List<string> { _labelUnread };
                    service.RemoveLabels(message.Id, labelToRemove, userId: userId);
                }
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return messages;
        }

        /// <summary>
        /// Gets Gmail latest message body text for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="markRead">Boolean value to mark retrieved latest email as read. Default - 'false'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Email message body in 'text/plain' format.</returns>
        public static string GetLatestMessage(this GmailService gmailService, string query, bool markRead = false, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                string requiredMessage = null;
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var messageRequest = service.Users.Messages.Get(userId, latestMessage.Id);
                    messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Full;
                    var latestMessageDetails = messageRequest.Execute();
                    MessagePart requiredMessagePart = null;
                    if (latestMessageDetails.Payload.MimeType == "text/plain")
                        requiredMessagePart = latestMessageDetails.Payload;
                    else if (latestMessageDetails.Payload.MimeType == "text/html")
                        requiredMessagePart = latestMessageDetails.Payload;
                    else
                    {
                        if (latestMessageDetails.Payload.Parts != null)
                        {
                            requiredMessagePart = latestMessageDetails.Payload.Parts.FirstOrDefault(x => x.MimeType == "text/plain");
                            if (requiredMessagePart == null || requiredMessagePart.Body.Data == "" || requiredMessagePart.Body.Data == null)
                                requiredMessagePart = latestMessageDetails.Payload.Parts.FirstOrDefault(x => x.MimeType == "text/html");
                        }
                    }
                    if (requiredMessagePart != null)
                    {
                        byte[] data = Convert.FromBase64String(requiredMessagePart.Body.Data.Replace('-', '+').Replace('_', '/').Replace(" ", "+"));
                        requiredMessage = Encoding.UTF8.GetString(data);
                        if (markRead)
                        {
                            var labelToRemove = new List<string> { _labelUnread };
                            service.RemoveLabels(latestMessage.Id, labelToRemove, userId: userId);
                        }
                    }
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return requiredMessage;
            }
            else
            {
                if (disposeGmailService)
                    service.DisposeGmailService();
                return null;
            }
        }

        /// <summary>
        /// Gets Gmail latest message attachments for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="directoryPath">Directory path to download files into. Throws 'DirectoryNotFoundException' if path not found. Similar downloaded files in same path are overwritten.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Count of email message attachments downloaded.</returns>
        /// <exception cref="DirectoryNotFoundException">Throws 'DirectoryNotFoundException' if directory path not found.</exception>
        public static int GetMessageAttachments(this GmailService gmailService, string query, string directoryPath, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            if (!Directory.Exists(directoryPath))
                throw new DirectoryNotFoundException(string.Format("Path - '{0}' Not Found.", directoryPath));
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            int count = 0;
            if (messages.Count > 0)
            {
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var messageRequest = service.Users.Messages.Get(userId, latestMessage.Id);
                    messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Full;
                    var latestMessageDetails = messageRequest.Execute();
                    if (latestMessageDetails.Payload != null && latestMessageDetails.Payload.Parts.Count > 0)
                    {
                        foreach (var part in latestMessageDetails.Payload.Parts)
                        {
                            if (part.Filename != "")
                            {
                                var messageAttachmentRequest = service.Users.Messages.Attachments.Get(userId, latestMessageDetails.Id, part.Body.AttachmentId);
                                var messageAttachmentResponse = messageAttachmentRequest.Execute();
                                var messageAttachmentData = Convert.FromBase64String(messageAttachmentResponse.Data.Replace('-', '+').Replace('_', '/').Replace(" ", "+"));
                                File.WriteAllBytes(Path.Combine(directoryPath, part.Filename), messageAttachmentData);
                                count++;
                            }
                        }
                    }
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return count;
            }
            else
            {
                if (disposeGmailService)
                    service.DisposeGmailService();
                return count;
            }
        }

        /// <summary>
        /// Gets Gmail messages attachments for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="directoryPath">Directory path to download files into. Throws 'DirectoryNotFoundException' if path not found. Similar downloaded files in same path are overwritten.
        /// Directories are created inside this path using message id to download all attachments linked to a particular message if present.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Dictionary with message id and count of attachments downloaded for a particular message id.</returns>
        /// <exception cref="DirectoryNotFoundException">Throws 'DirectoryNotFoundException' if directory path not found.</exception>
        public static Dictionary<string, int> GetMessagesAttachments(this GmailService gmailService, string query, string directoryPath, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            var originalDirectoryPath = directoryPath;
            var attachmentInfo = new Dictionary<string, int>();
            if (!Directory.Exists(directoryPath))
                throw new DirectoryNotFoundException(string.Format("Path - '{0}' Not Found.", directoryPath));
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Full;
                var latestMessageDetails = messageRequest.Execute();
                if (latestMessageDetails.Payload != null && latestMessageDetails.Payload.Parts.Count > 0)
                {
                    int count = 0;
                    foreach (var part in latestMessageDetails.Payload.Parts)
                    {
                        if (part.Filename != "")
                        {
                            directoryPath = Path.Combine(originalDirectoryPath, message.Id);
                            if (!Directory.Exists(directoryPath))
                                Directory.CreateDirectory(directoryPath);
                            var messageAttachmentRequest = service.Users.Messages.Attachments.Get(userId, latestMessageDetails.Id, part.Body.AttachmentId);
                            var messageAttachmentResponse = messageAttachmentRequest.Execute();
                            var messageAttachmentData = Convert.FromBase64String(messageAttachmentResponse.Data.Replace('-', '+').Replace('_', '/').Replace(" ", "+"));
                            File.WriteAllBytes(Path.Combine(directoryPath, part.Filename), messageAttachmentData);
                            count++;
                        }
                    }
                    if (count > 0)
                        attachmentInfo.Add(message.Id, count);
                }
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return attachmentInfo;
        }

        /// <summary>
        /// Sends Gmail message.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="emailContentType">'EmailContentType' enum value. Email body 'PLAIN' for 'text/plain' format', 'HTML' for 'text/html' format'.</param>
        /// <param name="to">'To' email id value. Comma separated value for multiple 'to' email ids. Throws 'FormatException' for invalid email id value.</param>
        /// <param name="cc">'Cc' email id value. Comma separated value for multiple 'cc' email ids. Throws 'FormatException' for invalid email id value.</param>
        /// <param name="bcc">'Bcc' email id value. Comma separated value for multiple 'bcc' email ids. Throws 'FormatException' for invalid email id value.</param>
        /// <param name="subject">'Subject' for email value.</param>
        /// <param name="body">'Body' for email 'text/plain' or 'text/html' value.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <exception cref="FormatException">Throws 'FormatException' for invalid email id value in 'to', 'cc' & 'bcc'.</exception>
        public static void SendMessage(this GmailService gmailService, EmailContentType emailContentType, string to, string cc = "", string bcc = "", string subject = "", string body = "", string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            string payload = "";
            var toList = to.Split(',');
            foreach (var email in toList)
                if (!email.IsValidEmail())
                    throw new FormatException(string.Format("Not a valid 'To' email address. Email: '{0}'", email));
            if (cc != "")
            {
                var ccList = cc.Split(',');
                foreach (var email in ccList)
                    if (!email.IsValidEmail())
                        throw new FormatException(string.Format("Not a valid 'Cc' email address. Email: '{0}'", email));
            }
            if (bcc != "")
            {
                var bccList = bcc.Split(',');
                foreach (var email in bccList)
                    if (!email.IsValidEmail())
                        throw new FormatException(string.Format("Not a valid 'Bcc' email address. Email: '{0}'", email));
            }
            if (emailContentType.Equals(EmailContentType.PLAIN))
            {
                payload = $"To: {to}\r\n" +
                               $"Cc: {cc}\r\n" +
                               $"Bcc: {bcc}\r\n" +
                               $"Subject: {subject}\r\n" +
                               "Content-Type: text/plain; charset=utf-8\r\n\r\n" +
                               $"{body}";
            }
            else if (emailContentType.Equals(EmailContentType.HTML))
            {
                payload = $"To: {to}\r\n" +
                               $"Cc: {cc}\r\n" +
                               $"Bcc: {bcc}\r\n" +
                               $"Subject: {subject}\r\n" +
                               "Content-Type: text/html; charset=utf-8\r\n\r\n" +
                               $"{body}";
            }
            byte[] data = Encoding.UTF8.GetBytes(payload);
            var rawMessage = Convert.ToBase64String(data).Replace("+", "-").Replace("/", "_").Replace("=", "");
            var message = new Message
            {
                Raw = rawMessage
            };
            var sendRequest = service.Users.Messages.Send(message, userId);
            sendRequest.Execute();
            if (disposeGmailService)
                service.DisposeGmailService();
        }

        /// <summary>
        /// Sends Gmail message with attachments.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="emailContentType">'EmailContentType' enum value. Email body 'PLAIN' for 'text/plain' format', 'HTML' for 'text/html' format'.</param>
        /// <param name="to">'To' email id value. Comma separated value for multiple 'to' email ids. Throws 'FormatException' for invalid email id value.</param>
        /// <param name="attachments">List of attachment file paths. Throws 'FileNotFoundException' if file path not found.</param>
        /// Gmail standard attachment size limits and file prohibition rules apply.
        /// <param name="cc">'Cc' email id value. Comma separated value for multiple 'cc' email ids. Throws 'FormatException' for invalid email id value.</param>
        /// <param name="bcc">'Bcc' email id value. Comma separated value for multiple 'bcc' email ids. Throws 'FormatException' for invalid email id value.</param>
        /// <param name="subject">'Subject' for email value.</param>
        /// <param name="body">'Body' for email 'text/plain' or 'text/html' value.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <exception cref="FormatException">Throws 'FormatException' for invalid email id value in 'to', 'cc' & 'bcc'.</exception>
        /// <exception cref="FileNotFoundException">Throws 'FileNotFoundException' if file path not found.</exception>
        public static void SendMessage(this GmailService gmailService, EmailContentType emailContentType, string to, List<string> attachments, string cc = "", string bcc = "", string subject = "", string body = "", string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            var mimeMessage = new MimeMessage();
            var toList = to.Split(',');
            foreach (var email in toList)
            {
                if (!email.IsValidEmail())
                    throw new FormatException(string.Format("Not a valid 'To' email address. Email: '{0}'", email));
                else
                    mimeMessage.To.Add(new MailboxAddress(null, email));
            }
            if (cc != "")
            {
                var ccList = cc.Split(',');
                foreach (var email in ccList)
                {
                    if (!email.IsValidEmail())
                        throw new FormatException(string.Format("Not a valid 'Cc' email address. Email: '{0}'", email));
                    else
                        mimeMessage.Cc.Add(new MailboxAddress(null, email));
                }
            }
            if (bcc != "")
            {
                var bccList = bcc.Split(',');
                foreach (var email in bccList)
                {
                    if (!email.IsValidEmail())
                        throw new FormatException(string.Format("Not a valid 'Bcc' email address. Email: '{0}'", email));
                    else
                        mimeMessage.Bcc.Add(new MailboxAddress(null, email));
                }
            }
            mimeMessage.Subject = subject;
            var builder = new BodyBuilder();
            if (emailContentType.Equals(EmailContentType.PLAIN))
            {
                builder.TextBody = body;
            }
            else if (emailContentType.Equals(EmailContentType.HTML))
            {
                builder.HtmlBody = body;
            }
            foreach (var attachment in attachments)
            {
                if (File.Exists(attachment))
                    builder.Attachments.Add(attachment);
                else
                    throw new FileNotFoundException(string.Format("Attachment file '{0}' not found.", attachment));
            }
            mimeMessage.Body = builder.ToMessageBody();
            byte[] data;
            using (MemoryStream stream = new MemoryStream())
            {
                mimeMessage.WriteTo(stream);
                data = stream.ToArray();
            }
            var rawMessage = Convert.ToBase64String(data).Replace("+", "-").Replace("/", "_").Replace("=", "");
            var message = new Message
            {
                Raw = rawMessage
            };
            var sendRequest = service.Users.Messages.Send(message, userId);
            sendRequest.Execute();
            if (disposeGmailService)
                service.DisposeGmailService();
        }

        /// <summary>
        /// Moves Gmail latest message for a specified query criteria to trash.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was moved to trash or not.</returns>
        public static bool MoveMessageToTrash(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var isMoved = false;
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var moveToTrashRequest = service.Users.Messages.Trash(userId, latestMessage.Id);
                    moveToTrashRequest.Execute();
                    isMoved = true;
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return isMoved;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Moves Gmail messages for a specified query criteria to trash.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Count of email messages moved to trash.</returns>
        public static int MoveMessagesToTrash(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            int counter = 0;
            var service = gmailService;
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var moveToTrashRequest = service.Users.Messages.Trash(userId, message.Id);
                moveToTrashRequest.Execute();
                counter++;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Moves Gmail latest message for a specified query criteria from trash to inbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was untrashed and moved to inbox or not.</returns>
        public static bool UntrashMessage(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var isMoved = false;
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var untrashMessageRequest = service.Users.Messages.Untrash(userId, latestMessage.Id);
                    untrashMessageRequest.Execute();
                    var labelToAdd = new List<string> { _labelInbox };
                    service.AddLabels(latestMessage.Id, labelToAdd, userId: userId);
                    isMoved = true;
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return isMoved;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Moves Gmail messages for a specified query criteria from trash to inbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Count of email messages untrashed and moved to inbox.</returns>
        public static int UntrashMessages(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            int counter = 0;
            var service = gmailService;
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var untrashMessageRequest = service.Users.Messages.Untrash(userId, message.Id);
                untrashMessageRequest.Execute();
                var labelToAdd = new List<string> { _labelInbox };
                service.AddLabels(message.Id, labelToAdd, userId: userId);
                counter++;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Marks Gmail latest message for a specified query criteria as spam.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was marked as spam or not.</returns>
        public static bool ReportSpamMessage(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { _labelSpam },
                RemoveLabelIds = new List<string> { _labelInbox }
            };
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var isModified = false;
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                    modifyMessageRequest.Execute();
                    isModified = true;
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return isModified;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Marks Gmail messages for a specified query criteria as spam.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Count of email messages marked as spam.</returns>
        public static int ReportSpamMessages(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { _labelSpam },
                RemoveLabelIds = new List<string> { _labelInbox }
            };
            int counter = 0;
            var service = gmailService;
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, message.Id);
                modifyMessageRequest.Execute();
                counter++;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Marks Gmail latest message for a specified query criteria as not spam and moves to inbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was marked as not spam or not.</returns>
        public static bool UnspamMessage(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { _labelInbox },
                RemoveLabelIds = new List<string> { _labelSpam }
            };
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var isModified = false;
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                    modifyMessageRequest.Execute();
                    isModified = true;
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return isModified;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Marks Gmail messages for a specified query criteria as not spam and moves them to inbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Count of email messages marked as not spam.</returns>
        public static int UnspamMessages(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { _labelInbox },
                RemoveLabelIds = new List<string> { _labelSpam }
            };
            int counter = 0;
            var service = gmailService;
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, message.Id);
                modifyMessageRequest.Execute();
                counter++;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Marks Gmail latest message for a specified query criteria as read.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was marked as read or not.</returns>
        public static bool MarkMessageAsRead(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var mods = new ModifyMessageRequest
            {
                RemoveLabelIds = new List<string> { _labelUnread }
            };
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var isModified = false;
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                    modifyMessageRequest.Execute();
                    isModified = true;
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return isModified;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Marks Gmail messages for a specified query criteria as read.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Count of email messages marked as read.</returns>
        public static int MarkMessagesAsRead(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var mods = new ModifyMessageRequest
            {
                RemoveLabelIds = new List<string> { _labelUnread }
            };
            int counter = 0;
            var service = gmailService;
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, message.Id);
                modifyMessageRequest.Execute();
                counter++;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Marks Gmail latest message for a specified query criteria as unread.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was marked as unread or not.</returns>
        public static bool MarkMessageAsUnread(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { _labelUnread }
            };
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var isModified = false;
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                    modifyMessageRequest.Execute();
                    isModified = true;
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return isModified;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Marks Gmail messages for a specified query criteria as unread.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Count of email messages marked as unread.</returns>
        public static int MarkMessagesAsUnread(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { _labelUnread }
            };
            int counter = 0;
            var service = gmailService;
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, message.Id);
                modifyMessageRequest.Execute();
                counter++;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Modifies the labels on the latest message for a specified query criteria.
        /// Requires - 'labelsToAdd' And/Or 'labelsToRemove' param value. Throws 'ArgumentException' if none supplied.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="labelsToAdd">Label values to add. Default - 'null'.</param>
        /// <param name="labelsToRemove">Label values to remove. Default - 'null'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Boolean value to confirm if the email message labels for the criteria were modified or not.</returns>
        /// <exception cref="ArgumentException">Throws 'ArgumentException' if none of 'labelsToAdd' and 'labelsToRemove' value is supplied</exception>
        public static bool ModifyMessage(this GmailService gmailService, string query, List<string> labelsToAdd = null, List<string> labelsToRemove = null, string userId = "me", bool disposeGmailService = true)
        {
            if (labelsToAdd == null && labelsToRemove == null)
                throw new ArgumentException("Parameters 'labelsToAdd' / 'labelsToRemove' required.");
            var mods = new ModifyMessageRequest();
            if (labelsToAdd != null)
                mods.AddLabelIds = labelsToAdd;
            if (labelsToRemove != null)
                mods.RemoveLabelIds = labelsToRemove;
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var isModified = false;
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null)
                {
                    var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                    modifyMessageRequest.Execute();
                    isModified = true;
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return isModified;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Modifies the labels on the messages for a specified query criteria.
        /// Requires - 'labelsToAdd' And/Or 'labelsToRemove' param value. Throws 'ArgumentException' if none supplied.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="labelsToAdd">Label values to add. Default - 'null'.</param>
        /// <param name="labelsToRemove">Label values to remove. Default - 'null'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Count of email messages with labels modified.</returns>
        /// <exception cref="ArgumentException">Throws 'ArgumentException' if none of 'labelsToAdd' and 'labelsToRemove' value is supplied</exception>
        public static int ModifyMessages(this GmailService gmailService, string query, List<string> labelsToAdd = null, List<string> labelsToRemove = null, string userId = "me", bool disposeGmailService = true)
        {
            if (labelsToAdd == null && labelsToRemove == null)
                throw new ArgumentException("Parameters 'labelsToAdd' / 'labelsToRemove' required.");
            var mods = new ModifyMessageRequest();
            if (labelsToAdd != null)
                mods.AddLabelIds = labelsToAdd;
            if (labelsToRemove != null)
                mods.RemoveLabelIds = labelsToRemove;
            int counter = 0;
            var service = gmailService;
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, message.Id);
                modifyMessageRequest.Execute();
                counter++;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Gets the labels on the latest message for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>List of email message labels.</returns>
        public static List<Label> GetMessageLabels(this GmailService gmailService, string query, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            List<Message> result = new List<Message>();
            List<Message> messages = new List<Message>();
            List<Label> labels = new List<Label>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                ListMessagesResponse response = request.Execute();
                if (response.Messages != null)
                    result.AddRange(response.Messages);
                request.PageToken = response.NextPageToken;
            } while (!string.IsNullOrEmpty(request.PageToken));
            foreach (var message in result)
            {
                var messageRequest = service.Users.Messages.Get(userId, message.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Minimal;
                var currentMessage = messageRequest.Execute();
                messages.Add(currentMessage);
            }
            if (messages.Count > 0)
            {
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                if (latestMessage != null && latestMessage.LabelIds.Count > 0)
                {
                    foreach (var labelId in latestMessage.LabelIds)
                    {
                        var getLabelsRequest = service.Users.Labels.Get(userId, labelId);
                        var getLabelsResponse = getLabelsRequest.Execute();
                        labels.Add(getLabelsResponse);
                    }
                }
                if (disposeGmailService)
                    service.DisposeGmailService();
                return labels;
            }
            else
            {
                if (disposeGmailService)
                    service.DisposeGmailService();
                return labels;
            }
        }

        /// <summary>
        /// Creates a new Gmail label of type - 'user'.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="labelName">Label name value, should be unique.</param>
        /// <param name="labelBackgroundColor">Label background hex color value. Default - '#666666'.</param>
        /// <param name="labelTextColor">Label text hex color value. Default - '#ffffff'.</param>
        /// <param name="labelListVisibility">'LabelListVisibility' enum value. Default - 'LABEL_SHOW'.</param>
        /// 'LABEL_SHOW' - 'Show the label in the label list'.
        /// 'LABEL_SHOW_IF_UNREAD' - 'Show the label if there are any unread messages with that label'.
        /// 'LABEL_HIDE' - 'Do not show the label in the label list'.
        /// <param name="messageListVisibility">'MessageListVisibility' enum value. Default - 'SHOW'.</param>
        /// 'SHOW' - 'Show the label in the message list'.
        /// 'HIDE' - 'Do not show the label in the message list'.
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Created user label.</returns>
        public static Label CreateUserLabel(this GmailService gmailService, string labelName, string labelBackgroundColor = "#666666", string labelTextColor = "#ffffff", LabelListVisibility labelListVisibility = LabelListVisibility.LABEL_SHOW, MessageListVisibility messageListVisibility = MessageListVisibility.SHOW, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            var requiredLabelListVisibility = "";
            var requiredMessageListVisibility = "";
            if (labelListVisibility.Equals(LabelListVisibility.LABEL_SHOW))
                requiredLabelListVisibility = "labelShow";
            else if (labelListVisibility.Equals(LabelListVisibility.LABEL_SHOW_IF_UNREAD))
                requiredLabelListVisibility = "labelShowIfUnread";
            else if (labelListVisibility.Equals(LabelListVisibility.LABEL_HIDE))
                requiredLabelListVisibility = "labelHide";
            if (messageListVisibility.Equals(MessageListVisibility.SHOW))
                requiredMessageListVisibility = "show";
            else if (messageListVisibility.Equals(MessageListVisibility.HIDE))
                requiredMessageListVisibility = "hide";
            var labelColor = new LabelColor()
            {
                BackgroundColor = labelBackgroundColor,
                TextColor = labelTextColor
            };
            var labelBody = new Label()
            {
                Name = labelName,
                Color = labelColor,
                LabelListVisibility = requiredLabelListVisibility,
                MessageListVisibility = requiredMessageListVisibility,
                Type = "user"
            };
            var createUserLabelRequest = service.Users.Labels.Create(labelBody, userId);
            var label = createUserLabelRequest.Execute();
            if (disposeGmailService)
                service.DisposeGmailService();
            return label;
        }

        /// <summary>
        /// Updates the specified Gmail user label.
        /// Updates label type of - 'user/not defined' with label type - 'user', ignores label updates of type - 'system'.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="oldLabelName">Old existing label name (case sensitive) value, also used to search existing label to be modified.</param>
        /// Use 'ListUserLabels()' to get the correct information for existing labels.
        /// <param name="newLabelName">New label name value, should be unique.</param>
        /// <param name="labelBackgroundColor">Label background hex color value. Default - '#666666'.</param>
        /// <param name="labelTextColor">Label text hex color value. Default - '#ffffff'.</param>
        /// <param name="labelListVisibility">'LabelListVisibility' enum value. Default - 'LABEL_SHOW'.</param>
        /// 'LABEL_SHOW' - 'Show the label in the label list'.
        /// 'LABEL_SHOW_IF_UNREAD' - 'Show the label if there are any unread messages with that label'.
        /// 'LABEL_HIDE' - 'Do not show the label in the label list'.
        /// <param name="messageListVisibility">'MessageListVisibility' enum value. Default - 'SHOW'.</param>
        /// 'SHOW' - 'Show the label in the message list'.
        /// 'HIDE' - 'Do not show the label in the message list'.
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Updated user label.</returns>
        public static Label UpdateUserLabel(this GmailService gmailService, string oldLabelName, string newLabelName, string labelBackgroundColor = "#666666", string labelTextColor = "#ffffff", LabelListVisibility labelListVisibility = LabelListVisibility.LABEL_SHOW, MessageListVisibility messageListVisibility = MessageListVisibility.SHOW, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            var requiredLabelListVisibility = "";
            var requiredMessageListVisibility = "";
            if (labelListVisibility.Equals(LabelListVisibility.LABEL_SHOW))
                requiredLabelListVisibility = "labelShow";
            else if (labelListVisibility.Equals(LabelListVisibility.LABEL_SHOW_IF_UNREAD))
                requiredLabelListVisibility = "labelShowIfUnread";
            else if (labelListVisibility.Equals(LabelListVisibility.LABEL_HIDE))
                requiredLabelListVisibility = "labelHide";
            if (messageListVisibility.Equals(MessageListVisibility.SHOW))
                requiredMessageListVisibility = "show";
            else if (messageListVisibility.Equals(MessageListVisibility.HIDE))
                requiredMessageListVisibility = "hide";
            var labelColor = new LabelColor()
            {
                BackgroundColor = labelBackgroundColor,
                TextColor = labelTextColor
            };
            var labelBody = new Label()
            {
                Name = newLabelName,
                Color = labelColor,
                LabelListVisibility = requiredLabelListVisibility,
                MessageListVisibility = requiredMessageListVisibility,
                Type = "user"
            };
            var listLabelRequest = service.Users.Labels.List(userId);
            var listLabelResponse = listLabelRequest.Execute();
            var label = listLabelResponse.Labels.FirstOrDefault(x => x.Name.Equals(oldLabelName) && !x.Type.Equals("system"));
            if (label != null)
            {
                var updateUserLabelRequest = service.Users.Labels.Update(labelBody, userId, label.Id);
                var updatedLabel = updateUserLabelRequest.Execute();
                if (disposeGmailService)
                    service.DisposeGmailService();
                return updatedLabel;
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return label;
        }

        /// <summary>
        /// Immediately and permanently deletes the specified label and removes it from any messages and threads that it is applied to.
        /// In-built 'system' type labels cannot be deleted e.g INBOX, DRAFTS, SENT, SPAM etc.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="labelName">Label name value.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Boolean value to confirm if the label was deleted or not.</returns>
        public static bool DeleteLabel(this GmailService gmailService, string labelName, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            var listLabelRequest = service.Users.Labels.List(userId);
            var listLabelResponse = listLabelRequest.Execute();
            var label = listLabelResponse.Labels.FirstOrDefault(x => x.Name.Equals(labelName));
            if (label != null)
            {
                var deleteLabelRequest = service.Users.Labels.Delete(userId, label.Id);
                deleteLabelRequest.Execute();
                if (disposeGmailService)
                    service.DisposeGmailService();
                return true;
            }
            else
            {
                if (disposeGmailService)
                    service.DisposeGmailService();
                return false;
            }
        }

        /// <summary>
        /// Lists all labels in the Gmail user's mailbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="disposeGmailService">Boolean value to choose whether to dispose Gmail service instance used or not. Default - 'true'.</param>
        /// <returns>Lists of Gmail user's mailbox labels.</returns>
        public static List<Label> ListUserLabels(this GmailService gmailService, string userId = "me", bool disposeGmailService = true)
        {
            var service = gmailService;
            List<Label> labels = new List<Label>();
            var listLabelsRequest = service.Users.Labels.List(userId);
            var listLabelsResponse = listLabelsRequest.Execute();
            foreach (var label in listLabelsResponse.Labels)
            {
                labels.Add(label);
            }
            if (disposeGmailService)
                service.DisposeGmailService();
            return labels;
        }

        /// <summary>
        /// Checks email format.
        /// </summary>
        /// <param name="email">Email to validate.</param>
        /// <returns>Boolean value to confirm if email Id format is valid or not.</returns>
        private static bool IsValidEmail(this string email)
        {
            string pattern = @"^[^0-9](?!\.)(""([^""\r\\]|\\[""\r\\])*""|"
                            + @"([-a-z0-9!#$%&'*+/=?^_`{|}~]|(?<!\.)\.)*)(?<!\.)"
                            + @"@[a-z0-9][\w\.-]*[a-z0-9]\.[a-z][a-z\.]*[a-z]$";
            var regex = new Regex(pattern, RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(100));
            return regex.IsMatch(email);
        }

        /// <summary>
        /// Modifies Gmail message for labels to be removed.
        /// </summary>
        /// <param name="service">Gmail Service initializer value.</param>
        /// <param name="messageId">'Message Id' to modify.</param>
        /// <param name="labelsToRemove">'Labels' to remove.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        private static void RemoveLabels(this GmailService service, string messageId, List<string> labelsToRemove, string userId = "me")
        {
            ModifyMessageRequest mods = new ModifyMessageRequest()
            {
                RemoveLabelIds = labelsToRemove
            };
            service.Users.Messages.Modify(mods, userId, messageId).Execute();
        }

        /// <summary>
        /// Modifies Gmail message for labels to be added.
        /// </summary>
        /// <param name="service">Gmail Service initializer value.</param>
        /// <param name="messageId">'Message Id' to modify.</param>
        /// <param name="labelsToAdd">'Labels' to add.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        private static void AddLabels(this GmailService service, string messageId, List<string> labelsToAdd, string userId = "me")
        {
            ModifyMessageRequest mods = new ModifyMessageRequest()
            {
                AddLabelIds = labelsToAdd
            };
            service.Users.Messages.Modify(mods, userId, messageId).Execute();
        }
    }
}