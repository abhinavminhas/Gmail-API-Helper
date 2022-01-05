using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace GmailAPIHelper
{
    /// <summary>
    /// Gmail Helper.
    /// </summary>
    public static class GmailHelper
    {
        private static List<string> _scopes;
        private static string _applicationName;

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
        /// Sets the credentials path to be used.
        /// </summary>
        /// <param name="tokenPathType">'TokenPathType' enum value. 'HOME' for users home directory, 'WORKING_DIRECTORY' for working directory, 'CUSTOM' for any other custom path to be used.</param>
        /// <param name="tokenPath">Token path value in case of 'TokenPathType - CUSTOM' value.</param>
        /// <returns>Credentials file path.</returns>
        private static string SetCredentialPath(TokenPathType tokenPathType, string tokenPath = "")
        {
            string credPath = "";
            if (tokenPathType == TokenPathType.HOME)
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    credPath = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%") + "\\" + "token.json";
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                    credPath = Environment.GetEnvironmentVariable("HOME") + "/" + "token.json";
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                    credPath = Environment.GetEnvironmentVariable("HOME") + "/" + "token.json";
                else
                    throw new Exception("OS Platform: Not 'Windows/Linux/OSX' Platform.");
            }
            else if (tokenPathType == TokenPathType.WORKING_DIRECTORY)
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    credPath = Environment.CurrentDirectory + "\\" + "token.json";
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                    credPath = Environment.CurrentDirectory + "/" + "token.json";
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                    credPath = Environment.CurrentDirectory + "/" + "token.json";
                else
                    throw new Exception("OS Platform: Not 'Windows/Linux/OSX' Platform.");
            }
            else if (tokenPathType == TokenPathType.CUSTOM)
                credPath = tokenPath;
            return credPath;
        }

        /// <summary>
        /// Connects to the 'Gmail Service'.
        /// </summary>
        /// <param name="applicationName">'Application Name' value created in Gmail API Console.</param>
        /// <param name="tokenPathType">'TokenPathType' enum value. 'HOME' for users home directory, 'WORKING_DIRECTORY' for working directory, 'CUSTOM' for any other custom path to be used.
        /// Default value - 'WORKING_DIRECTORY'.</param>
        /// <param name="tokenPath">'token.json' path to save generated token from gmail authentication/authorization. 
        /// Always asks in case of change in gmail authentication or valid token file missing in the given path. Default path is blank, required for 'TokenPathType - CUSTOM'.</param>
        /// <returns>Gmail Service.</returns>
        public static GmailService GetGmailService(string applicationName, TokenPathType tokenPathType = TokenPathType.WORKING_DIRECTORY, string tokenPath = "")
        {
            _scopes = new List<string>();
            _applicationName = applicationName;
            _scopes.Add(GmailService.Scope.GmailMetadata);
            _scopes.Add(GmailService.Scope.GmailReadonly);
            _scopes.Add(GmailService.Scope.GmailModify);
            _scopes.Add(GmailService.Scope.GmailLabels);
            _scopes.Add(GmailService.Scope.GmailSend);
            UserCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                //The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
                string credPath = SetCredentialPath(tokenPathType, tokenPath);
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    _scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }
            //Create Gmail API service.
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = _applicationName
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
        /// Returns Gmail latest complete message with metadata for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="markRead">Boolean value to mark retrieved latest message as read. Default - 'false'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Email message matching the search criteria.</returns>
        public static Message GetMessage(this GmailService gmailService, string query, bool markRead = false, string userId = "me")
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
                var messageRequest = service.Users.Messages.Get(userId, requiredLatestMessage.Id);
                messageRequest.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Full;
                requiredLatestMessage = messageRequest.Execute();
                if (markRead)
                {
                    var labelToRemove = new List<string> { "UNREAD" };
                    service.RemoveLabels(requiredLatestMessage.Id, labelToRemove, userId: userId);
                }
                service.DisposeGmailService();
                return requiredLatestMessage;
            }
            else
            {
                service.DisposeGmailService();
                return null;
            }
        }

        /// <summary>
        /// Returns Gmail messages with metadata for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="markRead">Boolean value to mark retrieved messages as read. Default - 'false'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>List of email messages matching the search criteria.</returns>
        public static List<Message> GetMessages(this GmailService gmailService, string query, bool markRead = false, string userId = "me")
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
                    var labelToRemove = new List<string> { "UNREAD" };
                    service.RemoveLabels(message.Id, labelToRemove, userId: userId);
                }
            }
            service.DisposeGmailService();
            return messages;
        }

        /// <summary>
        /// Returns Gmail latest message body text for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="markRead">Boolean value to mark retrieved latest email as read. Default - 'false'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Email message body in 'text/plain' format.</returns>
        public static string GetLatestMessage(this GmailService gmailService, string query, bool markRead = false, string userId = "me")
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
                        if (requiredMessagePart.Body.Data == "" || requiredMessagePart.Body.Data == null)
                            requiredMessagePart = latestMessageDetails.Payload.Parts.FirstOrDefault(x => x.MimeType == "text/html");
                    }
                }
                if (requiredMessagePart != null)
                {
                    byte[] data = Convert.FromBase64String(requiredMessagePart.Body.Data.Replace('-', '+').Replace('_', '/').Replace(" ", "+"));
                    requiredMessage = Encoding.UTF8.GetString(data);
                    if (markRead)
                    {
                        var labelToRemove = new List<string> { "UNREAD" };
                        service.RemoveLabels(latestMessage.Id, labelToRemove, userId: userId);
                    }
                }
                else
                    requiredMessagePart = null;
                service.DisposeGmailService();
                return requiredMessage;
            }
            else
            {
                service.DisposeGmailService();
                return null;
            }
        }

        /// <summary>
        /// Sends Gmail message.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="emailContentType">'EmailContentType' enum value. 'PLAIN' for 'text/plain' format', 'HTML' for 'text/html' format'.</param>
        /// <param name="to">'To' email id value. Comma separated value for multiple 'to' email ids.</param>
        /// <param name="cc">'Cc' email id value. Comma separated value for multiple 'cc' email ids.</param>
        /// <param name="bcc">'Bcc' email id value. Comma separated value for multiple 'bcc' email ids.</param>
        /// <param name="subject">'Subject' for email value.</param>
        /// <param name="body">'Body' for email 'text/plain' or 'text/html' value.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        public static void SendMessage(this GmailService gmailService, EmailContentType emailContentType, string to, string cc = "", string bcc = "", string subject = "", string body = "", string userId = "me")
        {
            var service = gmailService;
            string payload = "";
            var toList = to.Split(',');
            foreach (var email in toList)
                if (!email.IsValidEmail())
                    throw new Exception(string.Format("Not a valid 'To' email address. Email: '{0}'", email));
            if (cc != "")
            {
                var ccList = cc.Split(',');
                foreach (var email in ccList)
                    if (!email.IsValidEmail())
                        throw new Exception(string.Format("Not a valid 'Cc' email address. Email: '{0}'", email));
            }
            if (bcc != "")
            {
                var bccList = bcc.Split(',');
                foreach (var email in bccList)
                    if (!email.IsValidEmail())
                        throw new Exception(string.Format("Not a valid 'Bcc' email address. Email: '{0}'", email));
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
            service.DisposeGmailService();
        }

        /// <summary>
        /// Moves Gmail latest message for a specified query criteria to trash.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was moved to trash or not.</returns>
        public static bool MoveMessageToTrash(this GmailService gmailService, string query, string userId = "me")
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
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                var moveToTrashRequest = service.Users.Messages.Trash(userId, latestMessage.Id);
                moveToTrashRequest.Execute();
                service.DisposeGmailService();
                return true;
            }
            service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Moves Gmail messages for a specified query criteria to trash.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Count of email messages moved to trash.</returns>
        public static int MoveMessagesToTrash(this GmailService gmailService, string query, string userId = "me")
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
            service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Moves Gmail latest message for a specified query criteria from trash to inbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was untrashed and moved to inbox or not.</returns>
        public static bool UntrashMessage(this GmailService gmailService, string query, string userId = "me")
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
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                var untrashMessageRequest = service.Users.Messages.Untrash(userId, latestMessage.Id);
                untrashMessageRequest.Execute();
                var labelToAdd = new List<string> { "INBOX" };
                service.AddLabels(latestMessage.Id, labelToAdd, userId: userId);
                service.DisposeGmailService();
                return true;
            }
            service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Moves Gmail messages for a specified query criteria from trash to inbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Count of email messages untrashed and moved to inbox.</returns>
        public static int UntrashMessages(this GmailService gmailService, string query, string userId = "me")
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
                var labelToAdd = new List<string> { "INBOX" };
                service.AddLabels(message.Id, labelToAdd, userId: userId);
                counter++;
            }
            service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Marks Gmail latest message for a specified query criteria as spam.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was marked as spam or not.</returns>
        public static bool ReportSpamMessage(this GmailService gmailService, string query, string userId = "me")
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { "SPAM" },
                RemoveLabelIds = new List<string> { "INBOX" }
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
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                modifyMessageRequest.Execute();
                service.DisposeGmailService();
                return true;
            }
            service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Marks Gmail messages for a specified query criteria as spam.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Count of email messages marked as spam.</returns>
        public static int ReportSpamMessages(this GmailService gmailService, string query, string userId = "me")
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { "SPAM" },
                RemoveLabelIds = new List<string> { "INBOX" }
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
            service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Marks Gmail latest message for a specified query criteria as not spam and moves to inbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was marked as not spam or not.</returns>
        public static bool UnspamMessage(this GmailService gmailService, string query, string userId = "me")
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { "INBOX" },
                RemoveLabelIds = new List<string> { "SPAM" }
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
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                modifyMessageRequest.Execute();
                service.DisposeGmailService();
                return true;
            }
            service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Marks Gmail messages for a specified query criteria as not spam and moves them to inbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Count of email messages marked as not spam.</returns>
        public static int UnspamMessages(this GmailService gmailService, string query, string userId = "me")
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { "INBOX" },
                RemoveLabelIds = new List<string> { "SPAM" }
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
            service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Marks Gmail latest message for a specified query criteria as read.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was marked as read or not.</returns>
        public static bool MarkMessageAsRead(this GmailService gmailService, string query, string userId = "me")
        {
            var mods = new ModifyMessageRequest
            {
                RemoveLabelIds = new List<string> { "UNREAD" }
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
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                modifyMessageRequest.Execute();
                service.DisposeGmailService();
                return true;
            }
            service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Marks Gmail messages for a specified query criteria as read.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Count of email messages marked as read.</returns>
        public static int MarkMessagesAsRead(this GmailService gmailService, string query, string userId = "me")
        {
            var mods = new ModifyMessageRequest
            {
                RemoveLabelIds = new List<string> { "UNREAD" }
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
            service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Marks Gmail latest message for a specified query criteria as unread.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Boolean value to confirm if the email message for the criteria was marked as unread or not.</returns>
        public static bool MarkMessageAsUnread(this GmailService gmailService, string query, string userId = "me")
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { "UNREAD" }
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
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                modifyMessageRequest.Execute();
                service.DisposeGmailService();
                return true;
            }
            service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Marks Gmail messages for a specified query criteria as unread.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Count of email messages marked as unread.</returns>
        public static int MarkMessagesAsUnread(this GmailService gmailService, string query, string userId = "me")
        {
            var mods = new ModifyMessageRequest
            {
                AddLabelIds = new List<string> { "UNREAD" }
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
            service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Modifies the labels on the latest message for a specified query criteria.
        /// Requires - 'labelsToAdd' And/Or 'labelsToRemove' param value.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="labelsToAdd">Label values to add. Default - 'null'.</param>
        /// <param name="labelsToRemove">Label values to remove. Default - 'null'.</param>
        /// <returns>Boolean value to confirm if the email message labels for the criteria were modified or not.</returns>
        public static bool ModifyMessage(this GmailService gmailService, string query, string userId = "me", List<string> labelsToAdd = null, List<string> labelsToRemove = null)
        {
            if (labelsToAdd == null && labelsToRemove == null)
                throw new NullReferenceException("Either 'Labels To Add' or 'Labels to Remove' required.");
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
                var latestMessage = messages.OrderByDescending(item => item.InternalDate).FirstOrDefault();
                var modifyMessageRequest = service.Users.Messages.Modify(mods, userId, latestMessage.Id);
                modifyMessageRequest.Execute();
                service.DisposeGmailService();
                return true;
            }
            service.DisposeGmailService();
            return false;
        }

        /// <summary>
        /// Modifies the labels on the messages for a specified query criteria.
        /// Requires - 'labelsToAdd' And/Or 'labelsToRemove' param value.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <param name="labelsToAdd">Label values to add. Default - 'null'.</param>
        /// <param name="labelsToRemove">Label values to remove. Default - 'null'.</param>
        /// <returns>Count of email messages with labels modified.</returns>
        public static int ModifyMessages(this GmailService gmailService, string query, string userId = "me", List<string> labelsToAdd = null, List<string> labelsToRemove = null)
        {
            if (labelsToAdd == null && labelsToRemove == null)
                throw new NullReferenceException("Either 'Labels To Add' or 'Labels to Remove' required.");
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
            service.DisposeGmailService();
            return counter;
        }

        /// <summary>
        /// Gets the labels on the latest message for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>List of email message labels.</returns>
        public static List<Label> GetMessageLabels(this GmailService gmailService, string query, string userId = "me")
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
                if (latestMessage.LabelIds.Count > 0)
                {
                    foreach (var labelId in latestMessage.LabelIds)
                    {
                        var getLabelsRequest = service.Users.Labels.Get(userId, labelId);
                        var getLabelsResponse = getLabelsRequest.Execute();
                        labels.Add(getLabelsResponse);
                    }
                }
                service.DisposeGmailService();
                return labels;
            }
            else
            {
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
        /// <returns>Created user label.</returns>
        public static Label CreateUserLabel(this GmailService gmailService, string labelName, string labelBackgroundColor = "#666666", string labelTextColor = "#ffffff", LabelListVisibility labelListVisibility = LabelListVisibility.LABEL_SHOW, MessageListVisibility messageListVisibility = MessageListVisibility.SHOW, string userId = "me")
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
        /// <returns>Updated user label.</returns>
        public static Label UpdateUserLabel(this GmailService gmailService, string oldLabelName, string newLabelName, string labelBackgroundColor = "#666666", string labelTextColor = "#ffffff", LabelListVisibility labelListVisibility = LabelListVisibility.LABEL_SHOW, MessageListVisibility messageListVisibility = MessageListVisibility.SHOW, string userId = "me")
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
                service.DisposeGmailService();
                return updatedLabel;
            }
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
        /// <returns>Boolean value to confirm if the label was deleted or not.</returns>
        public static bool DeleteLabel(this GmailService gmailService, string labelName, string userId = "me")
        {
            var service = gmailService;
            var listLabelRequest = service.Users.Labels.List(userId);
            var listLabelResponse = listLabelRequest.Execute();
            var label = listLabelResponse.Labels.FirstOrDefault(x => x.Name.Equals(labelName));
            if (label != null)
            {
                var deleteLabelRequest = service.Users.Labels.Delete(userId, label.Id);
                deleteLabelRequest.Execute();
                service.DisposeGmailService();
                return true;
            }
            else
            {
                service.DisposeGmailService();
                return false;
            }
        }

        /// <summary>
        /// Lists all labels in the Gmail user's mailbox.
        /// </summary>
        /// <param name="gmailService">'Gmail' service initializer value.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)'.</param>
        /// <returns>Lists of Gmail user's mailbox labels.</returns>
        public static List<Label> ListUserLabels(this GmailService gmailService, string userId = "me")
        {
            var service = gmailService;
            List<Label> labels = new List<Label>();
            var listLabelsRequest = service.Users.Labels.List(userId);
            var listLabelsResponse = listLabelsRequest.Execute();
            foreach (var label in listLabelsResponse.Labels)
            {
                labels.Add(label);
            }
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
            var regex = new Regex(pattern, RegexOptions.IgnoreCase);
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