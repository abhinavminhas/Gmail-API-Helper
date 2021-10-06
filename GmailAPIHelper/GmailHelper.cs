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
        /// 'Token Path Types' enum.
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
        /// Sets the credentials path to be used.
        /// </summary>
        /// <param name="tokenPathType">'TokenPathType' enum value. 'HOME' for users home directory, 'WORKING_DIRECTORY' for working directory, 'CUSTOM' for any other custom path to be used.</param>
        /// <param name="tokenPath">Token path value in case of 'TokenPathType - CUSTOM' value.</param>
        /// <returns></returns>
        private static string SetCredentialPath(TokenPathType tokenPathType, string tokenPath = "")
        {
            string credPath = "";
            if (tokenPathType == TokenPathType.HOME)
            {
                var isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
                var isLinux = RuntimeInformation.IsOSPlatform(OSPlatform.Linux);
                var isOSX = RuntimeInformation.IsOSPlatform(OSPlatform.OSX);
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
                var isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
                var isLinux = RuntimeInformation.IsOSPlatform(OSPlatform.Linux);
                var isOSX = RuntimeInformation.IsOSPlatform(OSPlatform.OSX);
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
        /// Always asks in case of change in gmail authentication or valid token file missing in the given path. Default path is users folder.</param>
        /// <returns>Gmail Service.</returns>
        public static GmailService GetGmailService(string applicationName, TokenPathType tokenPathType = TokenPathType.WORKING_DIRECTORY, string tokenPath = "")
        {
            _scopes = new List<string>();
            _applicationName = applicationName;
            _scopes.Add(GmailService.Scope.GmailReadonly);
            _scopes.Add(GmailService.Scope.GmailModify);
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
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = _applicationName
            });
            return service;
        }

        /// <summary>
        /// Returns Gmail latest message body text for a specified query criteria.
        /// </summary>
        /// <param name="gmailService">'Gmail' service value.</param>
        /// <param name="query">'Query' criteria for the email to search.</param>
        /// <param name="markRead">Boolean value to mark retrieved latest email read/unread. Default - 'false'.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)', takes from credentials.json</param>
        /// <returns>Email message body in Text/Plain.</returns>
        public static string GetLatestMessage(this GmailService gmailService, string query, bool markRead = false, string userId = "me")
        {
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
            List<Message> messages = new List<Message>();
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
                var requiredMessagePart = latestMessageDetails.Payload.Parts.FirstOrDefault(x => x.MimeType == "text/plain");
                if (requiredMessagePart != null)
                {
                    byte[] data = Convert.FromBase64String(requiredMessagePart.Body.Data.Replace('-', '+').Replace('_', '/').Replace(" ", "+").Replace("=", "+"));
                    requiredMessage = Encoding.UTF8.GetString(data);
                    if (markRead)
                    {
                        var labelToRemove = new List<string> { "UNREAD" };
                        ModifyMessage(service, latestMessage.Id, labelsToRemove: labelToRemove);
                    }
                }
                else
                    requiredMessagePart = null;
                return requiredMessage;
            }
            else
                return null;
        }

        /// <summary>
        /// Modifies Gmail message for labels criteria supplided.
        /// </summary>
        /// <param name="service">Gmail Service.</param>
        /// <param name="messageId">'Message Id' to modify.</param>
        /// <param name="userId">User's email address. 'User Id' for request to authenticate. Default - 'me (authenticated user)', takes from credentials.json</param>
        /// <param name="labelsToAdd">'Labels' to add</param>
        /// <param name="labelsToRemove">'Labels' to remove.</param>
        internal static void ModifyMessage(GmailService service, string messageId, string userId = "me", List<string> labelsToAdd = null, List<string> labelsToRemove = null)
        {
            if (labelsToAdd == null && labelsToRemove == null)
                throw new Exception("Modify Message: Please provide either labels to modify, 'Labels To Add' or 'Labels to Remove', both cannot be empty.");
            ModifyMessageRequest mods = new ModifyMessageRequest();
            if (labelsToAdd != null)
                mods.AddLabelIds = labelsToAdd;
            if (labelsToRemove != null)
                mods.RemoveLabelIds = labelsToRemove;
            try
            {
                service.Users.Messages.Modify(mods, userId, messageId).Execute();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
        }

    }
}