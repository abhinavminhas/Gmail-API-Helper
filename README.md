# Gmail-API-Helper
*Gmail API helper solution in .NET*. </br></br>
![Gmail-API-Helper (Build)](https://github.com/abhinavminhas/Gmail-API-Helper/actions/workflows/build.yml/badge.svg)
[![codecov](https://codecov.io/gh/abhinavminhas/Gmail-API-Helper/branch/main/graph/badge.svg?token=18ZV2GGET8)](https://codecov.io/gh/abhinavminhas/Gmail-API-Helper)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=abhinavminhas_Gmail-API-Helper&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=abhinavminhas_Gmail-API-Helper)
![maintainer](https://img.shields.io/badge/Creator/Maintainer-abhinavminhas-e65c00)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![NuGet](https://img.shields.io/nuget/v/GmailHelper?color=%23004880&label=Nuget)](https://www.nuget.org/packages/GmailHelper/)  

[![Security Rating](https://sonarcloud.io/api/project_badges/measure?project=abhinavminhas_Gmail-API-Helper&metric=security_rating)](https://sonarcloud.io/summary/new_code?id=abhinavminhas_Gmail-API-Helper)
[![Reliability Rating](https://sonarcloud.io/api/project_badges/measure?project=abhinavminhas_Gmail-API-Helper&metric=reliability_rating)](https://sonarcloud.io/summary/new_code?id=abhinavminhas_Gmail-API-Helper)
[![Maintainability Rating](https://sonarcloud.io/api/project_badges/measure?project=abhinavminhas_Gmail-API-Helper&metric=sqale_rating)](https://sonarcloud.io/summary/new_code?id=abhinavminhas_Gmail-API-Helper)

The [Gmail API](https://developers.google.com/gmail/api) is used to interact with users' Gmail inboxes and settings, supports several popular programming languages. This solution is the implementation of the same providing certain useful extension methods in .NET to retrieve/send/manipulate email messages and manage labels. The solution uses 'Gmail Search Query' to retrieve/manipulate required email messages.

## Download
The package is available and can be downloaded using [nuget.org](https://www.nuget.org/) package manager.  
- Package Name - [GmailHelper](https://www.nuget.org/packages/GmailHelper/).

## Features

1. Retrieve email message/messages based on query search.
2. Retrieve latest email message body based on query search.
3. Download message/messages attachments based on query search.
4. Send email messages (text/plain, text/html).
5. Send email messages with attachments (text/plain, text/html).
6. Trash/Untrash email message/messages based on query search.
7. Spam/Unspam email message/messages based on query search.
8. Mark email message/messages read/unread based on query search.
9. Modify email message/messages labels based on query search.
10. Retrieve latest email message labels based on query search.
11. Create/update/delete/list Gmail user labels.

    **NOTE:** 1. *Gmail query search operators information can be found **[here](https://support.google.com/mail/answer/7190)**. For examples checkout solution tests.*  
    &emsp;&emsp;&emsp; 2. *You can also use Gmail 'Search mail' search criteria option to create search query string.*  
    &emsp;&emsp;&emsp; <img src=https://user-images.githubusercontent.com/17473202/147176323-b4eb4963-1f5d-46e5-9fc1-bef4e8aaa2d2.png />  
    &emsp;&emsp;&emsp; 3. *Delete scope requires fully verified app, use move to trash instead.*  
    &emsp;&emsp;&emsp; 4. *For sending email with attachments, Gmail standard attachment size limits and file prohibition rules apply.*

## .NET Supported Versions

Gmail API solution is built on .NetStandard 2.0  

<img src="https://user-images.githubusercontent.com/17473202/137575806-fdebc1ff-4741-4ada-8974-0459c6e27830.png" />

## Configuration

**Reference:** ***[Authorizing Your App with Gmail](https://developers.google.com/gmail/api/auth/about-auth)***

#### Detailed Steps:

1. Create a Gmail account/use an existing Gmail account and login to **[Gmail API Console](https://console.cloud.google.com/apis/api/gmail)**  

    <img src="https://user-images.githubusercontent.com/17473202/138042516-7b388c1a-977a-4ff1-9d29-82aaeae3d8e7.png" />  
2. Under '**API & Services**' create a new project and provide  
    - Project name
    - Organisation, leave as is if none created.  

    <img src="https://user-images.githubusercontent.com/17473202/138042606-17b8e076-1bb2-4ea5-a722-fdd300204a6a.png" />  
    
3. Select the created project and go to '**Library**' under '**API & Services**'.  

    <img src="https://user-images.githubusercontent.com/17473202/138043014-7f41d875-3fca-4df1-a167-03a976cfb690.png" />  
4. Search '**Gmail API**' and enable the API for the logged in user account.  

    <img src="https://user-images.githubusercontent.com/17473202/138081584-16bf3582-56f2-41c4-b64b-41e6273a6de3.png" />  
    <img src="https://user-images.githubusercontent.com/17473202/138043204-2b876369-35d7-475f-aad8-ac3c2c21fe8c.png" />  
5. In the created project, open '**OAuth consent screen**' under '**API & Services**' and create a new app with **User Type** - '**External**' or none selected.  

    <img src="https://user-images.githubusercontent.com/17473202/138043286-9a670912-4747-4fbb-8988-59bc59305c37.png" />  

    In '**OAuth consent screen**' provide
    - Under '**App Information**'
        - App Name
        - User Support Email
    - Under '**Developer Contact Information**'
        - Email Addresses (can be same as '**User Support Email**' entered above)  
        
    <img src="https://user-images.githubusercontent.com/17473202/138043363-19e4d215-219d-4a96-9629-1f154af4335b.png" />  

    In '**Scopes**' search & select the '**Gmail API**' scope enabled previously.
    - Gmail API <https://mail.google.com/> scope  

    <img src="https://user-images.githubusercontent.com/17473202/138043411-60533281-f2b0-4374-82ab-1ce7ab7e55e1.png" />  

    Can leave the rest as is and save.  

    <img src="https://user-images.githubusercontent.com/17473202/138043492-e249dd83-f405-487e-88f9-d86c6dc0c88b.png" />  

    On created app under '**Publishing Status**' click  
    - PUBLISH APP  
    <img src="https://user-images.githubusercontent.com/17473202/138043559-e5d6512b-75eb-4c62-b999-33b213ef74ee.png" /><br>
    <img src="https://user-images.githubusercontent.com/17473202/138043618-3d23aa98-3869-433e-8831-82ee65823a23.png" />  

    **NOTE:** *Some added restricted/sensitive scopes may require app verification.*

6. Under '**API & Services**' go to '**Credentials**' and create new credentials of type '**OAuth client ID**'.  

    <img src="https://user-images.githubusercontent.com/17473202/138043764-9dcdb1e4-48f4-45c2-856c-470ca89faae6.png" />  
7. Provide below details
   - Application type as 'Desktop app'
   - Name  

    <img src="https://user-images.githubusercontent.com/17473202/138043806-11ca240f-1804-4e1c-aae3-28cd66208b67.png" />  
8. Download the json and save the file as '**credentials.json**'. This needs to be added to the solution to generate OAuth token.  

    <img src="https://user-images.githubusercontent.com/17473202/138043843-8bfbdac9-99c1-45d8-91bf-211c246aacda.png" />  
9. Include the '**credentials.json**' in the solution project and invoke **GmailHelper.GetGmailService()** function below to generate OAuth token.
   ``` csharp
   GmailHelper.GetGmailService(<Created Application Name>, <'token.json' file path>)
   ```
   Additional optional param to set '**credentials.json**' path if not in the solution project
   ``` csharp
   GmailHelper.GetGmailService(<Created Application Name>, <'token.json' file path>, <'credentials.json' file path>)
   ```

10. Sign-In to the login prompt presented with user account used to create the app.  

    <img src="https://user-images.githubusercontent.com/17473202/138043907-0d1f6f12-ba23-4331-9bc0-d97ab257e96d.png" />  
    
    **NOTE:** *'**Google hasn't verified this app**' may be prompted if the app is yet to be verified. Go to '**Advanced**' and continue.*  
11. Select the Gmail API access as added to the app scope '**Read, compose, send and permanently delete all your email from Gmail**' and continue.  

    <img src="https://user-images.githubusercontent.com/17473202/138043948-dd805f6e-faec-4c05-b164-78aabdca55d7.png" />  
12. Once you see '**Received verification code. You may now close this window.**', you can close the browser window and save the contents of '**token.json**' to reuse until credentials are changed or refreshed. The token path parameter in **GmailHelper.GetGmailService()** can be used to save and reuse the generated OAuth token again and again.  
    
    **NOTE:** *Keep contents of '**token.json**' & '**credentials.json**' safe to avoid any misuse. Checkout solution tests for more understanding on usage.*
13. For concurrent usage create a lock object for the used extension methods in a non-static class.  
    **Examples:**
``` csharp
public class Gmail {
  private static readonly object _lock = new object();

  public static Message GetMessage() {
    var service = GmailHelper.GetGmailService(applicationName: "GmailAPIHelper");
    Message message = null;
    lock(_lock) {
      message = GmailHelper.GetMessage(service, query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
    }
    return message;
  }
}
```
``` csharp
public class Gmail {
  private static readonly object _lock = new object();

  public static Message GetMessage() {
    Message message = null;
    lock(_lock) {
      message = GmailHelper.GetGmailService(applicationName: "GmailAPIHelper")
        .GetMessage(query: "[from:test.auto.helper@gmail.com][subject:'READ EMAIL']in:inbox is:read", markRead: true);
    }
    return message;
  }
}
```