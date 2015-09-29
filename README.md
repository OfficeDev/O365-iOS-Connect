#Office 365 Connect app for iOS#

[日本 (日本語)](/loc/README-ja.md) (Japanese)

[![Office 365 Connect sample](/readme-images/O365-iOS-Connect-video_play_icon.png)](https://youtu.be/3v__BnV61Rs "Click to see the sample in action")

Connecting to Office 365 is the first step every iOS app must take to start working with the rich data and services Office 365 offers. This sample shows how to connect to Office 365 and call the sendMail API to send an email from your Office 365 mail account. You can use this sample as a starting point to quickly connect your iOS apps to Office 365. It comes in both Objective-C and Swift versions.

**Table of contents**

* [Set up your environment](#set-up-your-environment)
* [Use CocoaPods to import the O365 iOS SDK](#use-cocoapods-to-import-the-o365-ios-sdk)
* [Register your app with Microsoft Azure](#register-your-app-with-microsoft-azure)
* [Get the Client ID and Redirect Uri into the project](#get-the-client-id-and-redirect-uri-into-the-project)
* [Code of Interest](#code-of-interest)
* [Questions and comments](#questions-and-comments)
* [Troubleshooting](#troubleshooting)
* [Additional resources](#additional-resources)


<a name="set-up-your-environment"></a>
## Set up your environment ##

To run the Office 365 Connect app for iOS, you need the following:


* [Xcode](https://developer.apple.com/) from Apple.
* An Office 365 account. You can get an Office 365 account by signing up for an [Office 365 Developer site](http://msdn.microsoft.com/library/office/fp179924.aspx). This will give you access to the APIs that you can use to create apps that target Office 365 data.
* A Microsoft Azure tenant to register your application. Azure Active Directory provides identity services that applications use for authentication and authorization. To create a trial subscription, and associate it with your O365 account, see [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment) and the section **To create a new Azure subscription and associate it with your Office 365 accounts** for more information.

  **Important**: If you already have an Azure subscription, you'll need to bind that subscription to your Office 365 account. To do this see [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment) and the section **Associate your Office 365 account with Azure AD to create and manage apps** for more information.


* Installation of [CocoaPods](https://cocoapods.org/) as a dependency manager. CocoaPods will allow you to pull the Office 365 and Azure Active Directory Authentication Library (ADAL) dependencies into the project.

Once you have an Office 365 account, an Azure AD account that is bound to your Office 365 Developer site, you'll need to perform the following steps:

1. Register your application with Microsoft Azure, and configure the appropriate Office 365 Exchange Online permissions. We'll show you how to do this later.
2. Install and use CocoaPods to get the Office 365 and ADAL authentication dependencies into your project. We'll show you how to do this later.
3. Enter the Azure app registration specifics (ClientID and RedirectUri) into the Office 365 Connect app.

<a name="use-cocoapods-to-import-the-o365-ios-sdk"></a>
## Use CocoaPods to import the O365 iOS SDK
Note: If you've never used **CocoaPods** before as a dependency manager you'll have to install it prior to getting your Office 365 iOS SDK dependencies into your project. If you already have it installed you may skip this and move on to **Getting the Office 365 SDK for iOS dependencies in your project**.

Enter both these lines of code from the **Terminal** app on your Mac.

    sudo gem install cocoapods
    pod setup

If the install and setup were successful, you should see the message **Setup completed in Terminal**. For more information on CocoaPods, and its usage, see [CocoaPods](https://cocoapods.org/).


**Getting the Office 365 SDK for iOS dependencies in your project.**
The O365 iOS Connect app already contains a podfile that will get the Office 365 and ADAL components (pods) into your project. It's located in either the **objective-c** or **swift** folder based on sample preference ("Podfile"). The example shows the contents of the file.

	target ‘test’ do

    pod 'ADALiOS', '~> 1.2.1'
    pod 'Office365/Outlook', '= 0.9.1'
    pod 'Office365/Discovery', '= 0.9.1'

	end


You'll simply need to navigate to the project directory in the **Terminal** (root of the project folder) and run the following command.


    pod install

Note: You should receive confirmation that these dependencies have been added to the project.  If there is a syntax error in the Podfile, you will encounter an error when you run the install command.

<a name="register-your-app-with-microsoft-azure"></a>
## Register your app with Microsoft Azure
1.	Sign in to the [Azure Management Portal](https://manage.windowsazure.com), using your Azure AD credentials.
2.	Click **Active Directory** on the left menu, then click the directory for your Office 365 developer site.
3.	On the top menu, click **Applications**.
4.	Click **Add** from the bottom menu.
5.	On the **What do you want to do page**, click **Add an application my organization is developing**.
6.	On the **Tell us about your application** page, specify **O365-iOS-Connect** for the application name and select **NATIVE CLIENT APPLICATION** for type.
7.	Click the arrow icon on the lower-right corner of the page.
8.	On the Application information page, specify a Redirect URI, for this example, you can specify http://localhost/connect, and then click the check box in the lower-right hand corner of the page. Remember this value for the section **Getting the ClientID and RedirectUri into the project**.
9.	Once the application has been successfully added, you will be taken to the Quick Start page for the application. From here, click Configure in the top menu.
10.	Under **permissions to other applications**, add the following permission: **Add the Office 365 Exchange Online** application. Next, click the check box in the bottom right corner to add the application. Finally, when you return to the **permissions to other applications** section, select **Send mail as a user** permission under **Delegated Permissions**.
13.	Copy the value specified for **Client ID** on the **Configure** page. Remember this value for the section **Getting the ClientID and RedirectUri into the project**.
14.	Click **Save** in the bottom menu.

<a name="get-the-client-id-and-redirect-uri-into-the-project"></a>
## Get the Client ID and Redirect Uri into the project

Finally you'll need to add the Client ID and Redirect Uri you recorded from the previous section **Register your app with Microsoft Azure**.

Browse the **O365-iOS-Connect** project directory and open up the workspace (O365-iOS-Connect.xcworkspace). In the **AuthenticationManager** file you'll see that the **ClientID** and **RedirectUri** values can be added to the top of the file. Supply the necessary values here:

    // You will set your application's clientId and redirect URI. You get
    // these when you register your application in Azure AD.
    static NSString * const REDIRECT_URL_STRING = @"ENTER_REDIRECT_URI_HERE";
    static NSString * const CLIENT_ID           = @"ENTER_CLIENT_ID_HERE";
    static NSString * const AUTHORITY           = @"https://login.microsoftonline.com/common";



<a name="code-of-interest"></a>
## Code of Interest

**Authentication with Azure AD**

The code for authenticating with Azure AD, which includes retrieval and management of your access tokens is located in AuthenticationManager.


**Outlook Services Client**

The code for creating your Outlook Services client is located in Office365ClientFetcher. The code in this file creates the Outlook Services client object required for performing API calls against the Office 365 Exchange service. The client requests will leverage the authentication code to get the application an access and refresh token to act on behalf of the user. The access and refresh token will be cached. The next time a user attempts to access the service, the access token will be issued. If the access token has expired, the client will issue the refresh token to get a new access token.


**Discovery Service**

The code for using the O365 Discovery service to retrieve the Exchange service endpoints/URLs is located in the internal method connectToOffice365 is called from the viewDidLoad method in SendMailViewController.


**Office 365 SendMail snippet**

The code for the operation to send mail is located in the sendMailMessage method in the SendMailViewController.


<a name="questions-and-comments"></a>
## Questions and comments

We'd love to get your input on this Office 365 iOS Connect sample. You can send your feedback to us in the [Issues](https://github.com/OfficeDev/O365-iOS-Connect) section of this repository. <br>
Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions are tagged with [Office365] and [API].

<a name="troubleshooting"></a>
## Troubleshooting

With the Xcode 7.0 update, App Transport Security is enabled for simulators and devices running iOS 9. See [App Transport Security Technote](https://developer.apple.com/library/prerelease/ios/technotes/App-Transport-Security-Technote/).

For this sample we have created a temporary exception for the following domain in the plist:

- outlook.office365.com

If these exceptions are not included, all calls into the Office 365 API will fail in this app when deployed to an iOS 9 simulator in Xcode.

<a name="additional-resources"></a>
## Additional resources

* [Office 365 code snippets for iOS](https://github.com/OfficeDev/O365-iOS-Snippets)
* [Office 365 Profile Sample for iOS](https://github.com/OfficeDev/O365-iOS-Profile)
* [Email Peek - An iOS app built using Office 365](https://github.com/OfficeDev/O365-iOS-EmailPeek)
* [Office 365 APIs documentation](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [File REST operations reference](http://msdn.microsoft.com/office/office365/api/files-rest-operations)
* [Calendar REST operations reference](http://msdn.microsoft.com/office/office365/api/calendar-rest-operations)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 API code samples and videos](https://msdn.microsoft.com/office/office365/howto/starter-projects-and-code-samples)

