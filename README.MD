# Office 365 Code Snippets for Android

**Table of contents**

* [Change History](#change-history)
* [Device requirements](#device-requirements)
* [Prerequisites](#prerequisites)
* [Open the project using Android Studio](#open-the-project-using-android-studio) 
* [Run the project](#run-the-project)
* [Understand the code](#understand-the-code)
* [Connnect to Office 365](#connect-to-office-365)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

This sample for Android gives you the common code routines that you write for access to email, calendar events, contacts, and files in Office 365. These code routines known as "snippets" are simple and self contained methods that you can copy and paste into your own code or use as reference for learning. This sample has snippets for the most common create, read, update, and delete operations on Office 365 content.  

This sample includes:

**Calendar**

* Get existing calendar events by date range.
* Create an event.
* Update an event.
* Delete an event.
* Accept an event invitation.
* Decline an event invitation.
* Get the status of an event attendee.

**OneDrive for Business Files**
 
* Get existing files and folders.
* Create text files.
* Delete files or folders.
* Read text file contents.
* Update text file contents.
* Upload or download files.

**Mail**

* Get existing mail items.
* Get mail items by subject.
* Delete mail item.
* Send new mail.
* Forward email.
* Reply to email.

**Contacts**

* Get existing contacts.
* Create a contact.
* Update a contact.
* Delete a contact.

We will continue adding snippets to this sample, so make sure to check back.

## Change History
April 6, 2015:

* Initial release


## Device requirements
To run the snippets project your device needs to meet the following requirements:

* A tablet with a screen size of 7 inches or larger.
* Android API level 16 or later.
 

## Set up your environment

### Prerequisites

To use the Office 365 Code Snippets for Android you need the following:

* The latest version of [Android Studio](http://developer.android.com/sdk/index.html).
* The [Gradle](http://www.gradle.org) build automation system version 2.2.1 or later.
* An Office 365 account. You can sign up for [an Office 365 Developer subscription](https://portal.microsoftonline.com/Signup/MainSignUp.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK) that includes the resources that you need to start building Office 365 apps.
* JDK 7
* 
* A registered Azure application with a client id and redirect uri value. The application must have the following permissions:
	* Windows Azure Active Directory
		* Enable sign-on and read users' profiles
	* Office 365 SharePoint Online
		* Edit or delete users' files
	* Office 365 Exchange Online
		* Read and write access to users' mail
		* Send mail as a user
		* Have full access to users' calendars
		* Have full access to users' contacts



## Open the project using Android Studio
Open the project by importing the settings.gradle file.
1. Install [Android Studio](http://developer.android.com/tools/studio/index.html#install-updates) and add the Android SDK packages according to the [instructions](http://developer.android.com/sdk/installing/adding-packages.html) on developer.android.com.
2. Download or clone the
   [Office 365 Code Snippets for Android](https://github.com/OfficeDev/O365-Android-Snippets).
3. Start Android Studio.
	1. From the **Welcome to Android Studio** dialog, choose **Import project (Eclipse ADT, Gradle, etc.)**. 
	2. Select the **settings.gradle** file in the **O365-Android-Snippets** folder and click **OK**.
4. Open the constants.java file in the com.microsoft.office365.snippetapp.helpers package.
	1. Find the CLIENT_ID constant and set its String value equal to the client id you registered in Azure Active Directory.
	2. Find the REDIRECT_URI constant and set its String value equal to the redirect URI you registered in Azure Active Directory.

> Note: The starter project declares the required dependencies using Gradle. The dependencies are:
> 
> * The [Azure Active Directory Authentication Library for Android](https://github.com/AzureAD/azure-activedirectory-library-for-android).
> * The [Office 365 SDK for Android](https://github.com/OfficeDev/Office-365-SDK-for-Android).



## Run the project

Once you've built the snippets project you can run it on an emulator or device.

1. Run the project.
3. Click the "Connect" item on the action bar overflow menu and enter your credentials.
4. Click on a task in the main activity list to run a task.
5. View the results of the task in the detail pane to the right of the list.
6. Optionally, clear the results of the task by clicking on the "Clear Results" action bar item.
7. Run all tasks in the list sequentially by clicking on the "Run All" action bar item.
  

## Understand the code

The starter project uses four objects, **CalendarSnippets**, **ContactsSnippets**, **EmailSnippets**, and **FileFolderSnippets** to manage interactions with Office 365. These objects have snippet methods that wrap calls to the **SharePointClient** and **OutlookClient** objects in the Office 365 SDK for Android. Look at the classes in the **com.microsoft.office365.snippetapp.O365Stories** package to see how the snippets are called. 


### Connect to Office 365

The Office 365 SDK for Android uses the Azure Active Directory Library (ADAL) for Android for connecting your app to Office 365. The ADAL provides protocol support for OAuth2, Web API integration with user level consent, and two-factor authentication.

The **AuthenticationController** object manages getting a token from ADAL and returning it to your application.

### Calendar API

The **CalendarAPIOperations** object wraps the API operations that create, update and delete calendar events in an Office 365 Exchange calendar. 

The **readEventsFromOutlook** method gets a list of events from the Office 365 calendar and loads pages of events into a local list. Changes, deletions, and additions to this list are posted asynchronously to the Office 365 calendar by the **pushUpdatedEventToOutlook**, **pushDeletedEventToOutlook**, and **pushCreatedEventToOutlook** methods.

### Files API

The **FilesFolderAPIOperations** object wraps the API operations that create, update, and delete files stored on OneDrive for Business.

The **getFilesAndFoldersFromService** method gets a list of all of the files and folders that are stored on OneDrive for Business and loads the list into a local array. Changes, deletions, and additions to the local list of files are posted asynchronously to OneDrive for Business by the **postUpdatedFileContents**,  **postUploadFileToServer**, **postDeleteSelectedFileFromServer**, and **postNewFileToServer** methods. 

The **getFilesAndFoldersFromService** method returns an **FilesFolderModel** object containing the selected files contents.

### Mail API

The **MailAPIOperations** object wraps the API operations that create, update and delete mail items in an Office 365 Exchange mailbox. 

The **getMessagesFromOutlook** method gets a list of mail items from the Office 365 mailbox and loads pages of items into a local list. Deletions to this list, and mail sent from the mailbox are posted asynchronously to the Office 365 calendar by the **pushDeleteMailItemToOutlook**, and **pushNewMailToOutlook** methods.

### Contacts API

The **ContactsAPIOperations** object wraps the API operations that create, update and delete contacts in Office 365 Exchange. 

The **getContactsFromOutlook** method gets a list of contacts from Office 365 and loads pages of contacts into a local list. Changes, deletions, and additions to this list are posted asynchronously to Office 365 by the **postUpdatedEvent**, **postDeletedEvent**, and **postCreatedEvent** methods.

## Questions and comments

We'd love to get your feedback on this Android starter kit. You can send your questions and suggestions to us:

* In the [Issues](https://github.com/OfficeDev/O365-Android-Start/issues) section of this repository.
* Send us an email to [docthis@microsoft.com](mailto:docthis@microsoft.com?subject=Feedback%20on%20the%20Office%20365%20Starter%20project%20for%20Android).
  
<a name="resources"/>
## Additional resources

* [Office 365 APIs documentation](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [File REST operations reference](https://msdn.microsoft.com/office/office365/api/files-rest-operations)
* [Calendar REST operations reference](http://msdn.microsoft.com/office/office365/api/calendar-rest-operations)
* [Mail REST operations reference](https://msdn.microsoft.com/office/office365/api/mail-rest-operations)
* [Microsoft Office 365 API Tools](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs starter projects and code samples](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)


## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
Microsoft and OneDrive are either registered trademarks or trademarks of Microsoft Corporation.

