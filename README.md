# Microsoft Graph Daemon

A small .NET Core console app demonstrating connecting to the Microsoft Graph as an application

## Setup

In order to run this app, you will need the following:

1. Azure Active Directory Tenant Id
1. Azure Actived Directory App Client Id
1. Azure Actived Directory App Client Secret
1. A SharePoint site set up with a document library

Your Azure AD Application will need application permissions to write to SharePoint, and will also need to be created using the Azure AD v2.0 endpoint. You can do this at the [Application Registration Portal](apps.dev.microsoft.com).

This application will create an excel document and upload it to the specified document library.

Update the appsettings.json file with your settings, and in the Program.cs file specify the name of your new document and the name of your document library.

## Run

Once you have updated the settings, all you need to do is hit F5 to run.

## MSAL

This uses the Microsoft.Identity.Client 1.1.4-preview0002 pulled in from Nuget, although this is labelled as a preview. Microsoft states that this is suitable for production use.

From MSAL [GitHub](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet):

> This library is suitable for use in a production environment. We provide the same production level support for this library as we do our current production libraries. During the preview we may make changes to the API, internal cache format, and other mechanisms of this library, which you will be required to take along with bug fixes or feature improvements. This may impact your application. For instance, a change to the cache format may impact your users, such as requiring them to sign in again. An API change may require you to update your code. When we provide the General Availability release we will require you to update to the General Availability version within six months, as applications written using a preview version of library may no longer work.

This also uses the class MSALCache which has been pulled from the Microsoft sample repository from the Microsoft Daemon example on github here: [https://github.com/Azure-Samples/active-directory-dotnet-daemon-v2](https://github.com/Azure-Samples/active-directory-dotnet-daemon-v2)
