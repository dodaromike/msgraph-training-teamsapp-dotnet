---
page_type: sample
description: This sample demonstrates how to use the Microsoft Graph .NET SDK to access data in Office 365 from Microsoft Teams apps.
products:
- ms-graph
- microsoft-graph-calendar-api
- office-exchange-online
languages:
- csharp
- dotnet
---

# Microsoft Graph sample Microsoft Teams app

This sample demonstrates how to use the Microsoft Graph .NET SDK to access data in Office 365 from Microsoft Teams apps.

## Prerequisites

Before you start this tutorial, you should have the following installed on your development machine.

- [.NET SDK](https://dotnet.microsoft.com/download).
- [ngrok](https://ngrok.com/)

You should also have a Microsoft work or school account in a Microsoft 365 tenant that has [enabled custom Teams app sideloading](/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant#enable-custom-teams-apps-and-turn-on-custom-app-uploading). If you don't have a Microsoft work or school account, or your organization has not enabled custom Teams app sideloading, you can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 developer subscription.

### Run ngrok

Microsoft Teams does not support local hosting for apps. The server hosting your app must be available from the cloud using HTTPS endpoints. For debugging locally, you can use ngrok to create a public URL for your locally-hosted project.

1. Open CLI in the project direcory, and run 
    ```Shell
    dotnet run
    ```
 1. Note 'Now listening on:' port number in the output:
    ```
    Building...
    info: Microsoft.Hosting.Lifetime[14]
        Now listening on: https://localhost:7209
    info: Microsoft.Hosting.Lifetime[14]
        Now listening on: http://localhost:5274
    info: Microsoft.Hosting.Lifetime[0]
        Application started. Press Ctrl+C to shut down.
    info: Microsoft.Hosting.Lifetime[0]
        Hosting environment: Development
    info: Microsoft.Hosting.Lifetime[0]
    ```

1. Open your CLI and run the following command to start ngrok.

    ```Shell
    ngrok http 5274
    ```

1. Once ngrok starts, copy the HTTPS Forwarding URL. It should look like `https://50153897dd4d.ngrok.io`. You'll need this value to .

> [!IMPORTANT]
> If you are using the free version of ngrok, the forwarding URL changes every time you restart ngrok. It's recommended that you leave ngrok running until you complete this tutorial to keep the same URL. If you have to restart ngrok, you'll need to update your URL everywhere that it is used and reinstall the app in Microsoft Teams.

## Register the app

In this exercise, you will create a new Azure AD web application registration using the Azure Active Directory admin center.

1. Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.

1. Select **New registration**. On the **Register an application** page, set the values as follows, where `YOUR_NGROK_URL` is the ngrok forwarding URL you copied in the previous section.

    - Set **Name** to `Teams Graph Tutorial`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Web` and set the value to `YOUR_NGROK_URL/authcomplete`.

1. Select **Register**. On the **Teams Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.

1. Select **Authentication** under **Manage**. Locate the **Implicit grant** section and enable **Access tokens** and **ID tokens**. Select **Save**.

1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and select **Add**.

1. Copy the client secret value before you leave this page. You will need it in the next step.

    > [!IMPORTANT]
    > This client secret is never shown again, so make sure you copy it now.

1. Select **API permissions** under **Manage**, then select **Add a permission**.

1. Select **Microsoft Graph**, then **Delegated permissions**.

1. Select the following permissions, then select **Add permissions**.

    - **Calendars.ReadWrite** - this will allow the app to read and write to the user's calendar.
    - **MailboxSettings.Read** - this will allow the app to get the user's time zone, date format, and time format from their mailbox settings.

### Configure Teams single sign-on

In this section you'll update the app registration to support [single sign-on in Teams](/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso).

1. Select **Expose an API**. Select the **Set** link next to **Application ID URI**. Insert your ngrok forwarding URL domain name (with a forward slash "/" appended to the end) between the double forward slashes and the GUID. The entire ID should look similar to: `api://50153897dd4d.ngrok.io/ae7d8088-3422-4c8c-a351-6ded0f21d615`.

1. In the **Scopes defined by this API** section, select **Add a scope**. Fill in the fields as follows and select **Add scope**.

    - **Scope name:** `access_as_user`
    - **Who can consent?: Admins and users**
    - **Admin consent display name:** `Access the app as the user`
    - **Admin consent description:** `Allows Teams to call the app's web APIs as the current user.`
    - **User consent display name:** `Access the app as you`
    - **User consent description:** `Allows Teams to call the app's web APIs as you.`
    - **State: Enabled**

1. In the **Authorized client applications** section, select **Add a client application**. Enter a client ID from the following list, enable the scope under **Authorized scopes**, and select **Add application**. Repeat this process for each of the client IDs in the list.

    - `1fec8e78-bce4-4aaf-ab1b-5451cc387264` (Teams mobile/desktop application)
    - `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Teams web application)

## Running the sample

### Create Manifest and Add app using Teams Developer Portal.
-TBD-

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
