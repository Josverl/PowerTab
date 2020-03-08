# PowerTab, (a PowerApp better integrated in a Teams Tab)

Generating a Teams app [from a PowerApp is simple](https://docs.microsoft.com/en-us/powerapps/maker/canvas-apps/embed-teams-app), right ?  
Unfortunately after adding such a `Teamsified` PowerApp to your Team and Channel, the resulting _Teams Power App_ is pretty static, and it is not aware of the Team or Channel (or Chat) it is pinned to.
And as such you may need to add additional configuration and navigation, or perhaps even create and maintain multiple copies of your PowerApp.

This can be done a little smarter, by configuring Teams to pass the relevant information as query string parameters to your Power App
using [url placeholder values](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context#getting-context-by-inserting-url-placeholder-values).  
Note:  Additional context parameters are/may be documented as part of the [javascript sdk Teams Context](https://docs.microsoft.com/en-us/javascript/api/@microsoft/teams-js/microsoftteams.context?view=msteams-client-js-latest). 

Then you can add a little logic to the PowerApp to read and make use of this information.

This can be achieved by using a custom configuration page, that takes the PowerApp Application ID, and performs the relevant configuration so that Teams will pass the relevant information to your PowerApp.
This page is only involved during the initial configuration, after this its just Teams and your PowerApp (with some additional information)  
The relevant context information includes this such as : 
* The Teams ID and Teams Display Name 
* The channel ID and Channel Display Name
* The theme (dark, light, accessible) 

## Use cases 

- Display the Teams and Channel name
- use the Theme to determine when to render a high contrast UX  
- use the Teams ID 
- Use the entity ID in a deeplink to show information regarding a specific  Mention 

## General steps :
1. Create a PowerApp
2. Save & Publish your PowerApp 
3. Download the Teams manifest for your PowerApp 
4.  Import and Open the Teams App Manifest in [Teams App Studio](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/app-studio-overview#manifest-editor) 
5. Change the Teams app manifest to to change the configurable tab as in the below section
6. Use the PowerTab configuration page to configure which parameters you want to pass to your PowerApp
7. Retrieve the teams context parameters in PowerApp 
8. Use this Teams context in your PowerApp to drive an even better user and collaboration experience      

One example it to use the teams Context information to create a [Deeplink](Deeplinks.md) to your tab,  
that you can send in a adaptive card. That deeplink will allow the recipient to return directly to your tab, and can optionally include relevant information on their action / response in the subEntityID

## Teams manifest - Team tab (a.k.a. Configurable Tab) 

You should create or change a Teams manifest to use the PowerTab config page hosted as the PoC or to your own copy.
Configuration page: `https://powertab.azurewebsites.net/config.html?PowerAppId=<ID_Of_Your_PowerApp>`

after the change it should look something like this:  
![manifest editor showing Tab configuration](img/configured_teams_tab.png)

Example [Manifest](manifest/manifest.json)


# Live PowerTab site 
for demos and Proof of concepts feel free to use the live configuration page located at:
https://powertab.azurewebsites.net/config.html

**Note: If you use the PoC page, please be aware it has limited capacity and may change or stop working with no prior notice.**
## PowerApp onStart or onLoad

In PowerApp you can retrieve the Teams context, and store them in a global variable using the Param function `Param("TeamId")`.
As there may be cases when the PowerApp is started outside of Teams, you can use the Coalesce function to supply a sensible default  

Below is a sample to store the teams context in a single global variable (a record) 
so that you can retrieve the Teams Displayname as `glbTeamsContext.teamName`

`PowerApp App.OnStart: `
``` PowerApp 
// If specified, the Teams App ID can be used to generate deeplinks to this Tab
Set(TeamsAppID, "12345678-1234-1234-1234-123456789ABC");

// get Teams context in App.OnStart as a single record / variable
Set(
    glbTeamsContext,
    {
        source: Coalesce(Param("source"), "web"),
        groupId: Coalesce(Param("groupId"), "00000000-0000-0000-0000-000000000000"),
        teamId: Coalesce(Param("teamId"), "19:[team-id]@thread.skype"),
        channelId: Coalesce( Param("channelId"), "19:[channel-id]@thread.skype"),
        teamName: Coalesce( Param("teamName"), "Team unknown"),
        channelName: Coalesce( Param("channelName"), "Channel unknown"),
        chatId: Coalesce( Param("chatId"), "19:[chat-id]@thread.skype"),

        theme: Coalesce( Param("theme"), "light"),

        channelType: Coalesce( Param("channelType"), ""),
        teamSiteUrl: Coalesce( Param("teamSiteUrl"), ""),

        locale: Coalesce( Param("locale"), ""),
        entityId: Coalesce( Param("entityId"), ""),
        subEntityId: Coalesce( Param("subEntityId"), ""),
        
        tid: Coalesce( Param("tid"), ""),
        isFullScreen: Coalesce( Param("isFullScreen"), ""),
        userLicenseType: Coalesce( Param("userLicenseType"), ""),
        tenantSKU: Coalesce( Param("tenantSKU"), "")
    }
); 
```
_Note: The parameters names are case sensitive._


**Alternatively you can just retrieve a few specific parameters.**
``` PowerApp 
// get Teams context from the QueryString in App.OnStart
Set( glbsource, Param("source"));
Set( glbTeamId, Param("teamId"));
Set( glbChannelID, Param("channelId"));
Set( glbTheme, Param("theme") , "light"));

```



## Best option to create a PowerApp for embedding in Teams 

for best results : 
- Start in a Teams channel
- Add a Tab,
- Select PowerApps 
- Click the link at the bottom : `Or create an app in Power Apps` 
![image of link in Teams to create PowerApp ](./img/create_in_PowerApps.png)

This will start a PowerApp that has the additional benefits of 
- A **Teams Purple theme** for your PowerApp 
- **Automatic scaling / adjustment** of the screen size

# The Details 

## Teams Tab configuration 

The following configuration is saved for the Tab: 
 - **contentUrl: getUrl(idPowerApp, true)**  
   This is the complete url to run the PowerApp in context of Teams, so this includes all the Teams replacement tokens that have been selected.  
    - contentUrl: `https://apps.powerapps.com/play/providers/Microsoft.PowerApps/apps/b78d9c7e-9e44-4217-9cf3-f46dd4be475f?source=teamstab&groupId={groupId}&teamId={teamId}&channelId={channelId}&chatId={chatId}&theme={theme}&teamSiteUrl={teamSiteUrl}` 


 - **websiteUrl: getUrl(idPowerApp, false)**  
   This is the complete url to run the PowerApp in a browser. This does not includes any of the Teams context replacement tokens, as they would not be replaced with the actual teams context. To allow an App to determine how it is opened, the only query string parameter provided is : `source=teamsopenwebsite`

    - websiteUrl: `https://apps.powerapps.com/play/providers/Microsoft.PowerApps/apps/b78d9c7e-9e44-4217-9cf3-f46dd4be475f?source=teamsopenwebsite`

 - **entityId: idPowerApp**  
    the PowerApp ID (a GUID)

 - suggestedDisplayName: "PowerTab"  
    The default name for the Tab 


## Most supported URL parameters 

Parameters :
 - "teamName": "The name of the current team",
 - "channelId": "The channel ID in the format 19:[id]@thread.skype",
 - "channelName": "The name of the current channel",
 - "chatId": "The chat ID in the in the format 19:[id]@thread.skype",
 - "locale": "The current locale of the user formatted as languageId-countryId (for example, en-us)",
 - "entityId": "The developer-defined unique ID for the entity this content points to. (the PowerApp ID )",
 - "subEntityId": "The developer-defined unique ID for the sub-entity this content points to",
 - "loginHint": "A value suitable as a login hint for Azure AD. This is usually the login name of the current user, in their home tenant",
 - "userPrincipalName": "The User Principal Name of the current user, in the current tenant",
 - "userObjectId": "The Azure AD object id of the current user, in the current tenant",
 - "tid": "The Azure AD tenant ID of the current user",
 - "groupId": "Guid identifying the current O365 Group ID",
 - "theme": "The current UI theme: default | dark | contrast",
 - "isFullScreen": "Indicates whether the tab is in full-screen mode",
 - "userLicenseType": "Indicates the user licence type in the given SKU (for example, student or teacher)",
 - "tenantSKU": "Indicates the SKU category of the tenant (for example, EDU)",
 - "channelType": "microsoftTeams.ChannelType.Private | microsoftTeams.ChannelType.Regular" -->

Notes:  
- Parameter names are the same for Teams as for the PowerApp
- camelCasing may lead to mistakes, as the url string parameters are Case sensitive.

#### References:

- [Teams Context URL Placeholders](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context#getting-context-by-inserting-url-placeholder-values)
- [Teams Context URL Placeholders - Private Channels](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context#retrieving-context-in-private-channels)
- [Teams SDK - Context](https://docs.microsoft.com/en-us/javascript/api/@microsoft/teams-js/microsoftteams.context?view=msteams-client-js-latest)

