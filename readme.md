# PowerTab, (a PowerApp better integrated in a Teams Tab)

Generating a Teams app [from a PowerApp is simple](https://docs.microsoft.com/en-us/powerapps/maker/canvas-apps/embed-teams-app), right ?
unfortunatly after adding such a Teamsified PowerApp to your Team / channel, the Powerapp is pretty static, and it is not aware of the Team or Channel it is pinned to, and as such you may need to add additional configuration and navigation,
or perhaps even create and mantain multiple copies of your PowerApp.

This can be done a little smarter, by configuring Teams to pass the relevant information as query string parameters to your powerapp.
youten can read these paramaters in your powerapp, and then use them as you need.

relevant context information includes: 
* the Teams ID and Display Name 
* the channel ID and Display Name
* the theme ( dark / light / accessible) 


General steps :
1. Create a powerapp 
2. Publish it
3. Change the Teams app manifest to use the PowerTab configuration page
4. Use the PowerTab configuration page to configure which parameters you want to pass to your PowerApp
5. Retrieve the teams context parameters in PowerApp 
6. Use this Teams context in your PowerApp to drive an even better user and collaboration experience      


# Live powertab site 
for demos and Proof of concepts feel free to use the live configuration page located at:
https://powertab.azurewebsites.net/config.html


## Teams manifest - Configurable Tab  

You should create a teams manifest that uses the PowerTab config page hosted as the PoC or to your own copy.

configuration page: `https://powertab.azurewebsites.net/config.html?PowerAppId=<ID_Of_Your_PowerApp>`

Note: please do not use the PoC page for production as it has limited capacity and there may be changes without no prior notice. 

## Powerapp onStart or onLoad

In powerApp you can retrieve the teams context, and store them in global variables 

below is the simples form, or alternativly you can collect them as a single glbTeamsContext record.  

``` powerapp
Set( glbsource, Coalesce(Param("source"), "unknown") );
Set( glbTeamId, Coalesce(Param("teamId"), "unknown") );
Set( glbChannelID,Coalesce( Param("channelId"), "unknown") );
Set( glbUPN,Coalesce( Param("upn") , "unknown"));
Set( glbTenantID,Coalesce( Param("tenantId"), "unknown") );

Set( glbTheme,Coalesce( Param("theme") , "unknown"));
Set( glbLoginHint,Coalesce( Param("loginHint") , "unknown"));
Set( glbTid,Coalesce( Param("tid"), "unknown") );

```
Note:  the parameters names are case sensitive 

# details 

## All  parameters 


Parameters :
 - "teamName": "The name of the current team",
 - "channelId": "The channel ID in the format 19:[id]@thread.skype",
 - "channelName": "The name of the current channel",
 - "chatId": "The chat ID in the in the format 19:[id]@thread.skype",
 - "locale": "The current locale of the user formatted as languageId-countryId (for example, en-us)",
 - "entityId": "The developer-defined unique ID for the entity this content points to",
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
- camelCasing may lead to mistakes, as the url string parameters are Case sensative.
