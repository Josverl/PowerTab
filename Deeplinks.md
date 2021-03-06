# Creating Deeplinks to your Tab 

## Deeplink to a Channel tab including a sub-entity ID 

Note : the Entity ID is defined by the configuration page.
In this sample that is a static ID 

**Onsubmit of a Button** 
``` PowerApps
// deeplink to a configurable tab with a SubEntityId 

//  build the context that we want to pass as part of the deeplink  
UpdateContext({
    locContext: "{""subEntityId"":""Item_42""" & 
                        ",""channelId"":""" & TeamsCtx.channelId & """}" 
});

// https://teams.microsoft.com/l/entity/<teamsAppId>/<entityId>?webUrl=<entityWebUrl>&label=<entityLabel>&context=<context>
// the teams AppId is the Teams Manifest AppId, 
// or the Teams Internal AppId that can be found using the graph API 

UpdateContext({
   locDeepLink : 
        "https://teams.microsoft.com/l/entity/" & TeamsCtx.teamsAppID & "/" & TeamsCtx.entityId 
        & "?webUrl=" & EncodeUrl( "https://powertab.azurewebsites.net/") 
        & "&label=" & EncodeUrl("Item 42") 
        & "&context=" & EncodeUrl( locContext)
});

// post the adaptive card via Power Automate  
PowerAppsbutton_1.Run( 
    txtMessage,
    locDeepLink,
    TeamsCtx.groupId,
    TeamsCtx.channelId)
```

## Detecting and acting on a 'incoming link' that contains a subentity ID 
Requirement: configure your Tab/bot to pass the subentityIDs in an adaptive card or other link.

**App.OnStart**
```
// if a subentity was provided, go directly to that screen 
If( Len( TeamsCtx.subEntityId ) > 0 , Navigate(subEntityScreen, ScreenTransition.None));

```
