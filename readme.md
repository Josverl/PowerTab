#powerapp 

# onStart or onLoad

Note:  the parameters names are case sensitive 
``` powerapp
Set( glbsource, Coalesce(Param("source"), "unknown") );
Set( glbTeamId, Coalesce(Param("teamId"), "unknown") );
Set( glbChannelID,Coalesce( Param("channelId"), "unknown") );
Set( glbUPN,Coalesce( Param("upn") , "unknown"));
Set( glbTenantID,Coalesce( Param("tenantId"), "unknown") );

Set( glbTheme,Coalesce( Param("theme") , "unknown"));
Set( glbLoginHint,Coalesce( Param("loginHint") , "unknown"));
Set( glbTid,Coalesce( Param("tid"), "unknown") );

Set( glbTest,Coalesce( Param("test"), "unknown") );
```