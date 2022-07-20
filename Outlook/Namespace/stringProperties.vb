option explicit

sub main()

    dim ns        as outlook.namespace
    dim connMode  as olExchangeConnectionMode
    dim connMode_ as string

    set ns   = application.getNamespace("MAPI")

    connMode = ns.exchangeConnectionMode

    if     connMode    = olCachedConnectedDrizzle  then
           connMode_   ="olCachedConnectedDrizzle"
    elseif connMode    = olCachedConnectedFull     then
           connMode_   ="olCachedConnectedFull"
    elseif connMode    = olCachedConnectedHeaders  then
           connMode_   ="olCachedConnectedHeaders"
    elseif connMode    = olCachedDisconnected      then
           connMode_   ="olCachedDisconnected"
    elseif connMode    = olCachedOffline           then
           connMode_   ="olCachedOffline"
    elseif connMode    = olDisconnected            then
           connMode_   ="olDisconnected"
    elseif connMode    = olNoExchange              then
           connMode_   ="olNoExchange"
    elseif connMode    = olOffline                 then
           connMode_   ="olOffline"
    elseif connMode    = olOnline                  then
           connMode_   ="olOnline"
    end if


    debug.print "currentUser                 : " &ns.currentUser
    debug.print "currentProfileName          : " &ns.currentProfileName
    debug.print "defaultStore                : " &ns.defaultStore
    debug.print "exchangeConnectionMode      : " &   connMode_
'   debug.print "exchangeMailboxServerName   : " &ns.exchangeMailboxServerName
    debug.print "exchangeMailboxServerVersion: " &ns.exchangeMailboxServerVersion
    debug.print "offline                     : " &ns.offline
    debug.print "type                        : " &ns.type


end sub
