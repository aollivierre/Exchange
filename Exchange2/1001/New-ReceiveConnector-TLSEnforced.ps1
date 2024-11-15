﻿New-ReceiveConnector -Name "Dedicated TLS for coughlin.ca" `
-Bindings 0.0.0.0:2526 `
-RemoteIPRanges 13.107.6.152/31,13.107.18.10/31,13.107.128.0/22,23.103.160.0/20,40.96.0.0/13,40.104.0.0/15,52.96.0.0/14,131.253.33.215/32,132.245.0.0/16,150.171.32.0/22,204.79.197.215/32,40.92.0.0/15,40.107.0.0/16,52.100.0.0/14,52.238.78.88/32,104.47.0.0/17 `
-AuthMechanism TLS `
-PermissionGroups AnonymousUsers `
-RequireTLS $true