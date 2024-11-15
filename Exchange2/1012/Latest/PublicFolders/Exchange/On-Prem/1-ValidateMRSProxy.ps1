# https://learn.microsoft.com/en-us/exchange/architecture/mailbox-servers/mrs-proxy-endpoint?view=exchserver-2019

Get-WebServicesVirtualDirectory | Format-Table -Auto Identity,MRSProxyEnabled



# Identity                        MRSProxyEnabled
# --------                        ---------------
# GLB-EX01\EWS (Default Web Site)            True
