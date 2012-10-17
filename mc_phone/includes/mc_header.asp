<%
'Response.ExpiresAbsolute=#July 7,2001 12:00:00#
'Response.Expires = -1
'Response.AddHeader "Pragma", "no-cache"
'Response.AddHeader "cache-control", "no-cache, must-revalidate"
Response.Buffer = True
'Response.AddHeader "X-Pb-CacheOn", "0"
'Response.AddHeader "X-Pb-Flush", "1"
'Response.AddHeader "X-Pb-SendCache", "no"
'Response.AddHeader "X-Pb-CompressionLevel", "0"
Dim SessionID,SessionVars,UpdateVars
SetAppVarsFromDB()
GetContext
%>