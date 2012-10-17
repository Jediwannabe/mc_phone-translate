<%
'====================================================================================================================================
' MC MOBILE WAP
' Copyright (c) 2004-2005 Maintenance Connection Inc.
'====================================================================================================================================

Dim db, pdb, sql, PostUpdateSQL, con, rs, cmd, r, errormessage, errortext, ua, card, cardtitle, cardcurrent, cardcurrentlevel, cardfrom, cardfromlevel, cardlevelupdownamount, CardSkipLevel, newcontextpage
Dim ConnectionKey, ContainerGuid, ContainerCode, ContainerTypeCode, DBServerName, UserGuid, Password
Dim txtmembername, txtpassword, launchit 
Dim searchby, searchvalue, sqlwhere, sqlewhere, wopk, assetpk, pk, newrecord
Dim IsWindows, IsCE, IsPocketIE, IsPPC, IsSmartPhone, IsBlackBerry, IsWAP20
Dim PageSize,PagePos,TabIndex,ChoiceIndex, GlobalFieldLength
Dim FN,FT,POSTFIELDS
Dim debug, HeaderMSG, isIframe, returnurl
Dim IsBOF, IsEOF, IsOpen

Dim retval,retval2,rsdata1,rsdata1recs,rsdata2,rsdata2recs,rsdata3,rsdata3recs,rsdata4,rsdata4recs
Dim rscounter,retvalcounter,ei
Dim IsLargeScreen, GlobalPageSize

Dim HStyleBegin, HStyleEnd
Dim LStyleBegin, LStyleEnd
Dim Lang, SetVars, SetVarsDebug, Fields, FieldsB
Dim LiveMRO, MROGadget, MROPortlet
Dim MROLive, MCVirtualDirectory

'====================================================================================================
' SET APPLICATION VERSION INFORMATION
'====================================================================================================
If Not Application("MCVersion") = "3.0" Then
	Application.Lock
		Application("MCVersion") = "3.0" 
	Application.UnLock
End If

LStyleBegin = "<font style=""font-size:9pt;"">"
LStyleEnd = "</font>"

r = ""
%>

<!--#INCLUDE FILE="mc_all.asp" --> 

<%
launchit = False
errortext = ""
postfields = ""
PostUpdateSQL = ""
HeaderMSG = ""
SetVars = ""
Fields = ""
FieldsB = ""
TabIndex = 0
IsBOF = False
IsEOF = False
IsOpen = True
CardSkipLevel = 0
NewRecord = False

If (InStr(UCase(Request.ServerVariables("HTTP_HOST")),"LIVEMRO") > 0) or _
   (InStr(UCase(Request.ServerVariables("PATH_INFO")),"LIVEMRO") > 0) Then
   LiveMRO = True
Else
   LiveMRO = False
End If

If (InStr(UCase(Request.ServerVariables("HTTP_HOST")),"MROGADGET") > 0) or _
   (InStr(UCase(Request.ServerVariables("PATH_INFO")),"MROGADGET") > 0) Then
   MROGadget = True
Else
   MROGadget = False
End If

If (InStr(UCase(Request.ServerVariables("HTTP_HOST")),"MROPORTLET") > 0) or _
   (InStr(UCase(Request.ServerVariables("PATH_INFO")),"MROPORTLET") > 0) Then
   MROPortlet = True
Else
   MROPortlet = False
End If

Set db = New ADOHelper
Set pdb = New ADOHelper
pdb.oledbstr = Application("app_dsn")		

'card = "myworkorders"
card = Request("card")

cardfrom = GetSession("cardfrom")
If Not GetSession("cardfromlevel") = "" Then
	cardfromlevel = CInt(GetSession("cardfromlevel"))
End If

cardlevelupdownamount = 0
newcontextpage = False

'If Request.ServerVariables("remote_host") = "63.200.89.124" or _ 
   'Request.ServerVariables("remote_host") = "66.94.9.51" or _
   'Request.ServerVariables("remote_host") = "66.94.9.52" or _
   'Request.ServerVariables("remote_host") = "216.9.250.63" or _
   'Request.ServerVariables("remote_host") = "216.9.250.62" or _
   'Request.ServerVariables("remote_host") = "x209.115.248.157" or _
   'Request.ServerVariables("remote_host") = "BRAD209.183.48.110" or _
   'Request.ServerVariables("remote_host") = "x67.126.220.59" or _
   'Request.ServerVariables("remote_host") = "CSAIL65.198.3.2" or _     
   'Request.ServerVariables("remote_host") = "66.174.77.146" or _
   'Request.ServerVariables("remote_host") = "66.174.77.140" or _
   'Request.ServerVariables("remote_host") = "TOM63.78.116.140" Then
If True Then 
   ' We are good
Else
	'Write out Error
End If

If card = "" Then
	'If Lang = "HTML" Then
		If (Not Request.QueryString("m") = "") and (Not Request.QueryString("p") = "") Then
			txtmembername = Server.HTMLEncode(Trim(Request("m")))
			txtpassword = Server.HTMLEncode(Trim(Request("p")))
			card = "authenticate"		
		ElseIf (Lang = "WML") and (Not Request.Cookies("m") = "") and (Not Request.Cookies("p") = "") Then
            If IsPocketIE or IsBlackBerry Then		
    			card = "login"
            Else
			    txtmembername = Request.Cookies("m")
			    txtpassword = Request.Cookies("p")		
			    card = "authenticate"		
			End If
		ElseIf (Lang = "HTML") and (MROLive) and (Not Request.Cookies("m") = "") and (Not Request.Cookies("p") = "") Then
			txtmembername = Request.Cookies("m")
			txtpassword = Request.Cookies("p")		
			card = "authenticate"		
		Else
			card = "login"
		End If	
	'Else
	'	txtmembername = "cbucher"
	'	txtpassword = ""
	'	card = "authenticate"			
	'End If
End If
%>