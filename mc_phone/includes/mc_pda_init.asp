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
Dim LookupSize

'====================================================================================================
' SET APPLICATION VERSION INFORMATION
'====================================================================================================
If Not Application("MCVersion") = "3.0" Then
	Application.Lock
		Application("MCVersion") = "3.0"
	Application.UnLock
End If

'====================================================================================================
' WHAT DEVICE ARE WE DEALING WITH?
'====================================================================================================
ua = UCase(Request.ServerVariables("HTTP_USER_AGENT"))
'Call OutputWAPError(ua)

IsLargeScreen = False
IsWindows = True
IsCE = False
IsPocketIE = False
IsPPC = False
IsSmartPhone = False
IsBlackBerry = False
IsWAP20 = False

'If InStr(ua,"WINDOWS") > 0 or InStr(ua,"IPHONE") or InStr(ua,"IPOD") > 0 or InStr(ua,"ITOUCH") > 0 or InStr(ua,"PLAYSTATION") > 0 Then
'	IsWindows = True
'End If

Lang = "HTML"
IsLargeScreen = True
GlobalPageSize = 20
GlobalFieldLength = "16"
LookupSize = 5

HStyleBegin = "<div style='font-size:14pt; font-weight:bold;' class='Font1'>"
HStyleEnd = "</div>"

LStyleBegin = "<div class='Font1' style='font-size:12pt;'>"
LStyleEnd = "</div>"
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

Dim iPhone
iPhone = False

If InStr(ua,"IPHONE") > 0 or InStr(ua,"ITOUCH") > 0 or InStr(ua,"IPOD") > 0 or InStr(ua,"PLAYSTATION") > 0 Then
    iPhone = True
End If

If LiveMRO or MROGadget or MROPortlet or iPhone Then
	MROLive = True
Else
	MROLive = False
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

If card = "" Then
	If (Not Request.QueryString("m") = "") and (Not Request.QueryString("p") = "") Then
		txtmembername = Server.HTMLEncode(Trim(Request("m")))
		txtpassword = Server.HTMLEncode(Trim(Request("p")))
		card = "authenticate"
	ElseIf (Lang = "HTML") and (Not Request.Cookies("m") = "") and (Not Request.Cookies("p") = "") Then
		txtmembername = Request.Cookies("m")
		txtpassword = Request.Cookies("p")
		card = "authenticate"
	Else
		card = "login"
	End If
End If
%>