<%
'====================================================================================================================================
' MC MOBILE WAP SUPPORTING FUNCTIONS
'====================================================================================================================================

'====================================================================================================================================
Dim backReferer

Sub StartMobileDocument(title)
%>	<html>
	<head>
	<title><% =title %></title>
<meta name="viewport" content="width=device-width,initial-scale=1, minimum-scale=1, maximum-scale=1,user-scalable=no" />
<meta name="apple-mobile-web-app-capable" content="yes">


	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
	<script src="js/jquery-1.6.min.js" type="text/javascript"></script>
	<script src="js/jquery-ui.min.js" type="text/javascript"></script>
	<script src="js/jquery.mc.js" type="text/javascript"></script>
	<link rel="stylesheet" href="css/jqueryui.css">
	<link rel="stylesheet" href="css/main.css">
	<script type="text/javascript">

		$(document).ready(function(){
			LoadDialogs();
		});

		$(window).resize(function () {

		});

	</script>
	</head>
	<body>
	<form name="mcform" method="get">
	<div name="maindiv" id="maindiv">
<%
	'debug = true
	'OutputWAPDebugHeader
	backReferer = Request.ServerVariables("http_referer")
	OutputSetVars
End Sub

Sub StartWMLCard(i,t)
End Sub

'====================================================================================================================================

Sub EndWMLCard()
	Call SetSession("sqlwhere" & CardCurrentLevel+1,"")
End Sub

'====================================================================================================================================

Sub EndWMLDocument()
%>
</div>
<div id="mcMainWindow">
	<div id="mcMainWindowLoading" class='loadingWindow'>
		<div style="position:relative; width:160px; left:50%; margin-left: -80px; top:25%; text-align:center;">
			<div><img src='images/logo_color_png.png' alt='' title='' /></div>
			<div style="padding-top:20px;"><img src="images/loading.gif" alt="" title="" /></div>
		</div>
	</div>
	<iframe id="mcMainWindowFrame" src="blank.htm" frameborder="0" scrolling="no" style="width:0px; height:0px;"></iframe>
</div>
</form>
</body>
</html>
<%
Response.End
End Sub

'====================================================================================================================================

Sub OutputWAPMsg(e)
	%>
	<div class='Font2' style='font-size:12pt;'>
	<% =e %>
	</div>
	<%
End Sub

'====================================================================================================================================

Sub OutputWAPDebug(e)
	If Debug Then
	%>
	<p>
	<small>
	<% =e %>
	</small>
	</p>
	<%
	End If
End Sub

'====================================================================================================================================

Sub OutputWAPDebugHeader()
	Call OutputWAPDebug(card)
End Sub

'====================================================================================================================================

Sub OutputWAPError(e)
	Call StartMobileDocument("")
	If IsPocketIE or IsBlackBerry Then
	If (lang="HTML") and (MROLive) Then %>
	<p align="center"><a href="http://www.maintenanceconnection.com" title="Maintenance Connection" target="_blank"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/mclive_login.gif" alt="MC Live"/></a></p><%
	Else %>
	<p align="center"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/mcmobile_login.gif" alt="MC Mobile"/></p><%
	End If
	End If %>
	<p><% If lang="HTML" Then %><font style="font-size:11pt;"><% End If %><% =e %><% If lang="HTML" Then %></font><% End If %></p>
	<p align="center">
	<a href="default.asp?card=login">Login</a>
	</p><%
	If lang = "WML" Then
	OutputBackButtonNative
	Else
	HTMLbuttonsBegin
	OutputBackButtonNative
	HTMLbuttonsEnd
	End If
	EndWMLDocument
	Response.End
End Sub

'====================================================================================================================================

Sub WAPbuttonsBegin()
    If IsPocketIE Then
        Response.Write "<p>"
    End If
End Sub

'====================================================================================================================================

Sub WAPbuttonsEnd()
    If IsPocketIE Then
        Response.Write "</p>"
    End If
End Sub

'====================================================================================================================================

Sub OutputButtonOrLinkStart(n,l)
    If n = "" Then
        n = "submitbutton"
    End If
    If l = "" Then
        l = "Submit"
    End If
    If IsPocketIE Then
        Response.Write "<anchor>"
    Else
        Response.Write "<do name=""" & n & """ type=""accept"" label=""" & l & """>"
    End If
End Sub

'====================================================================================================================================

Sub OutputButtonOrLinkEnd(l)
    If l = "" Then
        l = "Submit"
    End If
    If IsPocketIE Then
        Response.Write l & "</anchor> "
    Else
        Response.Write "</do>"
    End If
End Sub

'====================================================================================================================================

Sub rw(e)
	Call StartMobileDocument("")
	%>
	<p>
	<% =e %>
	</p>
	<%
	EndWMLDocument
End Sub

'====================================================================================================================================

Sub OutputBackButtonNative()
%>
<%
End Sub

'====================================================================================================================================

Sub CheckDB(byref db)
	If Not db.dok Then
		ErrorText = "There was a problem with your request: " & db.derror
		Set db = Nothing
		Call OutputWAPError(ErrorText)
	End If
End Sub

'====================================================================================================================================

Function ConvDate(strDate, strFormat)

	'%m Month as a decimal no. 02
	'%b Abbreviated month name Feb
	'%B Full month name February
	'%d Day of the month 23
	'%j Day of the year 54
	'%y Year without century 98
	'%Y Year with century 1998
	'%w Weekday as integer 5 (0 is Sunday)
	'%a Abbreviated day name Fri
	'%A Weekday Name Friday
	'%I Hour in 12 hour format 12
	'%H Hour in 24 hour format 24
	'%M Minute as an integer 01
	'%S Second as an integer 55
	'%P AM/PM Indicator PM
	'%% Actual Percent sign %%

   Dim intPosItem
   Dim intHourPart
   Dim strHourPart
   Dim strMinutePart
   Dim strSecondPart
   Dim strAMPM

   If not IsDate(strDate) Then
      ConvDate = strDate
      Exit Function
   End If

   intPosItem = Instr(strFormat, "%m")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("m",strDate) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%m")
   Loop

   intPosItem = Instr(strFormat, "%b")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      MonthName(DatePart("m",strDate),True) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%b")
   Loop

   intPosItem = Instr(strFormat, "%B")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      MonthName(DatePart("m",strDate),False) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%B")
   Loop

   intPosItem = Instr(strFormat, "%d")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("d",strDate) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%d")
   Loop

   intPosItem = Instr(strFormat, "%j")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("y",strDate) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%j")
   Loop

   intPosItem = Instr(strFormat, "%y")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      Right(DatePart("yyyy",strDate),2) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%y")
   Loop

   intPosItem = Instr(strFormat, "%Y")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("yyyy",strDate) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%Y")
   Loop

   intPosItem = Instr(strFormat, "%w")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      DatePart("w",strDate,1) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%w")
   Loop

   intPosItem = Instr(strFormat, "%a")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      WeekDayName(DatePart("w",strDate,1),True) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%a")
   Loop

   intPosItem = Instr(strFormat, "%A")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & _
      WeekDayName(DatePart("w",strDate,1),False) & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%A")
   Loop

   intPosItem = Instr(strFormat, "%I")
   Do While intPosItem > 0
      intHourPart = DatePart("h",strDate) mod 12
      if intHourPart = 0 then intHourPart = 12
      strFormat = Left(strFormat, intPosItem-1) & _
      intHourPart & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%I")
   Loop

   intPosItem = Instr(strFormat, "%H")
   Do While intPosItem > 0
      strHourPart = DatePart("h",strDate)
      if strHourPart < 10 Then strHourPart = "0" & strHourPart
      strFormat = Left(strFormat, intPosItem-1) & _
      strHourPart & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%H")
   Loop

   intPosItem = Instr(strFormat, "%M")
   Do While intPosItem > 0
      strMinutePart = DatePart("n",strDate)
      if strMinutePart < 10 then strMinutePart = "0" & strMinutePart
      strFormat = Left(strFormat, intPosItem-1) & _
      strMinutePart & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%M")
   Loop

   intPosItem = Instr(strFormat, "%S")
   Do While intPosItem > 0
      strSecondPart = DatePart("s",strDate)
      if strSecondPart < 10 then strSecondPart = "0" & strSecondPart
      strFormat = Left(strFormat, intPosItem-1) & _
      strSecondPart & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%S")
   Loop

   intPosItem = Instr(strFormat, "%P")
   Do While intPosItem > 0
      if DatePart("h",strDate) >= 12 then
         strAMPM = "PM"
      Else
         strAMPM = "AM"
      End If
      strFormat = Left(strFormat, intPosItem-1) & strAMPM & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%P")
   Loop

   intPosItem = Instr(strFormat, "%%")
   Do While intPosItem > 0
      strFormat = Left(strFormat, intPosItem-1) & "%" & _
      Right(strFormat, Len(strFormat) - (intPosItem + 1))
      intPosItem = Instr(strFormat, "%%")
   Loop

   ConvDate = strFormat

End Function

'====================================================================================================================================

Function EclipseString(strToCheck, inLength)

	If Len(strToCheck) > inLength Then
		EclipseString = Left(EclipseString, inLength) & "..."
	Else
		EclipseString = strToCheck
	End If
End Function

Function WAPValidate(strToCheck)

	WAPValidate = strToCheck

End Function

'====================================================================================================================================

Function WAPEncode(strToCheck)

	If lang = "WML" Then
	' Replace unsupported characters e.g. HTML code characters

	strToCheck = Replace(strToCheck,"&","!amp!")
	strToCheck = Replace(strToCheck,"'","!apos!")
	strToCheck = Replace(strToCheck,">","!gt!")
	strToCheck = Replace(strToCheck,"<","!lt!")
	'strToCheck = Replace(strToCheck," ","&nbsp;")
	strToCheck = Replace(strToCheck,"""","!quot!")
	strToCheck = Replace(strToCheck,"$","$$")

	End If

	WAPEncode = strToCheck

End Function

'====================================================================================================================================

Function WAPDecode(strToCheck)

	If lang = "WML" Then
	' Replace unsupported characters e.g. HTML code characters

	strToCheck = Replace(strToCheck,"!amp!","&")
	strToCheck = Replace(strToCheck,"!apos!","'")
	strToCheck = Replace(strToCheck,"!gt!",">")
	strToCheck = Replace(strToCheck,"!lt!","<")
	'strToCheck = Replace(strToCheck," ","&nbsp;")
	strToCheck = Replace(strToCheck,"!quot!","""")
	strToCheck = Replace(strToCheck,"$$","$")

	End If

	WAPDecode = strToCheck

End Function

'====================================================================================================================================

Function GetWOSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("WOSearchBy",SearchBy)
		Call SetSession("WOSearchValue",SearchValue)
	ElseIf Not GetSession("WOSearchBy") = "" Then
		searchby = GetSession("WOSearchBy")
		searchvalue = GetSession("WOSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND (WO.WOID LIKE '" & searchvalue & "%' OR WO.WOPK LIKE '" & searchvalue & "%') "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=1"">WO# = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND WO.REASON LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=2"">Reason = " & searchvalue & "</a> "
		Case "3"
			sqlwhere = " AND WO.ASSETID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=3"">Asset ID = " & searchvalue & "</a> "
		Case "4"
			sqlwhere = " AND WO.ASSETNAME LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=4"">Asset Name = " & searchvalue & "</a> "
		Case "5"
			sqlwhere = " AND WO.PROCEDUREID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=5"">Procedure ID = " & searchvalue & "</a> "
		Case "6"
			sqlwhere = " AND WO.PROCEDURENAME LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=6"">Procedure Name = " & searchvalue & "</a> "
		Case "7"
			sqlwhere = " AND (WO.TYPE LIKE '" & searchvalue & "%' OR WO.TYPEDESC LIKE '" & searchvalue & "%') "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=7"">Type = " & searchvalue & "</a> "
		Case "8"
			sqlwhere = " AND (WO.PRIORITY LIKE '" & searchvalue & "%' OR WO.PRIORITYDESC LIKE '" & searchvalue & "%') "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=8"">Priority = " & searchvalue & "</a> "
		Case "9"
			sqlwhere = " AND (WO.SUBSTATUS LIKE '" & searchvalue & "%' OR WO.SUBSTATUSDESC LIKE '" & searchvalue & "%') "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=WOSearch&amp;s=" & SessionID & "&amp;wosearchby=9"">Sub-Status = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then

		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a>"
		sqlewhere = sqlewhere & "</div>"
	End If


    ' MRO / Technician WorkCenter WO Filter
    Dim rs
    Set rs = db.RunSQLReturnRS("SELECT AssetPK FROM LaborAssetMgr WITH (NOLOCK) WHERE LaborPK = " & GetSession("UserPK"),"")
    If db.dok Then
	    If Not rs.eof Then
		    sqlwhere = sqlwhere & " AND (WO.ASSETPK IN (SELECT AssetPK FROM MC_GetAssetChildPK('" & RSConcat(rs,0) & "')))"
	    End If
	    Call closeobj(rs)
    End If

End Function

'====================================================================================================================================

Function GetCMSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("CMSearchBy",SearchBy)
		Call SetSession("CMSearchValue",SearchValue)
	ElseIf Not GetSession("CMSearchBy") = "" Then
		searchby = GetSession("CMSearchBy")
		searchvalue = GetSession("CMSearchValue")
	End If

    Dim ft
    ft = Trim(Request("ft"))
    If Not ft = "" Then
        Call SetSession("CMFT",ft)
    Else
        ft = GetSession("CMFT")
    End If
    If ft = "" Then
        ft = "CompanyID"
    End If

	sqlwhere = ""
	sqlewhere = ""

    Select Case UCase(ft)
        Case "COMPANYID"
        Case "VENDORID"
        Case "MANUFACTURERID"
    End Select

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Company.CompanyID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=CMSearch&amp;s=" & SessionID & "&amp;cmsearchby=1"">Company ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Company.CompanyName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=CMSearch&amp;s=" & SessionID & "&amp;cmsearchby=2"">Company Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetFASQLWhere()

	Dim searchbyvar,searchvaluevar,desc
	Select Case UCase(Card)
		Case "FAPRLOOKUP"
			searchbyvar = "FAPRSearchBy"
			searchvaluevar = "FAPRSearchValue"
			desc = "Problem"
			sqlwhere = " AND Failure.Type = 'P' "
		Case "FAFALOOKUP"
			searchbyvar = "FAFASearchBy"
			searchvaluevar = "FAFASearchValue"
			desc = "Failure"
			sqlwhere = " AND Failure.Type = 'F' "
		Case "FASOLOOKUP"
			searchbyvar = "FASOSearchBy"
			searchvaluevar = "FASOSearchValue"
			desc = "Solution"
			sqlwhere = " AND Failure.Type = 'S' "
	End Select

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession(searchbyvar,SearchBy)
		Call SetSession(searchvaluevar,SearchValue)
	ElseIf Not GetSession(searchbyvar) = "" Then
		searchby = GetSession(searchbyvar)
		searchvalue = GetSession(searchvaluevar)
	End If

	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = sqlwhere & " AND Failure.FailureID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & Replace(UCase(card),"LOOKUP","SEARCH") & "&amp;s=" & SessionID & "&amp;" & LCase(searchbyvar) & "=1"">" & desc & " ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = sqlwhere & " AND Failure.FailureName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & Replace(UCase(card),"LOOKUP","SEARCH") & "&amp;s=" & SessionID & "&amp;" & LCase(searchbyvar) & "=2"">" & desc & " Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetASSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("ASSearchBy",SearchBy)
		Call SetSession("ASSearchValue",SearchValue)
	ElseIf Not GetSession("ASSearchBy") = "" Then
		searchby = GetSession("ASSearchBy")
		searchvalue = GetSession("ASSearchValue")
	End If

    Dim ft
    ft = Trim(Request("ft"))
    If Not ft = "" Then
        Call SetSession("ASFT",ft)
    Else
        ft = GetSession("ASFT")
    End If
    If ft = "" Then
        ft = "AssetID"
    End If

	sqlwhere = ""
	sqlewhere = ""

    Select Case UCase(ft)
        Case "ASSETID"
        Case "PARENTID"
    End Select

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Asset.AssetID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=ASSearch&amp;s=" & SessionID & "&amp;assearchby=1"">Asset ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Asset.AssetName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=ASSearch&amp;s=" & SessionID & "&amp;assearchby=2"">Asset Name = " & searchvalue & "</a> "
	End Select

    If Not sqlewhere = "" Then

	    ' MRO / Technician WorkCenter Asset Filter
	    Dim rs
	    Set rs = db.RunSQLReturnRS("SELECT AssetPK FROM LaborAssetMgr WITH (NOLOCK) WHERE LaborPK = " & GetSession("UserPK"),"")
	    If db.dok Then
		    If Not rs.eof Then
			    sqlwhere = sqlwhere & " AND (Asset.ASSETPK IN (SELECT AssetPK FROM MC_GetAssetChildPK('" & RSConcat(rs,0) & "')))"
		    End If
		    Call closeobj(rs)
	    End If

		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"

	End If

End Function

'====================================================================================================================================

Function GetASSQLWhere2()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("ASSearchBy",SearchBy)
		Call SetSession("ASSearchValue",SearchValue)
	ElseIf Not GetSession("ASSearchBy") = "" Then
		searchby = GetSession("ASSearchBy")
		searchvalue = GetSession("ASSearchValue")
	End If

    Dim ft
    ft = Trim(Request("ft"))
    If Not ft = "" Then
        Call SetSession("ASFT",ft)
    Else
        ft = GetSession("ASFT")
    End If
    If ft = "" Then
        ft = "AssetID"
    End If

	sqlwhere = ""
	sqlewhere = ""

    Select Case UCase(ft)
        Case "ASSETID"
        Case "PARENTID"
    End Select

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Asset.AssetID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=ASSearch2&amp;s=" & SessionID & "&amp;assearchby=1"">Asset ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Asset.AssetName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=ASSearch2&amp;s=" & SessionID & "&amp;assearchby=2"">Asset Name = " & searchvalue & "</a> "
	End Select

    If Not sqlewhere = "" Then

	    ' MRO / Technician WorkCenter Asset Filter
	    Dim rs
	    Set rs = db.RunSQLReturnRS("SELECT AssetPK FROM LaborAssetMgr WITH (NOLOCK) WHERE LaborPK = " & GetSession("UserPK"),"")
	    If db.dok Then
		    If Not rs.eof Then
			    sqlwhere = sqlwhere & " AND (Asset.ASSETPK IN (SELECT AssetPK FROM MC_GetAssetChildPK('" & RSConcat(rs,0) & "')))"
		    End If
		    Call closeobj(rs)
	    End If

		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"

	End If

End Function

'====================================================================================================================================

Function GetCLSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("CLSearchBy",SearchBy)
		Call SetSession("CLSearchValue",SearchValue)
	ElseIf Not GetSession("CLSearchBy") = "" Then
		searchby = GetSession("CLSearchBy")
		searchvalue = GetSession("CLSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Classification.ClassificationID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=CLSearch&amp;s=" & SessionID & "&amp;clsearchby=1"">Classification ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Classification.ClassificationName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=CLSearch&amp;s=" & SessionID & "&amp;clsearchby=2"">Classification Name = " & searchvalue & "</a> "
	End Select

    If Not sqlewhere = "" Then

		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"

	End If

End Function

'====================================================================================================================================

Function GetACSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("ACSearchBy",SearchBy)
		Call SetSession("ACSearchValue",SearchValue)
	ElseIf Not GetSession("ACSearchBy") = "" Then
		searchby = GetSession("ACSearchBy")
		searchvalue = GetSession("ACSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Account.AccountID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=ACSearch&amp;s=" & SessionID & "&amp;acsearchby=1"">Account ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Account.AccountName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=ACSearch&amp;s=" & SessionID & "&amp;acsearchby=2"">Account Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetRCSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("RCSearchBy",SearchBy)
		Call SetSession("RCSearchValue",SearchValue)
	ElseIf Not GetSession("RCSearchBy") = "" Then
		searchby = GetSession("RCSearchBy")
		searchvalue = GetSession("RCSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND RepairCenter.RepairCenterID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=RCSearch&amp;s=" & SessionID & "&amp;rcsearchby=1"">Repair Center ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND RepairCenter.RepairCenterName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=RCSearch&amp;s=" & SessionID & "&amp;rcsearchby=2"">Repair Center Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

	If Not RCCheck("") Then
	   sqlwhere = sqlwhere & " AND (RepairCenter.RepairCenterPK IN (" & GetSession("RCDENY") & ") or RepairCenter.RepairCenterPK Is Null) "
	End If

End Function

'====================================================================================================================================

Function GetSHSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("SHSearchBy",SearchBy)
		Call SetSession("SHSearchValue",SearchValue)
	ElseIf Not GetSession("SHSearchBy") = "" Then
		searchby = GetSession("SHSearchBy")
		searchvalue = GetSession("SHSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Shop.ShopID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=SHSearch&amp;s=" & SessionID & "&amp;shsearchby=1"">Shop ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Shop.ShopName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=SHSearch&amp;s=" & SessionID & "&amp;shsearchby=2"">Shop Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetCASQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("CASearchBy",SearchBy)
		Call SetSession("CASearchValue",SearchValue)
	ElseIf Not GetSession("CASearchBy") = "" Then
		searchby = GetSession("CASearchBy")
		searchvalue = GetSession("CASearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Category.CategoryID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=CASearch&amp;s=" & SessionID & "&amp;casearchby=1"">Category ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Category.CategoryName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=CASearch&amp;s=" & SessionID & "&amp;casearchby=2"">Category Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetZNSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("ZNSearchBy",SearchBy)
		Call SetSession("ZNSearchValue",SearchValue)
	ElseIf Not GetSession("ZNSearchBy") = "" Then
		searchby = GetSession("ZNSearchBy")
		searchvalue = GetSession("ZNSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Zone.ZoneID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=ZNSearch&amp;s=" & SessionID & "&amp;znsearchby=1"">Zone ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Zone.ZoneName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=ZNSearch&amp;s=" & SessionID & "&amp;znsearchby=2"">Zone Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetPRSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("PRSearchBy",SearchBy)
		Call SetSession("PRSearchValue",SearchValue)
	ElseIf Not GetSession("PRSearchBy") = "" Then
		searchby = GetSession("PRSearchBy")
		searchvalue = GetSession("PRSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND ProcedureLibrary.ProcedureID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=PRSearch&amp;s=" & SessionID & "&amp;prsearchby=1"">Procedure ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND ProcedureLibrary.ProcedureName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=PRSearch&amp;s=" & SessionID & "&amp;prsearchby=2"">Procedure Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetDPSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("DPSearchBy",SearchBy)
		Call SetSession("DPSearchValue",SearchValue)
	ElseIf Not GetSession("DPSearchBy") = "" Then
		searchby = GetSession("DPSearchBy")
		searchvalue = GetSession("DPSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Department.DepartmentID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=DPSearch&amp;s=" & SessionID & "&amp;dpsearchby=1"">Department ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Department.DepartmentName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=DPSearch&amp;s=" & SessionID & "&amp;dpsearchby=2"">Department Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetTNSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("TNSearchBy",SearchBy)
		Call SetSession("TNSearchValue",SearchValue)
	ElseIf Not GetSession("TNSearchBy") = "" Then
		searchby = GetSession("TNSearchBy")
		searchvalue = GetSession("TNSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Tenant.TenantID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=TNSearch&amp;s=" & SessionID & "&amp;tnsearchby=1"">Customer ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Tenant.TenantName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=TNSearch&amp;s=" & SessionID & "&amp;tnsearchby=2"">Customer Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetPJSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("PJSearchBy",SearchBy)
		Call SetSession("PJSearchValue",SearchValue)
	ElseIf Not GetSession("PJSearchBy") = "" Then
		searchby = GetSession("PJSearchBy")
		searchvalue = GetSession("PJSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Project.ProjectID LIKE '" & searchvalue & "%' "
			sqlewhere = "Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=PJSearch&amp;s=" & SessionID & "&amp;pjsearchby=1"">Project ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Project.ProjectName LIKE '" & searchvalue & "%' "
			sqlewhere = "Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=PJSearch&amp;s=" & SessionID & "&amp;pjsearchby=2"">Project Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetLASQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("LASearchBy",SearchBy)
		Call SetSession("LASearchValue",SearchValue)
	ElseIf Not GetSession("LASearchBy") = "" Then
		searchby = GetSession("LASearchBy")
		searchvalue = GetSession("LASearchValue")
	End If

	'sqlwhere = AddGeneralFilters("LA")
    Dim ft
    ft = Trim(Request("ft"))
    If Not ft = "" Then
        Call SetSession("LAFT",ft)
    Else
        ft = GetSession("LAFT")
    End If
    If ft = "" Then
        ft = "LaborID"
    End If

	sqlwhere = ""
	sqlewhere = ""

    Select Case UCase(ft)
        Case "LABORID"
        Case "OPERATORID"
        Case "CONTACTID"
    End Select

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Labor.LaborID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=LASearch&amp;s=" & SessionID & "&amp;lasearchby=1"">Labor ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Labor.LaborName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=LASearch&amp;s=" & SessionID & "&amp;lasearchby=2"">Labor Name = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetINSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
			searchby = ""
			searchvalue = ""
		End If
		Call SetSession("INSearchBy",SearchBy)
		Call SetSession("INSearchValue",SearchValue)
	ElseIf Not GetSession("INSearchBy") = "" Then
		searchby = GetSession("INSearchBy")
		searchvalue = GetSession("INSearchValue")
	End If

    Dim ft
    ft = Trim(Request("ft"))
    If Not ft = "" Then
        Call SetSession("INFT",ft)
    Else
        ft = GetSession("INFT")
    End If
    If ft = "" Then
        ft = "PartID"
    End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Part.PartID LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=INSearch&amp;s=" & SessionID & "&amp;insearchby=1"">Part ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Part.PartName LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=INSearch&amp;s=" & SessionID & "&amp;insearchby=2"">Part Name = " & searchvalue & "</a> "
		Case "3"
			sqlwhere = " AND Part.PartDescription LIKE '" & searchvalue & "%' "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=INSearch&amp;s=" & SessionID & "&amp;insearchby=3"">Part Desc = " & searchvalue & "</a> "
		Case "4"
			sqlwhere = " AND Part.PartPK IN (Select PartPK FROM PartVendor WHERE VendorPartNumber LIKE '" & searchvalue & "%') "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=INSearch&amp;s=" & SessionID & "&amp;insearchby=3"">Vendor Part # = " & searchvalue & "</a> "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetSRSQLWhere()

	If Not Request("searchby") = "" Then
		searchby = Request("searchby")
		searchvalue = Trim(Request("searchvalue"))
		If searchby = "NONE" or searchvalue = "" Then
		    If Not searchby = "3" Then
			    searchby = ""
			    searchvalue = ""
			End If
		End If
		Call SetSession("SRSearchBy",SearchBy)
		Call SetSession("SRSearchValue",SearchValue)
	ElseIf Not GetSession("SRSearchBy") = "" Then
		searchby = GetSession("SRSearchBy")
		searchvalue = GetSession("SRSearchValue")
	End If

	sqlwhere = ""
	sqlewhere = ""

	Select Case UCase(searchby)
		Case "1"
			sqlwhere = " AND Location.LocationID LIKE '" & searchvalue & "%' " & _
                   	   " AND Location.RepairCenterPK = " & GetSession("RCPK") & " "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=SRSearch&amp;s=" & SessionID & "&amp;srsearchby=1"">Location ID = " & searchvalue & "</a> "
		Case "2"
			sqlwhere = " AND Location.LocationName LIKE '" & searchvalue & "%' " & _
                   	   " AND Location.RepairCenterPK = " & GetSession("RCPK") & " "
			sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: <a class='Font1' style='font-size:12pt;' href=""default.asp?card=SRSearch&amp;s=" & SessionID & "&amp;srsearchby=2"">Location Name = " & searchvalue & "</a> "
		Case "3"
			sqlwhere = " "
    		sqlewhere = "<div class='Font1' style='font-size:12pt;'>Search: All Repair Centers "
		Case Else
			sqlwhere = " AND Location.RepairCenterPK = " & GetSession("RCPK") & " "
	End Select

	If Not sqlewhere = "" Then
		sqlewhere = sqlewhere & "&nbsp;&nbsp;|&nbsp;&nbsp;<a class='Font1' style='font-size:12pt;' href=""default.asp?card=" & CardCurrent & "&amp;s=" & SessionID & "&amp;searchby=NONE"">Show All</a></div>"
	End If

End Function

'====================================================================================================================================

Function GetCardLevel()

	If Not Request("back") = "" Then
		GetCardLevel = (CardFromLevel - CLng(Request("back"))) + CardSkipLevel
	ElseIf UCase(CardCurrent) = UCase(GetSession("ParentCard" & CardFromLevel - 1)) Then
		GetCardLevel = CardFromLevel - 1
	ElseIf UCase(CardCurrent) = UCase(GetSession("ParentCard" & CardFromLevel - 2)) Then
		GetCardLevel = CardFromLevel - 2
	ElseIf UCase(CardCurrent) = UCase(GetSession("ParentCard" & CardFromLevel - 3)) Then
		GetCardLevel = CardFromLevel - 3
	ElseIf UCase(CardFrom) = UCase(CardCurrent) Then
		GetCardLevel = CardFromLevel
	Else
		GetCardLevel = CardFromLevel + CardSkipLevel + 1
	End If

    cardlevelupdownamount = GetCardLevel - cardfromlevel

End Function

'====================================================================================================================================

Sub OutputButtons()

	Select Case UCase(CardCurrent)

	Case "MYWORKORDERS","ALLWORKORDERS","ALLWORKORDERSU","ASHISTORY","CMLOOKUP","RCLOOKUP","SHLOOKUP","ASLOOKUP","CLLOOKUP","ASFORSERIALPARTLOOKUP","LTLOOKUP","ACLOOKUP","CALOOKUP","ZNLOOKUP","PRLOOKUP","DPLOOKUP","TNLOOKUP","PJLOOKUP","LALOOKUP","INLOOKUP","SRLOOKUP","FAPRLOOKUP","FAFALOOKUP","FASOLOOKUP","ASSETS"
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE or IsBlackBerry Then
			OutputBackButton
			OutputSearchButton
		Else
			OutputSearchButton
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="center">
		<%
		'OutputOptionsButton
		'HTMLbuttonsEnd
		'OptionsButtonBegin
		'OutputHomeButton
		OutputSearchButton
		'OptionsButtonEnd
		HTMLbuttonsEnd
		End If

	Case Else
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE or IsBlackBerry Then
			OutputBackButton
			OutputHomeButton
		Else
			OutputHomeButton
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		OutputHomeButton
		HTMLbuttonsEnd
		End If

	End Select

End Sub

'====================================================================================================================================

Sub OptionsButtonBegin()
	%>
	<div name="optionsmenu" id="optionsmenu" style="display:none; width:140px; border:2px outset #ffffff; height:76px; position:absolute; top: 213px; left:136px; background-color:#ffffff; overflow-y:scroll;">
	<%
End Sub

'====================================================================================================================================

Sub OptionsButtonEnd()
	%>
	</div>
	<%
End Sub

'====================================================================================================================================

Sub OutputBackButton()
	Dim backcard
	Dim ii
	ii = 0
	if ii = 1 then
	Select Case UCase(CardCurrent)
		Case "WOSEARCH","LASEARCH","RCSEARCH","SHSEARCH","ASSEARCH","CLSEARCH","ACSEARCH","CASEARCH","ZNSEARCH","PRSEARCH","DPSEARCH","TNSEARCH","PJSEARCH","CMSEARCH","INSEARCH","SRSEARCH","FAPRSEARCH","FAFASEARCH","FASOSEARCH","ASSEARCH2"
			If CardFrom = CardCurrent and Not SearchBy = "" Then
				backcard = CardCurrent
			Else
				backcard = GetSession("ParentCard" & (CardCurrentLevel-1))
			End If
        	If lang = "WML" Then
                Call OutputButtonOrLinkStart("backbutton","Back") %>
				<go href="default.asp" method="get">
					<postfield name="card" value="<% =backcard %>"/><%
					OutputBackButtonStandardPostFields %>
				</go><%
		        Call OutputButtonOrLinkEnd("Back")
			Else %>
			<input style="width:100%;" type="button" name="backbutton" value="Back" onclick="self.location.href='default.asp?card=<% =backcard %><% =OutputBackButtonStandardPostFields() %>';"/><%
			End If
		Case Else
        	If lang = "WML" Then
                Call OutputButtonOrLinkStart("backbutton","Back") %>
				<go href="default.asp" method="get">
					<postfield name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/><% OutputBackButtonStandardPostFields %>
				</go><%
		        Call OutputButtonOrLinkEnd("Back")
			Else
			Select Case UCase(CardCurrent)
			Case "LALOOKUP","CMLOOKUP","RCLOOKUP","SHLOOKUP","ASLOOKUP","CLLOOKUP","ASFORSERIALPARTLOOKUP","LTLOOKUP","ACLOOKUP","CALOOKUP","CALENDARLOOKUP","ZNLOOKUP","PRLOOKUP","DPLOOKUP","TNLOOKUP","PJLOOKUP","INLOOKUP","SRLOOKUP","FAPRLOOKUP","FAFALOOKUP","FASOLOOKUP" %>
			<input style="width:100%;" type="button" name="backbutton" value="Back" onclick="closeiframe();self.location.href='default.asp?card=<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %><% =OutputBackButtonStandardPostFields() %>';"/><%
			Case Else %>
			<input style="width:100%;" type="button" name="backbutton" value="Back" onclick="self.location.href='default.asp?card=<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %><% =OutputBackButtonStandardPostFields() %>';"/><%
			End Select
			End If
	End Select
	end if
End Sub

'====================================================================================================================================

Function OutputBackButtonStandardPostFields()
	If lang = "WML" Then %>
	<% If Not SessionID = "" Then %>
	<postfield name="s" value="<% =SessionID %>"/><%
	End If %>
	<% If Not GetSession("PagePos" & (CardCurrentLevel-1)) = "" Then %>
	<postfield name="pagepos" value="<% =GetSession("PagePos" & (CardCurrentLevel-1)) %>"/><%
	End If %>
	<% If Not WOPK = "" Then %>
	<postfield name="wopk" value="<% =WOPK %>"/><%
	End If %>
	<% If Not AssetPK = "" Then %>
	<postfield name="assetpk" value="<% =AssetPK %>"/><%
	End If %>
	<postfield name="back" value="1"/><%
	Else
		OutputBackButtonStandardPostFields = ""
		OutputBackButtonStandardPostFields = OutputBackButtonStandardPostFields & _
			"&s=" & SessionID
		OutputBackButtonStandardPostFields = OutputBackButtonStandardPostFields & _
			"&PagePos=" & GetSession("PagePos" & (CardCurrentLevel-1))
		OutputBackButtonStandardPostFields = OutputBackButtonStandardPostFields & _
			"&WOPK=" & WOPK
		OutputBackButtonStandardPostFields = OutputBackButtonStandardPostFields & _
			"&AssetPK=" & AssetPK
		OutputBackButtonStandardPostFields = OutputBackButtonStandardPostFields & _
			"&Back=1"
	End If
End Function

'====================================================================================================================================

Sub OutputHomeButton()
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("homebutton","Home") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="<% =GetSession("ParentCard1") %>"/>
			<postfield name="s" value="<% =SessionID %>"/>
		</go><%
	Call OutputButtonOrLinkEnd("Home")
	Else
	%>
	<input style="width:100%;" type="button" name="homebutton" value="Home" onclick="self.location.href='default.asp?card=<% =GetSession("ParentCard1") %>&s=<% =SessionID %>';"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub OutputOptionsButton()
	%>
	<input style="width:100%;" type="button" name="optionsbutton" value="Options" onclick="if (self.optionsmenu.style.display == '') {self.optionsmenu.style.display='none'} else {self.optionsmenu.style.display = ''};"/>
	<%
End Sub

'====================================================================================================================================

Function GetBackButton(inURL)
	Dim backcard
	Select Case UCase(CardCurrent)
		Case "WOSEARCH","LASEARCH","RCSEARCH","SHSEARCH","ASSEARCH","CLSEARCH","ACSEARCH","CASEARCH","ZNSEARCH","PRSEARCH","DPSEARCH","TNSEARCH","PJSEARCH","CMSEARCH","INSEARCH","SRSEARCH","FAPRSEARCH","FAFASEARCH","FASOSEARCH","ASSEARCH2"
			If CardFrom = CardCurrent and Not SearchBy = "" Then
				backcard = CardCurrent
			Else
				backcard = GetSession("ParentCard" & (CardCurrentLevel-1))
			End If
			GetBackButton = "<div onclick=""self.location.href='default.asp?card=" & backcard & OutputBackButtonStandardPostFields & "';"" style='float:left; padding-right:10px; padding-top:6px;'><img src='images/icons/48/Navigation 1 Left Black.png' alt='Back' title='Back' style='border:none; cursor:pointer; width:42px; height:42px;' /></div>"
		Case Else

			Select Case UCase(CardCurrent)
			Case "LALOOKUP","CMLOOKUP","RCLOOKUP","SHLOOKUP","ASLOOKUP","CLLOOKUP","ASFORSERIALPARTLOOKUP","LTLOOKUP","ACLOOKUP","CALOOKUP","CALENDARLOOKUP","ZNLOOKUP","PRLOOKUP","DPLOOKUP","TNLOOKUP","PJLOOKUP","INLOOKUP","SRLOOKUP","FAPRLOOKUP","FAFALOOKUP","FASOLOOKUP"
				GetBackButton = "<div onclick=""self.location.href='default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & OutputBackButtonStandardPostFields & "';"" style='float:left; padding-right:10px; padding-top:6px;'><img src='images/icons/48/Navigation 1 Left Black.png' alt='Back' title='Back' style='border:none; cursor:pointer; width:42px; height:42px;' /></div>"
			Case Else
				GetBackButton = "<div onclick=""self.location.href='default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & OutputBackButtonStandardPostFields & "';"" style='float:left; padding-right:10px; padding-top:6px;'><img src='images/icons/48/Navigation 1 Left Black.png' alt='Back' title='Back' style='border:none; cursor:pointer; width:42px; height:42px;' /></div>"
			End Select
	End Select
End Function

Function GetAddButton(inURL)
	GetAddButton = "<div onclick=""" & inURL & """ style='float:left; padding-right:10px; padding-top:6px;'><img src='images/icons/48/add.png' alt='Add' title='Add' style='border:none; cursor:pointer; width:42px; height:42px;' /></div>"
End Function

Function GetSearchButton()

	Dim SearchCard, ShowSearch
	ShowSearch = True

	Select Case UCase(CardCurrent)
	Case "CMLOOKUP"
		SearchCard = "CMSearch"
	Case "RCLOOKUP"
		SearchCard = "RCSearch"
	Case "SHLOOKUP"
		SearchCard = "SHSearch"
	Case "ASLOOKUP","ASFORSERIALPARTLOOKUP"
		SearchCard = "ASSearch"
	Case "ASSETS"
	    SearchCard = "ASSearch2"
	Case "CLLOOKUP"
	    SearchCard = "CLSearch"
	Case "ACLOOKUP"
		SearchCard = "ACSearch"
	Case "LTLOOKUP"
		SearchCard = ""
		ShowSearch = False
	Case "CALOOKUP"
		SearchCard = "CASearch"
	Case "ZNLOOKUP"
		SearchCard = "ZNSearch"
	Case "PRLOOKUP"
		SearchCard = "PRSearch"
	Case "DPLOOKUP"
		SearchCard = "DPSearch"
	Case "TNLOOKUP"
		SearchCard = "TNSearch"
	Case "PJLOOKUP"
		SearchCard = "PJSearch"
	Case "LALOOKUP"
		SearchCard = "LASearch"
	Case "INLOOKUP"
		SearchCard = "INSearch"
	Case "SRLOOKUP"
		SearchCard = "SRSearch"
	Case "FAPRLOOKUP"
		SearchCard = "FAPRSearch"
	Case "FAFALOOKUP"
		SearchCard = "FAFASearch"
	Case "FASOLOOKUP"
		SearchCard = "FASOSearch"
	Case Else
		SearchCard = "WOSearch"
	End Select

	If SearchCard <> "" And Cbool(ShowSearch) Then
		GetSearchButton = "<div onclick=""self.location.href='default.asp?card=" & SearchCard & "&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:6px;'><img src='images/icons/48/Find.png' alt='Search' title='Search' style='border:none; cursor:pointer; width:42px; height:42px;' /></div>"
	End If
End Function

Sub OutputSearchButton()

	Dim SearchCard

	Select Case UCase(CardCurrent)
	Case "CMLOOKUP"
		SearchCard = "CMSearch"
	Case "RCLOOKUP"
		SearchCard = "RCSearch"
	Case "SHLOOKUP"
		SearchCard = "SHSearch"
	Case "ASLOOKUP","ASFORSERIALPARTLOOKUP"
		SearchCard = "ASSearch"
	Case "ASSETS"
	    SearchCard = "ASSearch2"
	Case "CLLOOKUP"
	    SearchCard = "CLSearch"
	Case "ACLOOKUP"
		SearchCard = "ACSearch"
	Case "LTLOOKUP"
		SearchCard = "LTSearch"
	Case "CALOOKUP"
		SearchCard = "CASearch"
	Case "ZNLOOKUP"
		SearchCard = "ZNSearch"
	Case "PRLOOKUP"
		SearchCard = "PRSearch"
	Case "DPLOOKUP"
		SearchCard = "DPSearch"
	Case "TNLOOKUP"
		SearchCard = "TNSearch"
	Case "PJLOOKUP"
		SearchCard = "PJSearch"
	Case "LALOOKUP"
		SearchCard = "LASearch"
	Case "INLOOKUP"
		SearchCard = "INSearch"
	Case "SRLOOKUP"
		SearchCard = "SRSearch"
	Case "FAPRLOOKUP"
		SearchCard = "FAPRSearch"
	Case "FAFALOOKUP"
		SearchCard = "FAFASearch"
	Case "FASOLOOKUP"
		SearchCard = "FASOSearch"
	Case Else
		SearchCard = "WOSearch"
	End Select

End Sub

'====================================================================================================================================

Sub HTMLButtonsBegin
	%>
	<div style="padding-top:10px;">
	<%
End Sub

'====================================================================================================================================

Sub HTMLButtonsEnd
	%>
		</div><div style='clear:both;'></div>
	<%
End Sub

'====================================================================================================================================

Sub GetWOPK()
	Dim t,sql,sqlwhere
	t = 1
	sql = ""
	sqlwhere = ""
	Do While t <= 5
		sqlwhere = GetSession("sqlwhere" & CardCurrentLevel - t)
		If Not sqlwhere = "" Then
			Exit Do
		End If
		t = t + 1
	Loop
	If not sqlwhere = "" Then
	    If Not InStr(sqlwhere,"UNION ") > 0 Then
		    sql = "SELECT WO.WOPK " & _
			    sqlwhere
	    Else
	        sql = sqlwhere
	    End If
	End If
	If Not Request("nr") = "" and Not sql = "" Then
		wopk = Request("nr")
		Set rs = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		Do While Not rs.eof
			If CLng(wopk) = CLng(rs("wopk")) Then
				Exit Do
			End If
			rs.MoveNext
		Loop
		If Not rs.Eof Then
			rs.MoveNext
			If Not rs.eof Then
				wopk = rs("WOPK")
			Else
				IsEOF = True
			End If
		Else
			IsEOF = True
		End If
	ElseIf Not Request("pr") = "" and Not sql = "" Then
		wopk = Request("pr")
		Set rs = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		Do While Not rs.eof
			If CLng(wopk) = CLng(rs("wopk")) Then
				Exit Do
			End If
			rs.MoveNext
		Loop
		If Not rs.Bof Then
			rs.MovePrevious
			If Not rs.bof Then
				wopk = rs("WOPK")
			Else
				IsBOF = True
			End If
		Else
			IsBOF = True
		End If
	Else
		wopk = Request("WOPK")
		If wopk = "" Then
			wopk = GetSession("WOPK" & CardCurrentLevel)
		Else
			Call SetSession("WOPK" & CardCurrentLevel,WOPK)
		End If
	End If
End Sub

'====================================================================================================================================

Sub GetAssetPK()

	assetpk = Request("AssetPK")
	If assetpk = "" Then
		assetpk = GetSession("ASSETPK" & CardCurrentLevel)
	Else
		Call SetSession("ASSETPK" & CardCurrentLevel,assetpk)
	End If

End Sub

'====================================================================================================================================

Sub GetClassificationPK()

	classificationpk = Request("ClassificationPK")
	If classificationpk = "" Then
		classificationpk = GetSession("ClassificationPK" & CardCurrentLevel)
	Else
		Call SetSession("ClassificationPK" & CardCurrentLevel,classificationpk)
	End If

End Sub

'====================================================================================================================================

Sub GetPagePos()

	If Request("PagePos") = "" Then
		If (Not GetSession("PagePos" & CardCurrentLevel) = "") and (CardCurrentLevel < CardFromLevel) Then
			PagePos = CLng(GetSession("PagePos" & CardCurrentLevel))
		Else
			PagePos = 0
		End If
	Else
		PagePos = CLng(Request("PagePos"))
	End If

	TabIndex = 0
	ChoiceIndex = 1

End Sub

'====================================================================================================================================

Sub HTMLLoadingPage
            %>
<html>
<head>
<title>Maintenance Connection Mobile</title>
<meta http-equiv="imagetoolbar" content="no">
<meta name="MSSmartTagsPreventParsing" content="TRUE">
<script language="javascript">
function launch()
{
	var w = 402;
	var h = 610;

	var aW = self.screen.availWidth-10;
	var aH = self.screen.availHeight-30;

	if (aW >= ((1024*2)-50) || aH >= ((768*2)-50))
	{
	 	// Using dual montior display
	 	if (aW >= ((1024*2)-50))
	 	{
	 		aW = aW / 2;
	 	}
	 	else
	 	{
	 		aH = aH / 2;
	 	}
	}

	var w2 = (aW - w) / 2;
	var h2 = (aH - h) / 2;

	h2 = h2 - 20;

	if (h2 < 0)
	{
		h2 = 0;
	}
	if (w2 < 0)
	{
		w2 = 0;
	}

	self.mcmobilewindow = self.open('default.asp?card=login','mcmobile', "width="+w+",height="+h+",left="+w2+",top="+h2+",scrollbars=no,resizable=no,status=no,toolbar=no,menubar=no,location=no,directories=no");
}
</script>
<body bgcolor="#6475D7" oncontextmenu="return false;" style="scrollbar-base-color: #EAEAEA;" link="blue" vlink="blue" alink="blue">
<% =UCase(Request.ServerVariables("HTTP_USER_AGENT")) %>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" height="100%" align="center">
<div style="width:170px; padding-top:10px; padding-right:10px; padding-left:10px; padding-bottom:15px; border:4px outset #FFFFFF; background-color:#ffffff;">
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td>
<p><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/mcmobile_login.gif" alt="MC Mobile"/></p>
<p><font face="arial" style="font-size:9pt;">Clicking the Launch<br/>button below will<br/>open a new window.<br/><br/>If you use Pop-up<br/>blockers, they must<br/>be disabled.</font></p>
<p><input type="button" name="launch" value="      Launch      " onclick="launch();"></a></p>
</td>
</tr>
</table>
</div>
<br><br><br>&nbsp;
</td>
</tr>
</table>
</body>
</html>
<%
Response.End
End Sub

'====================================================================================================================================

Function stripyear(d)
	stripyear = Replace(d,"/"&Year(d),"")
End Function

'====================================================================================================================================

Sub FTF(fieldtype)
	Select Case UCase(fieldtype)
	Case "C"
		Fields = Fields & GetFieldTitle() & "<div style='clear:both;'></div>"
	Case "B"
		FieldsB = FieldsB & GetFieldTitle() & "<div style='clear:both;'></div>"
	End Select
End Sub

'====================================================================================================================================

Sub FTE(fieldtype)
	Select Case UCase(fieldtype)
	Case "C"
		If IsPocketIE Then
			Fields = Fields & " " & GetFieldTitle() & "<div style='clear:both;'></div>"
		ElseIf IsBlackBerry Then
			Fields = Fields & "<br/>"
		Else
			Fields = Fields & "<br/>"
		End If
		If Lang="WML" Then
			PostFields = PostFields & "<postfield name=""" & FN & """ value=""$(" & FN & r & ")""/>"
		End If
	Case "B"
		If IsPocketIE Then
			FieldsB = FieldsB & " " & GetFieldTitle() & "<div style='clear:both;'></div>"
		ElseIf IsBlackBerry Then
			FieldsB = FieldsB & "<br/>"
		Else
			FieldsB = FieldsB & "<br/>"
		End If
		If Lang="WML" Then
			PostFields = PostFields & "<postfield name=""" & FN & """ value=""$(" & FN & r & ")""/>"
		End If
	End Select
End Sub

'====================================================================================================================================

Function GetFieldTitle()

	Dim targetframe
	targetframe = ""

	If lang="HTML" Then
		targetframe = " target=""mciframe"""
	End If

	GetFieldTitle = FT
	Dim lookupIcon

	lookupIcon = "<div style='float:left; padding-right:5px;'><img src='images/search_page.png' alt='Lookup' title='Lookup' style='border:none;' /></div>"
	If IsOpen Then
		'response.write(Trim(UCase(FN)) & "<br />")
		Select Case Trim(UCase(FN))
		    Case "TARGETDATE","WORKDATE","MISCCOSTDATE","ASSIGNEDDATE","DATE","PURCHASEDDATE","INSTALLDATE","REPLACEDATE","DISPOSALDATE","WARRANTYEXPIRE","VALUEDATE"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Calendar', 400, 400, 'default.asp', '&card=CalendarLookup&amp;ft="&FN&"&amp;s=" & SessionID & "');"">" & FT & "</div>"
			Case "REPAIRCENTERID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Repair Center', 600, 400, 'default.asp', '&card=RClookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "SHOPID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Shops', 600, 400, 'default.asp', '&card=SHlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "LABORID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Labor/Contacts', 600, 400, 'default.asp', '&card=LAlookup&ft=LaborID&s=" & SessionID & "');"">" & FT & "</div>"
			Case "OPERATORID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Operators', 600, 400, 'default.asp', '&card=LAlookup&ft=OperatorID&s=" & SessionID & "');"">" & FT & "</div>"
			Case "CONTACTID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Contacts', 600, 400, 'default.asp', '&card=LAlookup&ft=ContactID&s=" & SessionID & "');"">" & FT & "</div>"
			Case "CATEGORYID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Categories', 600, 400, 'default.asp', '&card=CAlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "ZONEID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Zones', 600, 400, 'default.asp', '&card=ZNlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "PROCEDUREID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Procedures', 600, 400, 'default.asp', '&card=PRlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "DEPARTMENTID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Departments', 600, 400, 'default.asp', '&card=DPlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "TENANTID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Tenants', 600, 400, 'default.asp', '&card=TNlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "PROJECTID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Projects', 600, 400, 'default.asp', '&card=PJlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "ASSETID","PARENTID"
			    If Trim(UCase(FN)) = "ASSETID" Then
			        If Trim(UCase(CardCurrent)) = "ASDETAILS" Then
			            If Not AssetPK = "-1" Then
							GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Assets', 600, 580, 'default.asp', '&card=ASlookup&ft=AssetID&s=" & SessionID & "');"">" & FT & "</div>"
						Else
							GetFieldTitle = "<div style='float:left;' class='FieldFont'>" & FT & "</div>"
				        End If
				    Else
						GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Assets', 600, 580, 'default.asp', '&card=ASlookup&ft=AssetID&s=" & SessionID & "');"">" & FT & "</div>"
				    End If
				Else
					GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Assets', 600, 580, 'default.asp', '&card=ASlookup&ft=ParentID&s=" & SessionID & "');"">" & FT & "</div>"
				End If
			Case "CLASSIFICATIONID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Classifications', 600, 580, 'default.asp', '&card=CLlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "SERIALREPLACETOLOCATIONID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Locations', 600, 400, 'default.asp', '&card=ASForSerialPartLookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "ACCOUNTID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Accounts', 600, 400, 'default.asp', '&card=AClookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "PARTID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Parts', 600, 400, 'default.asp', '&card=INlookup&ft=PartID&s=" & SessionID & "');"">" & FT & "</div>"
			Case "ROTATINGPARTID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Rotating Parts', 600, 400, 'default.asp', '&card=INlookup&ft=RotatingPartID&s=" & SessionID & "');"">" & FT & "</div>"
			Case "LOCATIONID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Locations', 600, 400, 'default.asp', '&card=SRlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "COMPANYID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Companies', 600, 400, 'default.asp', '&card=CMlookup&ft=CompanyID&s=" & SessionID & "');"">" & FT & "</div>"
			Case "VENDORID"
				GetFieldTitle = lookupIcon & "<div class='FieldFont' onclick=""OpenWindow('Vendors', 600, 400, 'default.asp', '&card=CMlookup&ft=VendorID&s=" & SessionID & "');"">" & FT & "</div>"
			Case "MANUFACTURERID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Manufactures', 600, 400, 'default.asp', '&card=CMlookup&ft=ManufacturerID&s=" & SessionID & "');"">" & FT & "</div>"
			Case "PROBLEMID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Problems', 600, 400, 'default.asp', '&card=FAPRlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "FAILUREID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Failures', 600, 400, 'default.asp', '&card=FAFAlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "SOLUTIONID"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Solutions', 600, 400, 'default.asp', '&card=FASOlookup&s=" & SessionID & "');"">" & FT & "</div>"
			Case "LABORREPORT"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Labor', 600, 400, 'default.asp', '&card=LTlookup&amp;lt=actions&amp;lf=LaborReport&amp;lr=codedesc&s=" & SessionID & "');"">" & FT & "</div>"
			Case "PRIORITY"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Priorites', 600, 400, 'default.asp', '&card=LTlookup&amp;lt=wopriority&amp;lf=Priority&amp;lr=codename&s=" & SessionID & "');"">" & FT & "</div>"
		    Case "METER1UNITS"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Meter 1', 600, 400, 'default.asp', '&card=LTlookup&amp;lt=meterunits&amp;lf=Meter1Units&amp;lr=codename&s=" & SessionID & "');"">" & FT & "</div>"
		    Case "METER2UNITS"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Meter 2', 600, 400, 'default.asp', '&card=LTlookup&amp;lt=meterunits&amp;lf=Meter2Units&amp;lr=codename&s=" & SessionID & "');"">" & FT & "</div>"
			Case "TYPE"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('WO Type', 600, 400, 'default.asp', '&card=LTlookup&amp;lt=wotype&amp;lf=Type&amp;lr=codename&s=" & SessionID & "');"">" & FT & "</div>"
			Case "STATUS"
				GetFieldTitle = lookupIcon & "<div style='float:left;' class='FieldFont' onclick=""OpenWindow('Status', 600, 400, 'default.asp', '&card=LTlookup&amp;lt=assetstatus&amp;lf=Status&amp;lr=codename&s=" & SessionID & "');"">" & FT & "</div>"
			Case Else
				GetFieldTitle = "<div style='float:left;' class='FieldFont2'>" & FT & "</div>"
		End Select
	Else
		GetFieldTitle = "<div style='float:left;' class='FieldFont2'>" & FT & "</div>"
	End If

End Function

'====================================================================================================================================

Function ProcessPostUpdateSQL(db)

	If ErrorHandler Then
		On Error Resume Next
	End If

	If Not db.dok Then
		Exit Function
	End If

	If Not PostUpdateSQL = "" Then

		Call db.RunSQL(PostUpdateSQL,"")
		PostUpdateSQL = ""

		If Not db.dok Then
			Call OutputWAPError(db.derror)
		End If

	End If

End Function

'====================================================================================================================================

Function OutputFieldValue(rs,ft,pk,v,dvalue)
    On Error Resume Next
	If pk = "-1" Then
		Select Case ft
			Case "C"
				If Request(v) = "" Then
					OutputFieldValue = dvalue
				Else
					OutputFieldValue = WAPDecode(Request(v))
				End If
				Call AddToSetVars(v,"<setvar name=""" & v & """ value=""" & WAPValidate(OutputFieldValue) & """/>")
			Case "B"
				If Request(v) = "" Then
					OutputFieldValue = dvalue
				Else
					If UCase(Request(v)) = "Y" Then
						OutputFieldValue = True
					Else
						OutputFieldValue = False
					End If
				End If
				If OutputFieldValue Then
					Call AddToSetVars(v,"<setvar name=""" & v & """ value=""Y""/>")
				Else
					Call AddToSetVars(v,"<setvar name=""" & v & """ value=""N""/>")
				End If
			Case Else
				If Request(v) = "" Then
					OutputFieldValue = dvalue
				Else
					OutputFieldValue = WAPDecode(Request(v))
				End If
				Call AddToSetVars(v,"<setvar name=""" & v & """ value=""" & WAPValidate(OutputFieldValue) & """/>")
			End Select
	Else
		Select Case ft
			Case "C"
				If Request(v) = "" Then
					OutputFieldValue = NullCheck(rs(v))
					If err > 0 Then
					    Response.Write "Could not find field: " & v
					    Response.End
					End If
				Else
					OutputFieldValue = WAPDecode(Request(v))
				End If
				Call AddToSetVars(v,"<setvar name=""" & v & """ value=""" & WAPValidate(OutputFieldValue) & """/>")
			Case "B"
				If Request(v) = "" Then
					If UCase(v) = "DELETE" Then
						OutputFieldValue = dvalue
					Else
						OutputFieldValue = BitNullCheck(rs(v))
					    If err > 0 Then
					        Response.Write "Could not find field: " & v
					        Response.End
					    End If
					End If
				Else
					If UCase(Request(v)) = "Y" Then
						OutputFieldValue = True
					Else
						OutputFieldValue = False
					End If
				End If
				If OutputFieldValue Then
					Call AddToSetVars(v,"<setvar name=""" & v & """ value=""Y""/>")
				Else
					Call AddToSetVars(v,"<setvar name=""" & v & """ value=""N""/>")
				End If
			Case Else
				If Request(v) = "" Then
					OutputFieldValue = NullCheck(rs(v))
					If err > 0 Then
					    Response.Write "Could not find field: " & v
					    Response.End
					End If
				Else
					OutputFieldValue = WAPDecode(Request(v))
				End If
				Call AddToSetVars(v,"<setvar name=""" & v & """ value=""" & WAPValidate(OutputFieldValue) & """/>")
		End Select
	End If
End Function

'====================================================================================================================================

Sub AddToSetVars(v,s)
	If Not InStr(SetVars,"name=""" & v & """") > 0 Then
		If Request("back") = "" Then
			SetVars = SetVars & s
			SetVarsDebug = SetVarsDebug & Replace(Replace(s,"/>",""),"<","") & "<br/>"
		Else
			If Not Request(v) = "" Then
				SetVars = SetVars & s
				SetVarsDebug = SetVarsDebug & Replace(Replace(s,"/>",""),"<","") & "<br/>"
			End If
		End If
	End If
End Sub

'====================================================================================================================================

Sub OutputSetVars()
If (Not Request("back") = "") and (IsBlackBerry) Then
	Exit Sub
End If
If lang = "WML" Then
If Not SetVars = "" Then %>
<onevent type="onenterforward">
<refresh>
<% =SetVars %>
</refresh>
</onevent>
<%
End If
End If
End Sub

'====================================================================================================================================

Sub BuildFields(tFN,tFT,FieldType,FieldSize,FieldMask,FieldEmptyOK,FieldDefault,rs,PK)
	Dim tabindexoutput
	FN = tFN
	FT = tFT
	TabIndex = TabIndex + 1
	If lang="HTML" Then
		tabindexoutput = "tabindex=""" & CStr(TabIndex) & """ "
	Else
		tabindexoutput = "tabindex=""" & CStr(TabIndex) & """ "
	End If
	Call FTF(FieldType)
	Select Case UCase(FieldType)
		Case "C"
			Fields = Fields & "<div><input class='Textbox_Field' " & tabindexoutput & "size=""" & FieldSize & """ value=""" & OutputFieldValue(rs,FieldType,PK,FN,FieldDefault) & """ emptyok=""" & FieldEmptyOK & """ type=""text"" name=""" & FN & r & """ id=""" & FN & r & """ format=""" & FieldMask & """/></div>" & vbcrlf
		Case "B"
			FieldsB = FieldsB & "<div><select class='Textbox_Field' " & tabindexoutput & "iname=""" & FN & """ class=""test"" name=""" & FN & r & """ value="""
			If OutputFieldValue(rs,FieldType,PK,FN,FieldDefault) Then
			FieldsB = FieldsB & "Y"
			Else
			FieldsB = FieldsB & "N"
			End If
			FieldsB = FieldsB & """>" & vbCrLf
			FieldsB = FieldsB & "<option value=""N"""
			If Not OutputFieldValue(rs,FieldType,PK,FN,FieldDefault) and Lang = "HTML" Then
			FieldsB = FieldsB & " selected"
			End If
			FieldsB = FieldsB & "><onevent type=""onpick""><noop/></onevent>N</option>" & vbCrLf
			FieldsB = FieldsB & "<option value=""Y"""
			If OutputFieldValue(rs,FieldType,PK,FN,FieldDefault) and Lang = "HTML" Then
			FieldsB = FieldsB & " selected"
			End If
			FieldsB = FieldsB & "><onevent type=""onpick""><noop/></onevent>Y</option>" & vbCrLf
			FieldsB = FieldsB & "</select></div>" & vbCrLf
	End Select
	Call FTE(FieldType)
End Sub

'====================================================================================================================================

Sub OutputFields()
	If IsBlackBerry Then
		' Well - the BlackBerry will "auto-invoke" the first button when selecting a value in a
		' select input field if it is the last field on a form. So - instead of moving the
		' select options to the front - we simply make sure the submit button is the first button
		' on forms.
		'Response.Write "<p mode=""wrap"">" & FieldsB & Fields & "</p>"
		Response.Write "<p mode=""wrap"">" & Fields & FieldsB & "</p>"
	Else
		Response.Write "<p mode=""wrap"">" & Fields & FieldsB & "</p>"
	End If
End Sub

'====================================================================================================================================

Function AddGeneralFilters(p)
	AddGeneralFilters = ""
	If Not GetSession("SHPK") = "" Then
		AddGeneralFilters = " AND " & p & ".ShopPK = " & GetSession("SHPK") & " "
	End If
End Function

'====================================================================================================================================

Function SaveField(ByRef db, fn, ft, tn, tnf, t, c, required)
    On Error Resume Next
    Dim rs2
    If ft = "" Then
        ft = fn
    End If
    If tnf = "" Then
        tnf = fn
    End If
    Select Case t
    Case "LM" ' - Lookup Module
        SaveField = True
        If Not NullCheck(Request(fn)) = "" Then
	        sql = "SELECT * FROM " & tn & " WITH (NOLOCK) WHERE " & tnf & " = '" & NullCheck(Request(fn)) & "' "
	        Set rs2 = db.RunSQLReturnRS(sql,"")
	        Call CheckDB(db)
	        If rs2.eof Then
		        HeaderMSG = "The " & ft & " was not found."
		        SaveField = False
		        If Not c = "" Then
		            Execute("Call " & c)
		        End If
	        Else
		        rs(Replace(fn,"ID","PK")) = rs2(Replace(tnf,"ID","PK"))
		        rs(fn) = rs2(tnf)
		        rs(Replace(fn,"ID","Name")) = rs2(Replace(tnf,"ID","Name"))
		        On Error Goto 0
		        Select Case Trim(UCase(fn))
		        Case "CLASSIFICATIONID"
		            rs("Type") = rs2("Type")
		            rs("TypeDesc") = rs2("TypeDesc")
		            If Trim(rs2("Type")) = "A" or _
		               Trim(rs2("Type")) = "AL" Then
			            rs("IsLocation") = False
		            Else
			            rs("IsLocation") = True
		            End If
		            If Trim(rs2("Type")) = "AL" Then
			            rs("IsLinear") = True
		            Else
			            rs("IsLinear") = False
		            End If
		            rs("Model") = rs2("Model")
		            rs("ModelNumber") = rs2("ModelNumber")
		            rs("ModelNumberMFG") = rs2("ModelNumberMFG")
		            rs("ModelLine") = rs2("ModelLine")
		            rs("ModelLineDesc") = rs2("ModelLineDesc")
		            rs("ModelSeries") = rs2("ModelSeries")
		            rs("ModelSeriesDesc") = rs2("ModelSeriesDesc")
		            rs("SystemPlatform") = rs2("SystemPlatform")
		            rs("SystemPlatformDesc") = rs2("SystemPlatformDesc")
		            rs("ManufacturerPK") = rs2("ManufacturerPK")
		            rs("ManufacturerID") = rs2("ManufacturerID")
		            rs("ManufacturerName") = rs2("ManufacturerName")
		            rs("RiskLevel") = rs2("RiskLevel")
                    If NullCheck(rs2("PMCycleStartBy")) = "" Then
		                rs("PMCycleStartBy") = rs2("PMCycleStartBy")
		                rs("PMCycleStartByDesc") = rs2("PMCycleStartByDesc")
                    Else
		                rs("PMCycleStartBy") = "PM"
		                rs("PMCycleStartByDesc") = "PM Settings"
                    End If
		            rs("IsMeter") = rs2("IsMeter")
		            rs("Meter1TrackHistory") = rs2("Meter1TrackHistory")
		            rs("Meter2TrackHistory") = rs2("Meter2TrackHistory")
		            rs("Icon") = rs2("Icon")
		            rs("IsLocation") = rs2("IsLocation")
		            rs("Meter1Units") = rs2("Meter1Units")
		            rs("Meter1UnitsDesc") = rs2("Meter1UnitsDesc")
		            rs("Meter2Units") = rs2("Meter2Units")
		            rs("Meter2UnitsDesc") = rs2("Meter2UnitsDesc")
		        End Select
		        On Error Resume Next
	        End If
	        Call CloseObj(rs2)
        Else
            If required Then
		        HeaderMSG = "The " & ft & " is required."
		        SaveField = False
		        If Not c = "" Then
		            Execute("Call " & c)
		        End If
            Else
                If Not NewRecord Then
		            rs(Replace(fn,"ID","PK")) = Null
		            rs(fn) = Null
		            rs(Replace(fn,"ID","Name")) = Null
		        End If
		    End If
        End If
    Case "LT"
        SaveField = True
        If Not NullCheck(Request(fn)) = "" Then
    		sql = "SELECT * FROM LookupTableValues WITH (NOLOCK) WHERE LookupTable = '" & tn & "' AND CodeName = '" & NullCheck(Request(fn)) & "' "
	        Set rs2 = db.RunSQLReturnRS(sql,"")
	        Call CheckDB(db)
	        If rs2.eof Then
		        HeaderMSG = "The " & ft & " was not found."
		        SaveField = False
		        If Not c = "" Then
		            Execute("Call " & c)
		        End If
	        Else
		        rs(fn) = rs2("CodeName")
		        rs(fn&"DESC") = rs2("CodeDesc")
	        End If
	        Call CloseObj(rs2)
        Else
            If required Then
		        HeaderMSG = "The " & ft & " is required."
		        SaveField = False
		        If Not c = "" Then
		            Execute("Call " & c)
		        End If
            Else
                If Not NewRecord Then
    	            rs(fn) = Null
	                rs(fn&"Desc") = Null
	            End If
	        End If
        End If
    Case "B"
		If UCase(NullCheck(Request(fn))) = "Y" or UCase(NullCheck(Request(fn))) = "2" Then
			rs(fn) = True
		Else
			rs(fn) = False
		End If
    Case "D"
        If Not NullCheck(Request(fn)) = "" Then
	        rs(fn) = SQLdatetimeADO(Request(fn))
        Else
            If required Then
		        HeaderMSG = "The " & ft & " is required."
		        SaveField = False
		        If Not c = "" Then
		            Execute("Call " & c)
		        End If
            Else
                If Not NewRecord Then
    	            rs(fn) = Null
    	        End If
    	    End If
        End If
    Case Else
        If Not NullCheck(Request(fn)) = "" Then
	        rs(fn) = Trim(Request(fn))
        Else
            If required Then
		        HeaderMSG = "The " & ft & " is required."
		        SaveField = False
		        If Not c = "" Then
		            Execute("Call " & c)
		        End If
            Else
                If Not NewRecord Then
    	            rs(fn) = Null
    	        End If
    	    End If
        End If
    End Select
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for " & ft & " is invalid."
	    SaveField = False
        If Not c = "" Then
            Execute("Call " & c)
        End If
  End If
End Function

'====================================================================================================================================

Function GetLastDay(intMonthNum, intYearNum)
	Dim dNextStart
	If CLng(intMonthNum) = 12 Then
		dNextStart = CDate( "1/1/" & intYearNum)
	Else
		dNextStart = CDate(intMonthNum + 1 & "/1/" & intYearNum)
	End If
	GetLastDay = Day(dNextStart - 1)
End Function

'====================================================================================================================================

Sub Write_CalDay(sValue, sClass)
	Response.Write sValue
End Sub

'====================================================================================================================================

Function CheckDigits(d)
    CheckDigits = d
End Function

'====================================================================================================================================

Sub IPadHeader(rightContent, doHeaderLink)
%>
	<div class='Header'>
		<%If CBool(doHeaderLink) Then%>
			<div style='float:left; cursor:pointer;' onclick="self.location.href = 'default.asp?card=mainmenu&s=<%=SessionID%>';">
				<img src='images/blueIconMC.png' alt='' title='' style='position:relative; top:3px; left:3px; border:none;' />
			</div>
		<%Else%>
			<div style='float:left;'>
				<img src='images/blueIconMC.png' alt='' title='' style='position:relative; top:3px; left:3px; border:none;' />
			</div>
		<%End If%>
		<div style='float:right;'>
			<%=rightContent%>
		</div>
		<div style='clear:both;'></div>
	</div>
	<%If Trim(HeaderMSG) <> "" Then%>
	<div class='Font1' style='padding:3px; font-size:12pt; background-color:#feb7b7; border: 1px solid #c21e1e; color:#c21e1e;'>
		<%=HeaderMSG%>
	</div>
	<%End If%>
<%
End Sub
%>



