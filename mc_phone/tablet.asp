<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>
<!--#INCLUDE FILE="includes/mc_pda_init.asp" -->
<!--#INCLUDE FILE="includes/mc_pda_authenticate.asp" -->
<!--#INCLUDE FILE="includes/mc_pda_support.asp" -->
<%
'====================================================================================================
' SET DEBUG MODE
'====================================================================================================
debug = False

Select Case LCase(card)

case "login"

	Dim inputfieldtype
	Call StartMobileDocument("Maintenance Connection")
	txtmembername = Request.Cookies("m")
	inputfieldtype = "password"
		%>
		<p align="center"><a href="http://www.maintenanceconnection.com" title="Maintenance Connection" target="_blank"><img border="0" src="images/mclive_login.gif" alt="MC Live"/></a></p>

		<p><font style="font-size:10pt;">Please enter your Member ID and Password:</font></p>

		<b>Member&nbsp;ID:</b><input size="15" value="<% =txtmembername %>" tabindex="1" type="text" id="membername" name="membername" format="*M"/><br/>
		<b>Password:</b><input size="15" value="<% =txtpassword %>" tabindex="2" type="<% =inputfieldtype %>" id="password" name="password" format="*M"/>

		<script type="text/javascript">
		try{
			if ($('#membername').val() == ''){
				$('#membername').focus();
			}
			else{
				$('#password').focus();
			}
		} catch(e) {}
		</script>

		<input tabindex="-1" style="width:100%;" type="submit" name="login" value="Login"/>
		<input type="hidden" name="card" value="authenticate"/>
	<%
	EndWMLDocument

case "authenticate"
		authenticate

case "entityselect"
		entityselect

case "entitysearch"
		entitysearch

case "mainmenu"
		mainmenu

case "myworkorders"
		myworkorders

case "allworkorders"
		allworkorders

case "allworkordersu"
		allworkordersu

case "countinventory"
        countinventory

case "assettasks"
        assettasks

case "asoptions"
		asoptions

case "asdetails"
		asdetails

case "asmeters"
		asmeters

case "ashistory"
		ashistory

case "wosearch"
		wosearch

case "cmsearch"
		IsIframe=True
		cmsearch

case "faprsearch"
		IsIframe=True
		fasearch

case "fafasearch"
		IsIframe=True
		fasearch

case "fasosearch"
		IsIframe=True
		fasearch

case "rcsearch"
		IsIframe=True
		rcsearch

case "shsearch"
		IsIframe=True
		shsearch

case "assearch"
		IsIframe=True
		assearch

case "assearch2"
		assearch2

case "clsearch"
		IsIframe=True
		clsearch

case "acsearch"
		IsIframe=True
		acsearch

case "casearch"
		IsIframe=True
		casearch

case "znsearch"
		IsIframe=True
		znsearch

case "prsearch"
		IsIframe=True
		prsearch

case "dpsearch"
		IsIframe=True
		dpsearch

case "tnsearch"
		IsIframe=True
		tnsearch

case "pjsearch"
		IsIframe=True
		pjsearch

case "lasearch"
		IsIframe=True
		lasearch

case "insearch"
		IsIframe=True
		insearch

case "srsearch"
		IsIframe=True
		srsearch

case "wooptions"
		Call wooptions(card)

case "wooptionsh"
		Call wooptions(card)

case "calendarlookup"
		IsIframe=True
        calendarlookup

case "cmlookup"
		IsIframe=True
		cmlookup

case "faprlookup"
		IsIframe=True
		falookup

case "fafalookup"
		IsIframe=True
		falookup

case "fasolookup"
		IsIframe=True
		falookup

case "rclookup"
		IsIframe=True
		rclookup

case "shlookup"
		IsIframe=True
		shlookup

case "aslookup"
		IsIframe=True
		aslookup

case "cllookup"
		IsIframe=True
		cllookup

case "assets"
		assets

case "asforserialpartlookup"
        IsIframe=True
        asforserialpartlookup

case "ltlookup"
		IsIframe=True
		ltlookup

case "aclookup"
		IsIframe=True
		aclookup

case "calookup"
		IsIframe=True
		calookup

case "znlookup"
		IsIframe=True
		znlookup

case "prlookup"
		IsIframe=True
		prlookup

case "dplookup"
		IsIframe=True
		dplookup

case "tnlookup"
		IsIframe=True
		tnlookup

case "pjlookup"
		IsIframe=True
		pjlookup

case "lalookup"
		IsIframe=True
		lalookup

case "inlookup"
		IsIframe=True
		inlookup

case "srlookup"
		IsIframe=True
		srlookup

case "wotasks"
		wotasks

case "wolabor"
		wolabor

case "wopart"
		wopart

case "womisccost"
		womisccost

case "asspecs"
		asspecs

case "aslabor"
		aslabor

case "woassign"
		woassign

case "woonhold"
		woonhold

case "worespond"
		worespond

case "woissue"
		woissue

case "wocomplete"
		wocomplete

case "woclose"
		woclose

case "wonew"
		wonew

case "assetmenu"
		assetmenu

case "inventorymenu"
		inventorymenu

case "wotask"
		wotask

case "wolaborrec"
		wolaborrec

case "wopartrec"
		wopartrec

case "womisccostrec"
		womisccostrec

case "asspecsrec"
		asspecsrec

case "aslaborrec"
		aslaborrec

case "woassignrec"
		woassignrec

case "asphoto"
		asphoto

case "logoff"

	DeleteSessionSQL
	Call StartMobileDocument("Log-off")
		If IsPocketIE or IsBlackBerry Then
		If (lang="HTML") and (MROLive) Then %>
		<p align="center"><a href="http://www.maintenanceconnection.com" title="Maintenance Connection" target="_blank"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/mclive_login.gif" alt="MC Live"/></a></p><%
		Else %>
		<p align="center"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/mcmobile_login.gif" alt="MC Mobile"/></p><%
		End If
		End If
		%>
		<p align="center"><%
		If IsPocketIE Then
		Else %>
		&nbsp;
		<br/><%
		End If
		If (lang="HTML") and (MROLive) Then %>
		Thank you for using MC Live!<%
		Else %>
		Thank you for using MC Mobile!<%
		End If %>
		</p>
		<p align="center">
		<a href="default.asp?card=login">Login</a>
		</p><%
	If lang="HTML" Then
	HTMLbuttonsBegin
	HTMLbuttonsEnd
	End If
	EndWMLDocument

case Else

	Call StartMobileDocument("")
		OutputWAPMsg("The card was not found.")
		WAPbuttonsBegin
		OutputBackButton
		WAPbuttonsEnd
	EndWMLDocument

End Select

Response.End

'====================================================================================================================================
'====================================================================================================================================
'====================================================================================================================================
'====================================================================================================================================

Sub MainMenu()

	Dim db, sql, rs, rs2, wocount, wocount2, wocountu, rccount, shcount
	Dim WOA, WOE, WON, WOFilter, TWC_WOAllOpen, SYADJUSTIN, ASA, ASN

	CardCurrent = "MainMenu"
	CardTitle = "Home"
	CardCurrentLevel = 1

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call WOSave(db)

	ASA = GetAccessRight(db,"ASA",0)
	ASN = GetAccessRight(db,"ASN",0)
	WOA = GetAccessRight(db,"WOA",0)
	WOE = GetAccessRight(db,"WOE",0)
	WON = GetAccessRight(db,"WON",0)
	WOFilter = GetAccessRight(db,"WOFilter",0)
	TWC_WOAllOpen = GetAccessRight(db,"TWC_WOAllOpen",0)
	SYADJUSTIN = GetAccessRight(db,"SYADJUSTIN",0)

	If Not Request("RepairCenterID") = "" Then
		sql = "SELECT * FROM RepairCenter WITH (NOLOCK) WHERE RepairCenterID = '" & NullCheck(Request("RepairCenterID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If Not rs2.eof Then
			Call SetSession("RCPK",rs2("RepairCenterPK"))
			Call SetSession("RCID",rs2("RepairCenterID"))
			Call SetSession("RCNM",rs2("RepairCenterName"))
			Call SetSession("SHPK","")
			Call SetSession("SHID","All")
			Call SetSession("SHNM","All")
			Dim prefvalue, prefdesc, prefpk, shoppk
			If GetPreference(db,True,GetSession("RCPK"),"SU_DefaultShop",prefvalue, prefdesc, prefpk) Then
				sql = "SELECT * FROM Shop WITH (NOLOCK) WHERE ShopPK = " & prefpk & " "
				Set rs2 = db.RunSQLReturnRS(sql,"")
				Call CheckDB(db)
				If Not rs2.eof Then
					Call SetSession("SHPK",rs2("ShopPK"))
					Call SetSession("SHID",rs2("ShopID"))
					Call SetSession("SHNM",rs2("ShopName"))
				End If
			End If
		End If
	End If

	If Not Request("ShopID") = "" and Request("WONew") = "" Then
		If Request("ShopID") = "ALL" Then
			Call SetSession("SHPK","")
			Call SetSession("SHID","All")
			Call SetSession("SHNM","All")
		Else
			sql = "SELECT * FROM Shop WITH (NOLOCK) WHERE ShopID = '" & NullCheck(Request("ShopID")) & "' "
			Set rs2 = db.RunSQLReturnRS(sql,"")
			Call CheckDB(db)
			If Not rs2.eof Then
				Call SetSession("SHPK",rs2("ShopPK"))
				Call SetSession("SHID",rs2("ShopID"))
				Call SetSession("SHNM",rs2("ShopName"))
			End If
		End If
	End If

	If GetSession("SHPK") = "" Then
		Call SetSession("SHID","All")
		Call SetSession("SHNM","All")
	End If

	sql = _
	"SELECT WOCount = COUNT(DISTINCT WO.WOPK) " &_
	"FROM WO WITH (NOLOCK) " &_
	"INNER JOIN WOassign WITH (NOLOCK) ON WO.WOPK = WOassign.WOPK " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.Complete Is Null " &_
	"AND WOAssign.LaborPK = '" & GetSession("UserPK") & "' " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount = "0"
	Else
		wocount = rs("WOCount")
	End If

	sql = _
	"SELECT WOCount = COUNT(WO.WOPK) " &_
	"FROM WO WITH (NOLOCK) " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.Complete Is Null " &_
	"AND WO.RepairCenterPK = " & GetSession("RCPK") & " " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "

	sql = sql & AddGeneralFilters("WO")

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount2 = "0"
	Else
		wocount2 = rs("WOCount")
	End If

	sql = _
	"SELECT WOCount = COUNT(WO.WOPK) " &_
	"FROM WO WITH (NOLOCK) " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.Complete Is Null " &_
	"AND WO.IsAssigned = 0 " &_
	"AND WO.RepairCenterPK = " & GetSession("RCPK") & " " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "

	sql = sql & AddGeneralFilters("WO")

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocountu = "0"
	Else
		wocountu = rs("WOCount")
	End If

	sql = _
	"SELECT RCCount = COUNT(RepairCenterPK) " &_
	"FROM RepairCenter WITH (NOLOCK) " &_
	"WHERE Active = 1 "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		rccount = "0"
	Else
		rccount = rs("RCCount")
	End If

	sql = _
	"SELECT SHCount = COUNT(ShopPK) " &_
	"FROM Shop WITH (NOLOCK) " &_
	"WHERE Active = 1 " &_
	"AND RepairCenterPK = " & GetSession("RCPK") & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		shcount = "0"
	Else
		shcount = rs("SHCount")
	End If

	Call StartMobileDocument(CardTitle)
        If lang="WML" Then
        If IsPocketIE or IsBlackBerry Then %>
        <onevent type="onenterforward"><refresh><setvar name="password" value=""/></refresh></onevent><%
        End If
        End If
		If IsPocketIE or IsBlackBerry Then
		If Not GetSession("el") = "" Then %>
		<p align="center"><img<% If lang="HTML" Then %> border="0"<% End If %> src="<% =showclientimg(GetSession("el")) %>" alt="MC Mobile"/></p>
		<% Else
		If (lang="HTML") and (MROLive) Then %>
		<p align="center"><a href="http://www.maintenanceconnection.com" title="Maintenance Connection" target="_blank"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/mclive_login.gif" alt="MC Live"/></a></p><%
		Else %>
		<p align="center"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/mcmobile_login.gif" alt="MC Mobile"/></p><%
		End If
		End If
		Else
		%>
		<p mode="nowrap" align="center">
		<b><% =GetSession("en") %></b></p>
		<% End If %>
		<p align="center"><%
		If WOA Then
		If wocount > 0 Then %>
		<b><%
		End If %>
		<a href="default.asp?card=myworkorders&amp;s=<% =SessionID %>">My Work Orders (<% =wocount %>)</a><br/><%
		If wocount > 0 Then %>
		</b><%
		End If
		End If
		If TWC_WOAllOpen Then %>
		<a href="default.asp?card=allworkorders&amp;s=<% =SessionID %>">All Work Orders (<% =wocount2 %>)</a><br/>
		<a href="default.asp?card=allworkordersu&amp;s=<% =SessionID %>">Unassigned Work Orders (<% =wocountu %>)</a><br/><%
		End If
		%>
		<a href="default.asp?card=assettasks&amp;s=<% =SessionID %>">Asset Tasks</a><br/><%
		If WON Then %>
		<a href="default.asp?card=wonew&amp;s=<% =SessionID %>">New Work Order</a><br/><%
		End If
		If False and GetSession("mcsupport")="Y" Then %><a href="default.asp?card=prospects&amp;s=<% =SessionID %>">Prospects</a><br/><%
		End If
		If ASA Then %><a href="default.asp?card=assets&amp;s=<% =SessionID %>"><% If GetSession("mcsupport")="Y" Then %>Customers<% Else %>Assets<% End If %></a><br/><%
		End If
		If ASN Then %><a href="default.asp?card=asdetails&amp;s=<% =SessionID %>&amp;assetpk=-1"><% If GetSession("mcsupport")="Y" Then %>New Customer<% Else %>New Asset<% End If %></a><br/><%
		End If
		If SYADJUSTIN Then %>
			<a href="default.asp?card=countinventory&amp;s=<% =SessionID %>">Count Inventory</a><br/><%
		End If
		If False Then %><a href="default.asp?card=assetmenu&amp;s=<% =SessionID %>">Asset Menu</a><br/><%
		End If
		If False Then %><a href="default.asp?card=inventorymenu&amp;s=<% =SessionID %>">Inventory Menu</a><br/><%
		End If
		If WOFilter Then
		    If rccount > 1 or shcount > 1 Then
		        'If IsPocketIE or IsBlackBerry Then
		            Response.Write "<br/><b>Criteria:</b><br/>"
		        'End If
		    End If
		    If rccount > 1 Then %>
	    	    <a <% If lang="HTML" Then %>target="mciframe" <% End If %>href="default.asp?card=rclookup&amp;s=<% =SessionID %>"><% =GetSession("RCNM") %></a><%
		    End If %>
		    <% If shcount > 1 Then %>
		        <% If rccount > 1 Then %>
		            <% If IsPocketIE or IsBlackBerry Then %>
                        <% Response.Write " / " %>
                    <% Else %>
                        <% Response.Write "<br/>" %>
                    <% End If %>
		        <% End If %>
    		    <a <% If lang="HTML" Then %>target="mciframe" <% End If %>href="default.asp?card=shlookup&amp;s=<% =SessionID %>">Shop: <% =GetSession("SHID") %></a><br/><%
		    End If
		End If
		If False Then %><a href="default.asp?card=entertime&amp;s=<% =SessionID %>">Enter Time</a><br/><%
		End If
		If GetSession("ed") = "Y" Then
		If Not GetSession("searchby") = "" Then %>
		<br/><a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=<% =GetSession("SearchBy") %>">Change Company</a><br/><%
		Else %>
		<br/><a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=NONE">Change Company</a><br/><%
		End If %>
		<% End If %>
		<% If Not GetSession("ed") = "Y" Then %>
		<br/>
		<% End If %>
		<% If IsPocketIE Then %>
		<a href="default.asp?card=mainmenu&amp;s=<% =SessionID %>">Refresh</a><br/>
		<% End If %>
		<a href="default.asp?card=logoff&amp;s=<% =SessionID %>">Log-off</a><br/>
		</p><% If lang = "WML" Then %>
		<% If IsPocketIE or IsBlackBerry Then %>
		<do name="button1" type="accept" label="Log-off">
			<go href="default.asp" method="get">
				<postfield name="card" value="logoff"/>
				<postfield name="s" value="<% =SessionID %>"/>
			</go>
		</do>
		<do name="button2" type="accept" label="Refresh">
			<go href="default.asp" method="get">
				<postfield name="card" value="mainmenu"/>
				<postfield name="s" value="<% =SessionID %>"/>
			</go>
		</do>
		<% Else %>
		<do name="button2" type="accept" label="Refresh">
			<go href="default.asp" method="get">
				<postfield name="card" value="mainmenu"/>
				<postfield name="s" value="<% =SessionID %>"/>
			</go>
		</do>
		<do name="button1" type="accept" label="Log-off">
			<go href="default.asp" method="get">
				<postfield name="card" value="logoff"/>
				<postfield name="s" value="<% =SessionID %>"/>
			</go>
		</do><%
		End If
		Else
		HTMLbuttonsBegin
		%>
		<input style="width:100%;" type="button" name="LogOff" value="Log-Off" onclick="self.location.href='default.asp?card=logoff&s=<% =SessionID %>';"/>
		</td><td align="right">
		<input style="width:100%;" type="button" name="Refresh" value="Refresh" onclick="self.location.href='default.asp?card=mainmenu&s=<% =SessionID %>';"/>
		<%
		HTMLbuttonsEnd
		End If
		rs.close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOOptions(n)

	Dim db, sql, rs, wocount, wocount2, wocount3, wocount4, wocount5, wocount6, woid, reason, assetid, assetname, closeddate, photo, wostatus, wostatusdesc, wostatusdate, wostatustime, parentlocationall, parentequipmentall, requestedline, targetdate, FromAssetPK, departmentline

	CardTitle = "WO Options"
	If n = "" Then
		CardCurrent = "WOOptions"
	Else
		CardCurrent = n
	End If
	CardCurrentLevel = GetCardLevel()

	'Call ASPDebug

	GetWOPK

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call WOCloseSave(db)

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!') " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		Call OutputWAPError("The Work Order was not found.")
	Else
		woid = NullCheck(rs("WOID"))
		If IsPocketIE or IsBlackBerry Then
		reason = WAPValidate(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)))
		Else
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),500))
		End If
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(rs("IsOpen"))
		wostatus = WAPValidate(NullCheck(RS("Status")))
		wostatusdesc = WAPValidate(NullCheck(RS("StatusDesc")))
		requestedline = NullCheck(RS("RequesterName"))
		If Not requestedline = "" and Not NullCheck(RS("RequesterPhone")) = "" Then
		requestedline = requestedline & " - " & NullCheck(RS("RequesterPhone"))
	    End If

        departmentline = NullCheck(RS("DepartmentName"))

        parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!","<br/>")
        parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!","<br/>")
        If Not parentlocationall = "" Then
            parentlocationall = parentlocationall & "<br/>"
        End If
        If Not parentequipmentall = "" Then
            parentequipmentall = parentequipmentall & "<br/>"
        End If

		targetdate = WAPValidate(DateNullCheck(RS("TargetDate")))

		Select Case UCase(wostatus)
			Case "REQUESTED"
				wostatusdate = WAPValidate(DateNullCheck(RS("Requested")))
				wostatustime = WAPValidate(TimeNullCheck(RS("Requested")))
			Case "ISSUED"
				If NullCheck(RS("Responded")) = "" Then
					wostatusdate = WAPValidate(DateNullCheck(RS("Issued")))
					wostatustime = WAPValidate(TimeNullCheck(RS("Issued")))
				ElseIf NullCheck(RS("Complete")) = "" Then
					wostatusdesc = wostatusdesc & " / Responded"
					wostatusdate = WAPValidate(DateNullCheck(RS("Responded")))
					wostatustime = WAPValidate(TimeNullCheck(RS("Responded")))
				ElseIf NullCheck(RS("Finalized")) = "" Then
					wostatusdesc = wostatusdesc & " / Completed"
					wostatusdate = WAPValidate(DateNullCheck(RS("Complete")))
					wostatustime = WAPValidate(TimeNullCheck(RS("Complete")))
				Else
					wostatusdesc = wostatusdesc & " / Finalized"
					wostatusdate = WAPValidate(DateNullCheck(RS("Finalized")))
					wostatustime = WAPValidate(TimeNullCheck(RS("Finalized")))
				End If
			Case "DENIED"
				wostatusdate = WAPValidate(DateNullCheck(RS("Denied")))
				wostatustime = WAPValidate(TimeNullCheck(RS("Denied")))
			Case "CLOSED"
				wostatusdate = WAPValidate(DateNullCheck(RS("Closed")))
				wostatustime = WAPValidate(TimeNullCheck(RS("Closed")))
			Case "CANCELED"
				wostatusdate = WAPValidate(DateNullCheck(RS("Canceled")))
				wostatustime = WAPValidate(TimeNullCheck(RS("Canceled")))
			Case "ONHOLD"
				wostatusdate = WAPValidate(DateNullCheck(RS("OnHold")))
				wostatustime = WAPValidate(TimeNullCheck(RS("OnHold")))
		End Select
	End If

	sql = _
	"SELECT WOCount = COUNT(WOPK) " &_
	"FROM WOTask WITH (NOLOCK) " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount = "0"
	Else
		wocount = rs("WOCount")
	End If

	sql = _
	"SELECT WOCount = COUNT(WOPK) " &_
	"FROM WOLabor WITH (NOLOCK) " &_
	"WHERE RecordType = 2 AND WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount2 = "0"
	Else
		wocount2 = rs("WOCount")
	End If

	sql = _
	"SELECT WOCount = COUNT(WOPK) " &_
	"FROM WOPart WITH (NOLOCK) " &_
	"WHERE RecordType = 2 AND WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount3 = "0"
	Else
		wocount3 = rs("WOCount")
	End If

	sql = _
	"SELECT WOCount = COUNT(WOPK) " &_
	"FROM WOMiscCost WITH (NOLOCK) " &_
	"WHERE RecordType = 2 AND WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount4 = "0"
	Else
		wocount4 = rs("WOCount")
	End If

	If Not AssetPK = "" Then
        sqlwhere = " WHERE WO.AssetPK =" & AssetPK & " "
	    sql = _
	    "SELECT WOCount = Count(WO.WOPK) " + nl + _
	    "FROM         WO WITH (NOLOCK) LEFT OUTER JOIN " + nl + _
	    "                      LookupTableValues ltp WITH (NOLOCK) ON WO.Priority = ltp.CodeName AND ltp.LookupTable = 'PRIORITY' INNER JOIN " + nl + _
	    "                      LookupTableValues lts WITH (NOLOCK) ON WO.Status = lts.CodeName AND lts.LookupTable = 'WOSTATUS' INNER JOIN " + nl + _
	    "                      LookupTableValues ltas WITH (NOLOCK) ON WO.AuthStatus = ltas.CodeName AND ltas.LookupTable = 'WOAUTHSTATUS' LEFT OUTER JOIN " + nl + _
	    "                      AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK LEFT OUTER JOIN " + nl + _
	    "					   Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " + nl + _
	    sqlwhere + " " + nl + _
	    "UNION " + nl + _
	    "SELECT WOCount = Count(WO.WOPK) " + nl + _
	    "FROM         WOTask WITH (NOLOCK) INNER JOIN " + nl + _
        "                      WO WITH (NOLOCK) ON WOTask.WOPK = WO.WOPK LEFT OUTER JOIN " + nl + _
	    "                      LookupTableValues ltp WITH (NOLOCK) ON WO.Priority = ltp.CodeName AND ltp.LookupTable = 'PRIORITY' INNER JOIN " + nl + _
	    "                      LookupTableValues lts WITH (NOLOCK) ON WO.Status = lts.CodeName AND lts.LookupTable = 'WOSTATUS' INNER JOIN " + nl + _
	    "                      LookupTableValues ltas WITH (NOLOCK) ON WO.AuthStatus = ltas.CodeName AND ltas.LookupTable = 'WOAUTHSTATUS' LEFT OUTER JOIN " + nl + _
	    "                      AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK LEFT OUTER JOIN " + nl + _
	    "					   Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " + nl + _
	    REPLACE(sqlwhere,"WO.AssetPK =","WOTask.AssetPK = ") + " " + nl

		Set rs = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)

		If rs.eof Then
			wocount5 = "0"
		Else
		    wocount5 = CLng(rs("WOCount"))
		    rs.MoveNext
		    If Not NullCheck(rs("WOCount")) = "" Then
		        wocount5 = wocount5 + CLng(rs("WOCount"))
		    End If
		End If
	Else
		wocount5 = "0"
	End If

	sql = _
	"SELECT WOCount = COUNT(PK) " &_
	"FROM WOassign WITH (NOLOCK) " &_
	"WHERE (WOassign.WOPK = " & WOPK & " AND WOassign.IsAssigned = 1 AND (WOassign.Active = 1 or WOassign.Active Is Null)) "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount6 = "0"
	Else
		wocount6 = rs("WOCount")
	End If

	Dim AccessToIssue, AccessToAssign, AccessToRespond, AccessToComplete, AccessToFinalize, AccessToClose

    AccessToIssue = GetAccessRight(db,"WOIssue",0)
	AccessToAssign = GetAccessRight(db,"WOAssign",0)
	AccessToRespond = GetAccessRight(db,"WORespond",0)
	AccessToComplete = GetAccessRight(db,"WOComplete",0)
	AccessToFinalize = GetAccessRight(db,"WOFinalize",0)
	AccessToClose = GetAccessRight(db,"WOCloseFromDialog",0)

	Call StartMobileDocument(CardTitle)
	%>
		<% If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap"><%
		If Not IsBOF Then %>
		<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;pr=<% =WOPK %>">&lt;</a>&nbsp;<%
		End If %><b><a href="default.asp?card=<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>&amp;s=<% =SessionID %>&amp;back=1">WO #<% =WOID %></a></b><%
		If Not IsEOF Then %>
		<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;nr=<% =WOPK %>">&gt;</a><%
		End If %>
		<br/><%
		If Not AssetPK = "" Then %>
		<% =HStyleBegin %><% =ParentLocationAll %><% =ParentEquipmentAll %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
		<% End If %>
		<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
		<% If Not RequestedLine = "" Then %>
		<br/><% =HStyleBegin %><% =RequestedLine %><% =HStyleEnd %>
		<% End If %>

		<% If Not DepartmentLine = "" Then %>
		<br/><% =HStyleBegin %><% =DepartmentLine %><% =HStyleEnd %>
		<% End If %>

		<% If Not IsOpen Then %>
		<br/><% =HStyleBegin %><b>Closed: <% =wostatusdate %>&nbsp;<% =wostatustime %></b><% =HStyleEnd %>
		<% Else %>
		<br/><% =HStyleBegin %>Target Date:&nbsp; <% =TargetDate %><% =HStyleEnd %>
		<br/><% =HStyleBegin %><% =wostatusdesc %>: <% =wostatusdate %>&nbsp; <% =wostatustime %><% =HStyleEnd %>
		<% End If
		If lang = "WML" and IsPPC Then
		Response.Write "<br/>"
		Response.Write HStyleBegin
		OutputBackButton
		OutputHomeButton
		Response.Write HStyleEnd
		End If %>
		</p><%
		End If %>
		<p align="center" mode="wrap">
		<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Tasks (<% =wocount %>)</a><br/>
		<a href="default.asp?card=wolabor&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Labor (<% =wocount2 %>)</a><br/>
		<a href="default.asp?card=wopart&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Materials (<% =wocount3 %>)</a><br/>
		<% If IsPocketIE or IsBlackBerry Then %>
		<a href="default.asp?card=womisccost&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Other Costs (<% =wocount4 %>)</a><br/>
		<% If AccessToAssign Then %>
		<a href="default.asp?card=woassign&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Assignments (<% =wocount6 %>)</a><br/>
		<% End If %>
		<% Else %>
		<a href="default.asp?card=womisccost&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Other (<% =wocount4 %>)</a><br/>
		<% If AccessToAssign Then %>
		<a href="default.asp?card=woassign&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Assign (<% =wocount6 %>)</a><br/>
		<% End If %>
		<% End If %>
		<% If IsOpen Then %>
		<% If UCase(wostatus) = "REQUESTED" Then %>
		<% If AccessToIssue and (True or IsPocketIE or IsBlackBerry or IsWAP20) Then %>
		<b><a href="default.asp?card=woissue&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Issue</a></b><% If IsPPC Then %>&nbsp;<% Else %><br/><% End If %>
		<% End If %>
		<% End If %>
		<% If AccessToRespond and (True or IsPocketIE or IsBlackBerry or IsWAP20) Then %>
		<b><a href="default.asp?card=worespond&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Respond</a></b><% If IsPPC Then %>&nbsp;<% Else %><br/><% End If %>
		<% End If %>
		<% If Not UCase(wostatus) = "ONHOLD" and (IsPocketIE or IsBlackBerry or IsWAP20) Then %>
		<b><a href="default.asp?card=woonhold&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">On-Hold</a></b><br/>
		<% End If %>
		<% If AccessToComplete Then %>
		<b><a href="default.asp?card=wocomplete&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Complete</a></b><% If IsPPC Then %>&nbsp;<% Else %><br/><% End If %>
		<% End If %>
		<% If AccessToClose Then %>
		<b><a href="default.asp?card=woclose&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>">Close</a></b><br/>
		<% End If %>
		<% End If %>
		<a href="default.asp?card=ashistory&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Asset History (<% =wocount5 %>)</a><br/>
		</p>
		<% If (IsPocketIE or IsBlackBerry or IsWAP20) and Not Photo = "" and Not AssetPK = "" Then %>
		<p align="center"><a href="default.asp?card=asphoto&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>"><img<% If lang="HTML" Then %> border="0"<% End If %> src="<% =Application("ImageServer") & Replace(LCase(Photo),"main","wo") %>" alt="Asset Photo"/></a></p>
		<% End If
		OutputButtons
		rs.close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASOptions()

	Dim db, sql, rs, ascount, ascount2, ascount3, ascount4, ascount5, ascount6, ascount7, assetid, assetname, photo, parentlocationall, parentequipmentall, ASN

	CardTitle = "Asset Menu"
	CardCurrent = "ASOptions"

	GetAssetPK

	Set db = New ADOHelper

	Call ASDetailsSave(db)
	Call ASMetersSave(db)

    If Trim(UCase(CardFrom)) = "ASDETAILS" and Request("AssetPK") = "-1" Then
        CardSkipLevel = -1
    End If

	CardCurrentLevel = GetCardLevel()
    Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
    Call SetSession("CardFrom",CardCurrent)
    Call SetSession("CardFromLevel",CardCurrentLevel)

	ASN = GetAccessRight(db,"ASN",0)

	sql = _
	"SELECT Asset.*, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!') " &_
	"FROM Asset WITH (NOLOCK) INNER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
	"WHERE Asset.AssetPK = " & AssetPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If Not rs.eof Then
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		photo = NullCheck(rs("Photo"))
        parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!","<br/>")
        parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!","<br/>")
        If Not parentlocationall = "" Then
            parentlocationall = parentlocationall & "<br/>"
        End If
        If Not parentequipmentall = "" Then
            parentequipmentall = parentequipmentall & "<br/>"
        End If
	Else
		Call OutputWAPError("The Asset was not found.")
	End If

	sql = "SELECT ASCount = COUNT(PK) " &_
		  "FROM AssetSpecification WITH (NOLOCK) " &_
		  "WHERE AssetPK = " & AssetPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		ascount = "0"
	Else
		ascount = rs("ASCount")
	End If

	sql = "SELECT ASCount = COUNT(PK) " &_
		  "FROM AssetLabor WITH (NOLOCK) " &_
		  "WHERE AssetPK = " & AssetPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		ascount2 = "0"
	Else
		ascount2 = rs("ASCount")
	End If

	sql = "SELECT ASCount = COUNT(PK) " &_
		  "FROM AssetTask WITH (NOLOCK) " &_
		  "WHERE AssetPK = " & AssetPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		ascount3 = "0"
	Else
		ascount3 = rs("ASCount")
	End If

	sql = "SELECT ASCount = COUNT(PK) " &_
		  "FROM AssetPart WITH (NOLOCK) " &_
		  "WHERE AssetPK = " & AssetPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		ascount6 = "0"
	Else
		ascount6 = rs("ASCount")
	End If

	sql = "SELECT ASCount = COUNT(PK) " &_
		  "FROM AssetContract WITH (NOLOCK) " &_
		  "WHERE AssetPK = " & AssetPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		ascount7 = "0"
	Else
		ascount7 = rs("ASCount")
	End If

	sql = "SELECT ASCount = COUNT(PK) " &_
		  "FROM PMAsset WITH (NOLOCK) " &_
		  "WHERE AssetPK = " & AssetPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		ascount4 = "0"
	Else
		ascount4 = rs("ASCount")
	End If

    sqlwhere = " WHERE WO.AssetPK =" & AssetPK & " "
	sql = _
	"SELECT ASCount = Count(WO.WOPK) " + nl + _
	"FROM         WO WITH (NOLOCK) LEFT OUTER JOIN " + nl + _
	"                      LookupTableValues ltp WITH (NOLOCK) ON WO.Priority = ltp.CodeName AND ltp.LookupTable = 'PRIORITY' INNER JOIN " + nl + _
	"                      LookupTableValues lts WITH (NOLOCK) ON WO.Status = lts.CodeName AND lts.LookupTable = 'WOSTATUS' INNER JOIN " + nl + _
	"                      LookupTableValues ltas WITH (NOLOCK) ON WO.AuthStatus = ltas.CodeName AND ltas.LookupTable = 'WOAUTHSTATUS' LEFT OUTER JOIN " + nl + _
	"                      AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK LEFT OUTER JOIN " + nl + _
	"					   Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " + nl + _
	sqlwhere + " " + nl + _
	"UNION " + nl + _
	"SELECT ASCount = Count(WO.WOPK) " + nl + _
	"FROM         WOTask WITH (NOLOCK) INNER JOIN " + nl + _
    "                      WO WITH (NOLOCK) ON WOTask.WOPK = WO.WOPK LEFT OUTER JOIN " + nl + _
	"                      LookupTableValues ltp WITH (NOLOCK) ON WO.Priority = ltp.CodeName AND ltp.LookupTable = 'PRIORITY' INNER JOIN " + nl + _
	"                      LookupTableValues lts WITH (NOLOCK) ON WO.Status = lts.CodeName AND lts.LookupTable = 'WOSTATUS' INNER JOIN " + nl + _
	"                      LookupTableValues ltas WITH (NOLOCK) ON WO.AuthStatus = ltas.CodeName AND ltas.LookupTable = 'WOAUTHSTATUS' LEFT OUTER JOIN " + nl + _
	"                      AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK LEFT OUTER JOIN " + nl + _
	"					   Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " + nl + _
	REPLACE(sqlwhere,"WO.AssetPK =","WOTask.AssetPK = ") + " " + nl

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		ascount5 = "0"
	Else
		ascount5 = CLng(rs("ASCount"))
		rs.MoveNext
		If Not NullCheck(rs("ASCount")) = "" Then
		    ascount5 = ascount5 + CLng(rs("ASCount"))
		End If
	End If

	Call StartMobileDocument(CardTitle)
		If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap"><b><% =HStyleBegin %><% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)<% =HStyleEnd %></b></p><%
		End If %>
		<p align="center" mode="wrap">
		<a href="default.asp?card=asdetails&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Details</a><br/>
		<a href="default.asp?card=asmeters&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Meters</a><br/>
		<a href="default.asp?card=asspecs&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Specifications (<% =ascount %>)</a><br/>
		<% If False Then %>
		<a href="default.asp?card=asparts&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Materials (<% =ascount6 %>)</a><br/>
		<% End If %>
		<a href="default.asp?card=aslabor&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Labor / Contacts (<% =ascount2 %>)</a><br/>
		<% If False Then %>
		<a href="default.asp?card=ascontracts&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Contracts (<% =ascount7 %>)</a><br/>
		<a href="default.asp?card=aspms&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">PMs (<% =ascount4 %>)</a><br/>
		<a href="default.asp?card=astasks&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Tracked Tasks (<% =ascount3 %>)</a><br/>
		<% End If %>
		<a href="default.asp?card=ashistory&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">History (<% =ascount5 %>)</a><br/>
		<% If False and ASN Then %><br/><a href="default.asp?card=asdetails&amp;s=<% =SessionID %>&amp;assetpk=-1&amp;parentid=<% =AssetID %>"><% If GetSession("mcsupport")="Y" Then %>New Asset<% Else %>New Asset<% End If %></a><br/><%
		End If %>
		</p>
		<% If (IsPocketIE or IsBlackBerry) and Not Photo = "" Then %>
		<p align="center"><a href="default.asp?card=asphoto&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><img<% If lang="HTML" Then %> border="0"<% End If %> src="<% =Application("ImageServer") & Replace(LCase(Photo),"main","wo") %>" alt="Asset Photo"/></a></p>
		<% End If
		OutputButtons
		rs.close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASDetails()

	Dim db, sql, rs, ASE, assetid, assetname, photo, parentlocationall, parentequipmentall, newrec, IsLocation, rseof

    newcontextpage = True

	CardTitle = "Asset Details"
	CardCurrent = "ASDetails"
	CardCurrentLevel = GetCardLevel()

	GetAssetPK

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	ASE = GetAccessRight(db,"ASE",0)

    If Not AssetPK = "-1" Then
	    sql = _
	    "SELECT Asset.*, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!') " &_
	    "FROM Asset WITH (NOLOCK) INNER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
	    "WHERE Asset.AssetPK = " & AssetPK & " "

	    Set rs = db.RunSQLReturnRS(sql,"")
	    Call CheckDB(db)

	    If Not rs.eof Then
		    assetpk = NullCheck(rs("AssetPK"))
		    assetid = WAPValidate(NullCheck(RS("AssetID")))
		    assetname = WAPValidate(NullCheck(RS("AssetName")))
		    photo = NullCheck(rs("Photo"))
            parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!","<br/>")
            parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!","<br/>")
            If Not parentlocationall = "" Then
                parentlocationall = parentlocationall & "<br/>"
            End If
            If Not parentequipmentall = "" Then
                parentequipmentall = parentequipmentall & "<br/>"
            End If
	    Else
		    Call OutputWAPError("The Asset was not found.")
	    End If

	    sql = _
	    "SELECT     Asset.* " &_
	    "FROM       Asset WITH (NOLOCK)  " &_
	    "WHERE     (Asset.AssetPK = " & AssetPK & " )"

	    Set rs = db.RunSQLReturnRS_RW(sql,"")
	    Call CheckDB(db)

        IsLocation = BitNullCheck(rs("IsLocation"))

        If rs.eof Then
            rseof = True
        Else
            rseof = False
        End If

	End If

    If AssetPK = "-1" Then
	Call BuildFields("ParentID","Parent ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("ClassificationID","Classification ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("AssetID","Asset ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("AssetName","Asset Name","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Else
	'Call BuildFields("ParentID","Parent ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
    End If
    If IsLocation and Not AssetPK = "-1" Then
	Call BuildFields("Address","Address","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("City","City","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("State","State","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Zip","Zip","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Country","Country","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
    End If
	Call BuildFields("Model","Model","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Serial","Serial #","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("RotatingPartID","Inventory Item","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Priority","Priority","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("RiskLevel","Risk Level","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("VendorID","Vendor ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("ManufacturerID","Manufacturer ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("WarrantyExpire","Warranty Ends","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("PurchasedDate","Purchase Date","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("InstallDate","Install Date","C",GlobalFieldLength,"*M","true",CStr(DateNullCheck(Date())),rs,AssetPK)
	Call BuildFields("ReplaceDate","Replace Date","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("DisposalDate","Disposal Date","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("RepairCenterID","Repair Center ID","C",GlobalFieldLength,"*M","true",GetSession("RCID"),rs,AssetPK)
	Call BuildFields("ShopID","Shop ID","C",GlobalFieldLength,"*M","true",GetSession("SHID"),rs,AssetPK)
	Call BuildFields("AccountID","Account ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("DepartmentID","Department ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("TenantID","Customer ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("ContactID","Contact ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("OperatorID","Operator ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("ZoneID","Zone ID","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Vicinity","Vicinity","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
    If AssetPK = "-1" Then
	Call BuildFields("Address","Address","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("City","City","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("State","State","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Zip","Zip","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Country","Country","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
    End If
	Call BuildFields("Status","Sub-Status","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("IsUp","In-Service?","B","","","",True,rs,AssetPK)

	Call StartMobileDocument(CardTitle)
		If rseof and Not AssetPK = "-1" Then
			Call OutputWAPMsg("The Asset was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<% If AssetPK = "-1" Then %>
		    <b><% =HStyleBegin %>New Asset<% =HStyleEnd %></b>
			<% Else %>
		    <b><% =HStyleBegin %><% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)<% =HStyleEnd %></b>
		    <% End If %>
			</p>
			<% End If
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			' We do not need to check for ASN because they would not have AssetPK = -1 if they could not get here from the main menu
			If IsOpen and (ASE or AssetPK = "-1") Then
			ASDetailsSubmit rs
			End If
		Else
			If IsOpen and (ASE or AssetPK = "-1") Then
			ASDetailsSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen and (ASE or AssetPK = "-1") Then
		%>
		</td><td align="right">
		<%
		ASDetailsSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		Call CloseObj(rs)
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub ASDetailsSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="ASOptions"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="Assetpk" value="<% =AssetPK %>"/>
			<postfield name="POSTEDASDETAILS" value="Y"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="ASOptions"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="Assetpk" value="<% =AssetPK %>"/>
	<input type="hidden" name="POSTEDASDETAILS" value="Y"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub ASDetailsSave(ByRef db)

	'ASPDebug
	'Response.End

	Dim rs2, LaborPK, WorkDate, ParentPK

	If Not Request("POSTEDASDETAILS") = "" Then
    	Assetpk = Request("Assetpk")
		sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetPK = " & Assetpk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If Assetpk = "-1" Then
			rs.addnew()
		End If
		If Not rs.Eof Then
			If UCase(Request("Delete")) = "Y" or UCase(Request("Delete")) = "2" Then
				rs.delete()
				Call SetSession("AssetPK" & CardCurrentLevel,"")
                Call SetSession("AssetPK" & CardCurrentLevel+1,"")
			Else
				On Error Resume Next
				If Assetpk = "-1" Then
				    NewRecord = True
				    If Not NullCheck(Request("ParentID")) = "" Then
					    sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetID = '" & Request("ParentID") & "' "
					    Set rs2 = db.RunSQLReturnRS(sql,"")
					    Call CheckDB(db)
					    If rs2.eof Then
						    HeaderMSG = "The Parent ID was not found."
						    ASDetails
						Else
						    ParentPK = rs2("AssetPK")
					    End If
					Else
					    HeaderMSG = "The Parent ID is required."
					    ASDetails
					End If
					Call SaveField(db,"ClassificationID","Classification ID","Classification","","LM","ASDetails",True)
                    'Call SaveField(db,"AssetID","Asset ID","","","C","ASDetails",True)
				    If Not NullCheck(Request("AssetID")) = "" Then
					    sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetID = '" & Request("AssetID") & "' "
					    Set rs2 = db.RunSQLReturnRS(sql,"")
					    Call CheckDB(db)
					    If Not rs2.eof Then
						    HeaderMSG = "The Asset ID has already been assigned to another Asset (" & WAPValidate(NullCheck(rs2("AssetName"))) & ")."
						    ASDetails
						Else
						    rs("AssetID") = Trim(Request("AssetID"))
					    End If
					Else
					    HeaderMSG = "The Asset ID is required."
					    ASDetails
					End If
                    Call SaveField(db,"AssetName","Asset Name","","","C","ASDetails",True)
				End If

				If rs("IsLocation") or Assetpk = "-1" Then
                Call SaveField(db,"Address","","","","C","ASDetails",False)
                Call SaveField(db,"City","","","","C","ASDetails",False)
                Call SaveField(db,"State","","","","C","ASDetails",False)
                Call SaveField(db,"Zip","","","","C","ASDetails",False)
                Call SaveField(db,"Country","","","","C","ASDetails",False)
                End If
                Call SaveField(db,"Model","","","","C","ASDetails",False)
                Call SaveField(db,"Serial","","","","C","ASDetails",False)
                Call SaveField(db,"RotatingPartID","Inventory Item","Part","PartID","LM","ASDetails",False)
                Call SaveField(db,"Priority","","WOPriority","","LT","ASDetails",False)
                Call SaveField(db,"RiskLevel","Risk Level","","","N","ASDetails",False)
                Call SaveField(db,"VendorID","Vendor ID","Company","CompanyID","LM","ASDetails",False)
                Call SaveField(db,"ManufacturerID","Manufacturer ID","Company","CompanyID","LM","ASDetails",False)
                Call SaveField(db,"WarrantyExpire","Warranty Ends","","","D","ASDetails",False)
                Call SaveField(db,"PurchasedDate","Purchase Date","","","D","ASDetails",False)
                Call SaveField(db,"InstallDate","Install Date","","","D","ASDetails",False)
                Call SaveField(db,"ReplaceDate","Replace Date","","","D","ASDetails",False)
                Call SaveField(db,"DisposalDate","Disposal Date","","","D","ASDetails",False)
                Call SaveField(db,"RepairCenterID","Repair Center ID","RepairCenter","","LM","ASDetails",False)
                Call SaveField(db,"ShopID","Shop ID","Shop","","LM","ASDetails",False)
                Call SaveField(db,"AccountID","Account ID","Account","","LM","ASDetails",False)
                Call SaveField(db,"DepartmentID","Department ID","Department","","LM","ASDetails",False)
                Call SaveField(db,"TenantID","Customer ID","Tenant","","LM","ASDetails",False)
                Call SaveField(db,"ContactID","Contact ID","Labor","LaborID","LM","ASDetails",False)
                Call SaveField(db,"OperatorID","Opertator ID","Labor","LaborID","LM","ASDetails",False)
                Call SaveField(db,"ZoneID","Zone ID","Zone","","LM","ASDetails",False)
                Call SaveField(db,"Vicinity","","","","C","ASDetails",False)
                Call SaveField(db,"Status","Sub-Status","assetstatus","","LT","ASDetails",False)
                Call SaveField(db,"IsUp","In-Service","","","B","ASDetails",False)

				On Error Goto 0

			End If
			db.dobatchupdate rs
			If db.dok Then
			    If AssetPK = "-1" Then
				    AssetPK = rs("AssetPK")
                    Call SetSession("ASSETPK" & CardCurrentLevel,assetpk)
		            Set rs = db.RunSQLReturnRS_RW("SELECT TOP 0 * FROM AssetHierarchy","")
		            If Not db.dok Then
			           Call OutputWAPError(db.derror)
		            End If
		            rs.AddNew()
		            rs("System") = "MC"
		            rs("AssetPK") = AssetPK
		            rs("ParentPK") = ParentPK
		            db.dobatchupdate rs

		            If Not db.dok Then
			           Call OutputWAPError(db.derror)
		            End If

				End If
			Else
				Call OutputWAPError(db.derror)
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub ASMeters()

	Dim db, sql, rs, ASE, assetid, assetname, photo, parentlocationall, parentequipmentall, newrec, IsLocation, rseof

    newcontextpage = True

	CardTitle = "Asset Meters"
	CardCurrent = "ASMeters"
	CardCurrentLevel = GetCardLevel()

	GetAssetPK

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	ASE = GetAccessRight(db,"ASE",0)

    If Not AssetPK = "-1" Then
	    sql = _
	    "SELECT Asset.*, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!') " &_
	    "FROM Asset WITH (NOLOCK) INNER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
	    "WHERE Asset.AssetPK = " & AssetPK & " "

	    Set rs = db.RunSQLReturnRS(sql,"")
	    Call CheckDB(db)

	    If Not rs.eof Then
		    assetpk = NullCheck(rs("AssetPK"))
		    assetid = WAPValidate(NullCheck(RS("AssetID")))
		    assetname = WAPValidate(NullCheck(RS("AssetName")))
		    photo = NullCheck(rs("Photo"))
            parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!","<br/>")
            parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!","<br/>")
            If Not parentlocationall = "" Then
                parentlocationall = parentlocationall & "<br/>"
            End If
            If Not parentequipmentall = "" Then
                parentequipmentall = parentequipmentall & "<br/>"
            End If
	    Else
		    Call OutputWAPError("The Asset was not found.")
	    End If

	    sql = _
	    "SELECT     Asset.* " &_
	    "FROM       Asset WITH (NOLOCK)  " &_
	    "WHERE     (Asset.AssetPK = " & AssetPK & " )"

	    Set rs = db.RunSQLReturnRS_RW(sql,"")
	    Call CheckDB(db)

        IsLocation = BitNullCheck(rs("IsLocation"))

        If rs.eof Then
            rseof = True
        Else
            rseof = False
        End If

	End If

	Call BuildFields("Meter1Reading","Meter 1 Reading","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Meter1Units","Meter 1 Units","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Meter2Reading","Meter 2 Reading","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Meter2Units","Meter 2 Units","C",GlobalFieldLength,"*M","true","",rs,AssetPK)

	Call StartMobileDocument(CardTitle)
		If rseof and Not AssetPK = "-1" Then
			Call OutputWAPMsg("The Asset was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<% If AssetPK = "-1" Then %>
		    <b><% =HStyleBegin %>New Asset<% =HStyleEnd %></b>
			<% Else %>
		    <b><% =HStyleBegin %><% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)<% =HStyleEnd %></b>
		    <% End If %>
			</p>
			<% End If
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			' We do not need to check for ASN because they would not have AssetPK = -1 if they could not get here from the main menu
			If IsOpen and (ASE or AssetPK = "-1") Then
			ASMetersSubmit rs
			End If
		Else
			If IsOpen and (ASE or AssetPK = "-1") Then
			ASMetersSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen and (ASE or AssetPK = "-1") Then
		%>
		</td><td align="right">
		<%
		ASMetersSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		Call CloseObj(rs)
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub ASMetersSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="ASOptions"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="Assetpk" value="<% =AssetPK %>"/>
			<postfield name="POSTEDASMETERS" value="Y"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="ASOptions"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="Assetpk" value="<% =AssetPK %>"/>
	<input type="hidden" name="POSTEDASMETERS" value="Y"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub ASMetersSave(ByRef db)

	'ASPDebug
	'Response.End

	Dim rs2, LaborPK, WorkDate, ParentPK

	If Not Request("POSTEDASMETERS") = "" Then
    	Assetpk = Request("Assetpk")
		sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetPK = " & Assetpk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If Not rs.Eof Then
			If UCase(Request("Delete")) = "Y" or UCase(Request("Delete")) = "2" Then
				rs.delete()
				Call SetSession("AssetPK" & CardCurrentLevel,"")
                Call SetSession("AssetPK" & CardCurrentLevel+1,"")
			Else
			    On Error Resume Next

                Call SaveField(db,"Meter1Reading","Meter 1 Reading","","","C","ASMeters",False)
                Call SaveField(db,"Meter1Units","Meter 1 Units","meterunits","","LT","ASMeters",False)
                Call SaveField(db,"Meter2Reading","Meter 2 Reading","","","C","ASMeters",False)
                Call SaveField(db,"Meter2Units","Meter 2 Units","meterunits","","LT","ASMeters",False)

				On Error Goto 0

			End If
			db.dobatchupdate rs
			If db.dok Then
			Else
				Call OutputWAPError(db.derror)
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub WOSearch()

	Dim db, sql, rs, wocount, wocount2, wocountu, TWC_WOAllOpen

	CardCurrent = "WOSearch"
	CardTitle = "WO Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

    TWC_WOAllOpen = GetAccessRight(db,"TWC_WOAllOpen",0)

	sql = _
	"SELECT WOCount = COUNT(DISTINCT WO.WOPK) " &_
	"FROM WO WITH (NOLOCK) " &_
	"INNER JOIN WOassign WITH (NOLOCK) ON WO.WOPK = WOassign.WOPK " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.Complete Is Null " &_
	"AND WOAssign.LaborPK = '" & GetSession("UserPK") & "' " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount = "0"
	Else
		wocount = rs("WOCount")
	End If

	sql = _
	"SELECT WOCount = COUNT(WO.WOPK) " &_
	"FROM WO WITH (NOLOCK) " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.Complete Is Null " &_
	"AND WO.RepairCenterPK = " & GetSession("RCPK") & " " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "

	sql = sql & AddGeneralFilters("WO")

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocount2 = "0"
	Else
		wocount2 = rs("WOCount")
	End If

	sql = _
	"SELECT WOCount = COUNT(WO.WOPK) " &_
	"FROM WO WITH (NOLOCK) " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.IsAssigned = 0 " &_
	"AND WO.Complete Is Null " &_
	"AND WO.RepairCenterPK = " & GetSession("RCPK") & " " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "

	sql = sql & AddGeneralFilters("WO")

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		wocountu = "0"
	Else
		wocountu = rs("WOCount")
	End If

	searchby = UCase(Request("wosearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=wosearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=1">WO #</a><br/>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=2">Reason</a><br/>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=3">Asset ID</a><br/>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=4">Asset Name</a><br/>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=5">Procedure ID</a><br/>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=6">Procedure Name</a><br/>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=7">Type</a><br/>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=8">Priority</a><br/>
		<a href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=9">Sub-Status</a><br/>
		<a href="default.asp?card=myworkorders&amp;s=<% =SessionID %>&amp;back=1">My WOs (<% =wocount %>)</a><br/>
		<% If TWC_WOAllOpen Then %>
		<a href="default.asp?card=allworkorders&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All WOs (<% =wocount2 %>)</a><br/>
		<a href="default.asp?card=allworkordersu&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Unassigned WOs (<% =wocountu %>)</a><br/>
		<% End If %>
		<% Case "1" %>
		<b>WO&nbsp;#: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Reason: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "3" %>
		<b>Asset ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "4" %>
		<b>Asset Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "5" %>
		<b>Procedure ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "6" %>
		<b>Procedure Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "7" %>
		<b>Type: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "8" %>
		<b>Priority: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "9" %>
		<b>Sub-Status: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CMSearch()

	Dim db, sql, rs

	CardCurrent = "CMSearch"
	CardTitle = "Company Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("cmsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=cmsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=cmsearch&amp;s=<% =SessionID %>&amp;cmsearchby=1">Company ID</a><br/>
		<a href="default.asp?card=cmsearch&amp;s=<% =SessionID %>&amp;cmsearchby=2">Company Name</a><br/>
		<a href="default.asp?card=cmlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Companies</a><br/>
		<% Case "1" %>
		<b>Company ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Company Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub FASearch()

	Dim db, sql, rs, desc, desc2

	CardCurrent = Card
	Select Case UCase(Card)
		Case "FAPRSEARCH"
			CardTitle = "Problem Search"
			desc = "Problems"
			desc2 = "Problem"
		Case "FAFASEARCH"
			CardTitle = "Failure Search"
			desc = "Failures"
			desc2 = "Failure"
		Case "FASOSEARCH"
			CardTitle = "Solution Search"
			desc = "Solutions"
			desc2 = "Solution"
	End Select

	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request(LCase(Card)&"by"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=<% =card %>&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=<% =card %>&amp;s=<% =SessionID %>&amp;<% =LCase(card) %>by=1"><% =desc2 %> ID</a><br/>
		<a href="default.asp?card=<% =card %>&amp;s=<% =SessionID %>&amp;<% =LCase(card) %>by=2"><% =desc2 %> Name</a><br/>
		<a href="default.asp?card=<% =Replace(UCase(card),"SEARCH","LOOKUP") %>&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All <% =desc %></a><br/>
		<% Case "1" %>
		<b><% =desc2 %> ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b><% =desc2 %> Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASSearch()

	Dim db, sql, rs

	CardCurrent = "ASSearch"
	CardTitle = "Asset Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("assearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=assearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=assearch&amp;s=<% =SessionID %>&amp;assearchby=1">Asset ID</a><br/>
		<a href="default.asp?card=assearch&amp;s=<% =SessionID %>&amp;assearchby=2">Asset Name</a><br/>
		<a href="default.asp?card=aslookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Assets</a><br/>
		<% Case "1" %>
		<b>Asset ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Asset Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASSearch2()

	Dim db, sql, rs

	CardCurrent = "ASSearch2"
	CardTitle = "Asset Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("assearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=assearch2&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=assearch2&amp;s=<% =SessionID %>&amp;assearchby=1">Asset ID</a><br/>
		<a href="default.asp?card=assearch2&amp;s=<% =SessionID %>&amp;assearchby=2">Asset Name</a><br/>
		<a href="default.asp?card=assets&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Assets</a><br/>
		<% Case "1" %>
		<b>Asset ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Asset Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CLSearch()

	Dim db, sql, rs

	CardCurrent = "CLSearch"
	CardTitle = "Classification Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("clsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=clsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=clsearch&amp;s=<% =SessionID %>&amp;clsearchby=1">Classification ID</a><br/>
		<a href="default.asp?card=clsearch&amp;s=<% =SessionID %>&amp;clsearchby=2">Classification Name</a><br/>
		<a href="default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Classifications</a><br/>
		<% Case "1" %>
		<b>Classification ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Classification Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ACSearch()

	Dim db, sql, rs

	CardCurrent = "ACSearch"
	CardTitle = "Account Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("acsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=acsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=acsearch&amp;s=<% =SessionID %>&amp;acsearchby=1">Account ID</a><br/>
		<a href="default.asp?card=acsearch&amp;s=<% =SessionID %>&amp;acsearchby=2">Account Name</a><br/>
		<a href="default.asp?card=aclookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Accounts</a><br/>
		<% Case "1" %>
		<b>Account ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Account Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub RCSearch()

	Dim db, sql, rs

	CardCurrent = "RCSearch"
	CardTitle = "Repair Center Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("rcsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=rcsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=rcsearch&amp;s=<% =SessionID %>&amp;rcsearchby=1">Repair Center ID</a><br/>
		<a href="default.asp?card=rcsearch&amp;s=<% =SessionID %>&amp;rcsearchby=2">Repair Center Name</a><br/>
		<a href="default.asp?card=rclookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Repair Centers</a><br/>
		<% Case "1" %>
		<b>Repair Center ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Repair Center Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub SHSearch()

	Dim db, sql, rs

	CardCurrent = "SHSearch"
	CardTitle = "Shop Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("shsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=shsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=shsearch&amp;s=<% =SessionID %>&amp;shsearchby=1">Shop ID</a><br/>
		<a href="default.asp?card=shsearch&amp;s=<% =SessionID %>&amp;shsearchby=2">Shop Name</a><br/>
		<a href="default.asp?card=shlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Shops</a><br/>
		<% Case "1" %>
		<b>Shop ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Shop Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CASearch()

	Dim db, sql, rs

	CardCurrent = "CASearch"
	CardTitle = "Category Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("casearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=casearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=casearch&amp;s=<% =SessionID %>&amp;casearchby=1">Category ID</a><br/>
		<a href="default.asp?card=casearch&amp;s=<% =SessionID %>&amp;casearchby=2">Category Name</a><br/>
		<a href="default.asp?card=calookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Categories</a><br/>
		<% Case "1" %>
		<b>Category ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Category Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ZNSearch()

	Dim db, sql, rs

	CardCurrent = "ZNSearch"
	CardTitle = "Zone Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("znsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=znsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=znsearch&amp;s=<% =SessionID %>&amp;znsearchby=1">Zone ID</a><br/>
		<a href="default.asp?card=znsearch&amp;s=<% =SessionID %>&amp;znsearchby=2">Zone Name</a><br/>
		<a href="default.asp?card=znlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Zones</a><br/>
		<% Case "1" %>
		<b>Zone ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Zone Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub PRSearch()

	Dim db, sql, rs

	CardCurrent = "PRSearch"
	CardTitle = "Procedure Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("prsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=prsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=prsearch&amp;s=<% =SessionID %>&amp;prsearchby=1">Procedure ID</a><br/>
		<a href="default.asp?card=prsearch&amp;s=<% =SessionID %>&amp;prsearchby=2">Procedure Name</a><br/>
		<a href="default.asp?card=prlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Procedures</a><br/>
		<% Case "1" %>
		<b>Procedure ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Procedure Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub DPSearch()

	Dim db, sql, rs

	CardCurrent = "DPSearch"
	CardTitle = "Department Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("dpsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=dpsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=dpsearch&amp;s=<% =SessionID %>&amp;dpsearchby=1">Department ID</a><br/>
		<a href="default.asp?card=dpsearch&amp;s=<% =SessionID %>&amp;dpsearchby=2">Department Name</a><br/>
		<a href="default.asp?card=dplookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Departments</a><br/>
		<% Case "1" %>
		<b>Department ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Department Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub TNSearch()

	Dim db, sql, rs

	CardCurrent = "TNSearch"
	CardTitle = "Customer Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("tnsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=tnsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=tnsearch&amp;s=<% =SessionID %>&amp;tnsearchby=1">Customer ID</a><br/>
		<a href="default.asp?card=tnsearch&amp;s=<% =SessionID %>&amp;tnsearchby=2">Customer Name</a><br/>
		<a href="default.asp?card=tnlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Customers</a><br/>
		<% Case "1" %>
		<b>Customer ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Customer Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub PJSearch()

	Dim db, sql, rs

	CardCurrent = "PJSearch"
	CardTitle = "Project Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("pjsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=pjsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=pjsearch&amp;s=<% =SessionID %>&amp;pjsearchby=1">Project ID</a><br/>
		<a href="default.asp?card=pjsearch&amp;s=<% =SessionID %>&amp;pjsearchby=2">Project Name</a><br/>
		<a href="default.asp?card=pjlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Projects</a><br/>
		<% Case "1" %>
		<b>Project ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Project Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub LASearch()

	Dim db, sql, rs

	CardCurrent = "LASearch"
	CardTitle = "Labor Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("lasearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=lasearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=lasearch&amp;s=<% =SessionID %>&amp;lasearchby=1">Labor ID</a><br/>
		<a href="default.asp?card=lasearch&amp;s=<% =SessionID %>&amp;lasearchby=2">Labor Name</a><br/>
		<a href="default.asp?card=lalookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Labor</a><br/>
		<% Case "1" %>
		<b>Labor ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Labor Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub INSearch()

	Dim db, sql, rs

	CardCurrent = "INSearch"
	CardTitle = "Item Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("insearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=insearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=insearch&amp;s=<% =SessionID %>&amp;insearchby=1">Part ID</a><br/>
		<a href="default.asp?card=insearch&amp;s=<% =SessionID %>&amp;insearchby=2">Part Name</a><br/>
		<a href="default.asp?card=insearch&amp;s=<% =SessionID %>&amp;insearchby=3">Part Desc</a><br/>
		<a href="default.asp?card=insearch&amp;s=<% =SessionID %>&amp;insearchby=4">Vendor Part #</a><br/>
		<a href="default.asp?card=inlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Items</a><br/>
		<% Case "1" %>
		<b>Item ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Item Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "3" %>
		<b>Item Desc: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "4" %>
		<b>Vendor Part #: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub SRSearch()

	Dim db, sql, rs

	CardCurrent = "SRSearch"
	CardTitle = "Location Search"
	If CardFrom = CardCurrent and SearchBy = "" Then
		CardCurrentLevel = CardFromLevel
	Else
		CardCurrentLevel = CardFromLevel + 1
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	searchby = UCase(Request("srsearchby"))
	Call StartMobileDocument(CardTitle)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Search By:</b>
		<% Else %>
		<b><a href="default.asp?card=srsearch&amp;s=<% =SessionID %>">Search By:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=srsearch&amp;s=<% =SessionID %>&amp;srsearchby=1">Location ID</a><br/>
		<a href="default.asp?card=srsearch&amp;s=<% =SessionID %>&amp;srsearchby=2">Location Name</a><br/>
		<a href="default.asp?card=srlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">All Locations</a><br/>
		<a href="default.asp?card=srlookup&amp;s=<% =SessionID %>&amp;searchby=3">All Locations (All Repair Centers)</a><br/>
		<% Case "1" %>
		<b>Location ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Location Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "3" %>
		<b>Click Submit to View All Stock Rooms at All RCs</b>
		<% End Select %>
		</p>
		<%
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If Not SearchBy = "" Then
			SearchSubmit
			End If
		Else
			If Not SearchBy = "" Then
			SearchSubmit
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
		SearchSubmit
		End If
		HTMLbuttonsEnd
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub SearchSubmit()
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="post">
			<postfield name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="searchby" value="<% =searchby %>"/>
			<postfield name="searchvalue" value="$(searchvalue<% =r %>)"/>
			<postfield name="back" value="1"/>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="searchby" value="<% =searchby %>"/>
	<input type="hidden" name="back" value="1"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub ASPhoto()

	Dim db, sql, rs, woid, assetid, assetname, reason, photo, FromWO

	CardTitle = "Asset Photo"
	CardCurrent = "ASPhoto"
	CardCurrentLevel = GetCardLevel()

	' Either WOPK or AssetPK can be passed in so do not use GetWOPK or GetAssetPK functions here
	wopk = Request("WOPK")
	assetpk = Request("AssetPK")

	If assetpk = "" Then
		FromWO = True
	Else
		FromWO = False
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	If FromWO Then
		sql = _
		"SELECT WO.*, Asset.Photo as AssetPhoto " &_
		"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
		"WHERE WOPK = " & WOPK & " "
	Else
		sql = _
		"SELECT Asset.* " &_
		"FROM ASSET WITH (NOLOCK) " &_
		"WHERE AssetPK = " & AssetPK & " "
	End If

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		Call OutputWAPError("Asset Photo Not Found.")
	Else
		If FromWO Then
			woid = NullCheck(rs("WOID"))
			reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
			photo = NullCheck(rs("AssetPhoto"))
		Else
			photo = NullCheck(rs("Photo"))
		End If
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
	End If

	Call StartMobileDocument(CardTitle)
		If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap">
		<% If FromWO Then %>
		<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">WO #<% =WOID %></a></b><br/>
		<% If Not AssetPK = "" Then %>
		<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
		<% End If %>
		<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
		<% Else %>
		<b><% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %></b><br/>
		<% End If %>
		</p>
		<% End If %>
		<p mode="nowrap">
		<img<% If lang="HTML" Then %> border="0"<% End If %> src="<% =Application("ImageServer") & Photo %>" alt="Asset Photo"/><br/>
		<% =HStyleBegin %><% =WAPValidate(NullCheck(RS("AssetName"))) %> (<% =WAPValidate(NullCheck(RS("AssetID"))) %>)<% =HStyleEnd %>
		</p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOTasksOld()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo
	Dim CheckBoxType

	CardTitle = "WO Tasks"
	CardCurrent = "WOTasks"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

	pagesize = GlobalPageSize

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	If IsPPC Then
		CheckBoxType = "GRAPHIC"
	ElseIf IsSmartPhone or IsBlackBerry Then
		CheckBoxType = "TEXT"
	Else
		CheckBoxType = "NONE"
	End If

	Set db = New ADOHelper

	Call WOTaskSave(db)

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
	End If

	sql = _
	"SELECT * " &_
	"FROM WOTask WITH (NOLOCK) " &_
	"WHERE WOPK = " & WOPK & " " &_
	"ORDER BY TaskNo "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	Call StartMobileDocument(CardTitle)
		If rs.eof Then
			Call OutputWAPMsg("No Tasks Found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">WO #<% =WOID %></a></b><br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
			</p>
			<% End If %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				If BitNullCheck(rs("Header")) Then
				%>
				<b><% =LStyleBegin %><% =WAPValidate(Shorten(Replace(NullCheck(rs("TaskAction")),"%0D%0A",CHR(13) & CHR(10)),100)) %><% =LStyleEnd %></b><br/>
				<%
				Else
				If CheckBoxType = "GRAPHIC" Then
				If BitNullCheck(rs("Complete")) Then %>
				<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/true.gif"/></a> <% Else %>
				<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/false.gif"/></a> <% End If
				ElseIf CheckBoxType = "TEXT" Then
				If BitNullCheck(rs("Complete")) Then %>
				<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>">x </a><% Else %>
				<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>">_ </a><% End If
				Else
				If BitNullCheck(rs("Complete")) Then %><% =LStyleBegin %><u>x</u><% =LStyleEnd %>&nbsp;<% Else %>_&nbsp;<% End If %>
				<% End If %><% =LStyleBegin %><a href="default.asp?card=wotask&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>"><% =rs("TaskNo") %>: <% =WAPValidate(Shorten(Replace(NullCheck(rs("TaskAction")),"%0D%0A","&nbsp;"),100)) %></a><% =LStyleEnd %><br/>
				<%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOTaskSave(ByRef db)

	pk = Request("PK")
	If Not pk = "" Then
		sql = "SELECT * FROM WOTask WITH (NOLOCK) WHERE PK = " & pk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If Not rs.Eof Then
			If Not Request("flipit") = "" Then
				If BitNullCheck(rs("Complete")) Then
					rs("Complete") = False
				Else
					rs("Complete") = True
				End If
			Else
			    On Error Resume Next
				If Not NullCheck(Request("HoursActual")) = "" Then
					rs("HoursActual") =	NullCheck(Request("HoursActual"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "The value provided for Hours Actual is invalid."
					WOTask
				End If
				If Not NullCheck(Request("Measurement")) = "" Then
					rs("Measurement") =	NullCheck(Request("Measurement"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "The value provided for Measurement is invalid."
					WOTask
				End If
				If Not NullCheck(Request("Rate")) = "" Then
					rs("Rate") = NullCheck(Request("Rate"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "The value provided for Rate is invalid."
					WOTask
				End If
				If Not NullCheck(Request("Comments")) = "" Then
					rs("Comments") =	NullCheck(Request("Comments"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "The value provided for Comments is invalid."
					WOTask
				End If
				If UCase(NullCheck(Request("Fail"))) = "Y" or UCase(NullCheck(Request("Fail"))) = "2" Then
					rs("Fail") = True
				Else
					rs("Fail") = False
				End If
				If UCase(NullCheck(Request("Complete"))) = "Y" or UCase(NullCheck(Request("Complete"))) = "2" Then
					rs("Complete") = True
				Else
					rs("Complete") = False
				End If
				On Error Goto 0
			End If
			db.dobatchupdate rs
			If db.dok Then
				'pl = rs("pk")
			Else
				Call OutputWAPError(db.derror)
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub WOTask()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, FromAsset

    newcontextpage = True

	CardTitle = "WO Task"
	CardCurrent = "WOTask"
	CardCurrentLevel = GetCardLevel()

	If Not Trim(UCase(GetSession("ParentCard2"))) = "ASSETTASKS" Then
		GetWOPK
		FromAsset = False
	Else
		FromAsset = True
	End If

	pk = Request("PK")
	If pk = "" Then
		pk = GetSession("WOTASKPK" & CardCurrentLevel)
	Else
		Call SetSession("WOTASKPK" & CardCurrentLevel,pk)
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	If Not FromAsset Then
		sql = _
		"SELECT WO.*, Asset.Photo as AssetPhoto " &_
		"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
		"WHERE WOPK = " & WOPK & " "

		Set rs = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)

		If rs.eof Then
			woid = "Unknown"
			reason = ""
			photo = ""
			assetpk = ""
			assetid = ""
			assetname = ""
			isopen = True
		Else
			woid = NullCheck(rs("WOID"))
			reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
			photo = NullCheck(rs("AssetPhoto"))
			assetpk = NullCheck(rs("AssetPK"))
			assetid = WAPValidate(NullCheck(RS("AssetID")))
			assetname = WAPValidate(NullCheck(RS("AssetName")))
			isopen = BitNullCheck(rs("isopen"))
		End If

	Else
			woid = ""
			reason = ""
			photo = ""
			assetpk = ""
			assetid = ""
			assetname = ""
			isopen = True

	End If

	sql = _
	"SELECT * " &_
	"FROM WOTask WITH (NOLOCK) " &_
	"WHERE PK = " & PK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	Call BuildFields("Complete","Complete?","B","","","",True,rs,PK)
	Call BuildFields("Fail","Failed?","B","","","",False,rs,PK)
	Call BuildFields("HoursActual","Hour(s) (Actual)","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("Measurement","Measurement","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("Rate","Rating","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("Comments","Comments","C",GlobalFieldLength,"*M","true","",rs,PK)

	Call StartMobileDocument(CardTitle)
		If rs.eof Then
			Call OutputWAPMsg("The Task was not found")
		Else %>
			<% If (Not FromAsset) and (IsPocketIE or IsBlackBerry) Then %>
			<p align="center" mode="wrap">
			<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2">WO #<% =WOID %></a></b><br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
			</p>
			<% End If %>
			<p mode="wrap"><% =HStyleBegin %><b>
			<% If IsPocketIE Then %>
			<% =NullCheck(rs("TaskNo")) & ": " %><% =Replace(NullCheck(rs("TaskAction")),"%0D%0A",CHR(13) & CHR(10)) %><br/>
			<% Else %>
			<% =NullCheck(rs("TaskNo")) & ": " %><% =Shorten(Replace(NullCheck(rs("TaskAction")),"%0D%0A",CHR(13) & CHR(10)),500) %><br/>
			<% End If %>
			</b><% =HStyleEnd %></p><%
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If IsOpen Then
			WOTaskSubmit rs
			End If
		Else
			If IsOpen Then
			WOTaskSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen Then
		%>
		</td><td align="right">
		<%
		WOTaskSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WOTaskSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="post">
			<postfield name="card" value="WOTasks"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="wopk" value="<% =WOPK %>"/>
			<postfield name="pk" value="<% =PK %>"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="WOTasks"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wopk" value="<% =WOPK %>"/>
	<input type="hidden" name="pk" value="<% =PK %>"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub WOTasks()

	'Call ASPDebug
	'Response.End

	Dim db, sql, rs, woid, reason, assetid, assetname, photo
	Dim CheckBoxType, FromAsset

	CardTitle = "WO Tasks"
	CardCurrent = "WOTasks"
	CardCurrentLevel = GetCardLevel()

	If Not Trim(UCase(GetSession("ParentCard2"))) = "ASSETTASKS" Then
		GetWOPK
		FromAsset = False
	Else
		assetpk = GetSession("ASSETPK2")
		FromAsset = True
	End If

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = 7
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = 5
    Else
	    pagesize = GlobalPageSize
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call WOTaskSave(db)

	If FromAsset Then
		sql = _
		"SELECT Asset.*, Asset.Photo as AssetPhoto " &_
		"FROM Asset WITH (NOLOCK) " &_
		"WHERE AssetPK = " & AssetPK & " "
	Else
		sql = _
		"SELECT WO.*, Asset.Photo as AssetPhoto " &_
		"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
		"WHERE WOPK = " & WOPK & " "
	End If

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If FromAsset Then
		If rs.eof Then
			woid = "Unknown"
			reason = ""
			photo = ""
			assetpk = ""
			assetid = ""
			assetname = ""
			isopen = True
		Else
			woid = ""
			reason = ""
			photo = NullCheck(rs("AssetPhoto"))
			assetpk = NullCheck(rs("AssetPK"))
			assetid = WAPValidate(NullCheck(RS("AssetID")))
			assetname = WAPValidate(NullCheck(RS("AssetName")))
			isopen = True
		End If
	Else
		If rs.eof Then
			woid = "Unknown"
			reason = ""
			photo = ""
			assetpk = ""
			assetid = ""
			assetname = ""
			isopen = True
		Else
			woid = NullCheck(rs("WOID"))
			reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
			photo = NullCheck(rs("AssetPhoto"))
			assetpk = NullCheck(rs("AssetPK"))
			assetid = WAPValidate(NullCheck(RS("AssetID")))
			assetname = WAPValidate(NullCheck(RS("AssetName")))
			isopen = BitNullCheck(rs("isopen"))
		End If
	End If

	If IsPPC Then
		CheckBoxType = "GRAPHIC"
	ElseIf IsSmartPhone or IsBlackBerry Then
		CheckBoxType = "TEXT"
	Else
		CheckBoxType = "NONE"
	End If

	If Not IsOpen Then
		CheckBoxType = "NONE"
	End IF

	If FromAsset Then
	sql = _
	"SELECT     WOtask.PK, WOtask.TaskNo, WOtask.TaskAction, " &_
	"TaskAction2 = " &_
	"CASE " &_
	"WHEN WOTask.AssetPK Is Not Null Then '<b>' + Asset.AssetName + ' [' + Asset.AssetID + ']</b> ' + CASE WHEN Asset.Vicinity Is Not Null AND Asset.Vicinity <> '' Then RTrim(Asset.Vicinity) + ': ' Else '' END " &_
	"Else Null " &_
	"END, WOtask.Rate, WOtask.Measurement, WOtask.Initials, WOtask.Fail, WOtask.Complete, WOtask.Header, WOtask.Spec, WOtask.LineStyle, WOtask.LineStyleDesc, WOtask.RowVersionDate, Task.TrackTask, AssetSpecification.ValueHi, AssetSpecification.ValueLow, " &_
	"Labor.LaborPK, Labor.LaborID, Labor.LaborName, Tool.ToolPK, Tool.ToolID, Tool.ToolName, WOtask.HoursEstimated, WOtask.HoursActual, WOtask.MeasurementInitial, WOtask.Meter1, WOtask.Meter2, WOtask.AssetPK, WO.WOID " &_
	"FROM         WOtask WITH (NOLOCK) LEFT OUTER JOIN Task WITH (NOLOCK) ON Task.TaskPK = WOTask.TaskPK " &_
	"					INNER JOIN WO WITH (NOLOCK) ON WO.WOPK = WOTask.WOPK " &_
	"				    LEFT OUTER JOIN AssetSpecification WITH (NOLOCK) ON AssetSpecification.PK = WOTask.AssetSpecificationPK " &_
	"				    LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WOTask.AssetPK " &_
	"				    LEFT OUTER JOIN Labor WITH (NOLOCK) ON Labor.LaborPK = WOTask.CraftPK " &_
	"				    LEFT OUTER JOIN Tool WITH (NOLOCK) ON Tool.ToolPK = WOTask.ToolPK " &_
	"WHERE WO.IsOpen = 1 AND (WO.AssetPK = " & AssetPK & " OR WOTask.AssetPK = " & AssetPK & ") " &_
	"ORDER BY WO.WOPK DESC, WOTask.TaskNo "
	Else
	sql = _
	"SELECT     WOtask.PK, WOtask.TaskNo, WOtask.TaskAction, " &_
	"TaskAction2 = " &_
	"CASE " &_
	"WHEN WOTask.AssetPK Is Not Null Then '<b>' + Asset.AssetName + ' [' + Asset.AssetID + ']</b> ' + CASE WHEN Asset.Vicinity Is Not Null AND Asset.Vicinity <> '' Then RTrim(Asset.Vicinity) + ': ' Else '' END " &_
	"Else Null " &_
	"END, WOtask.Rate, WOtask.Measurement, WOtask.Initials, WOtask.Fail, WOtask.Complete, WOtask.Header, WOtask.Spec, WOtask.LineStyle, WOtask.LineStyleDesc, WOtask.RowVersionDate, Task.TrackTask, AssetSpecification.ValueHi, AssetSpecification.ValueLow, " &_
	"Labor.LaborPK, Labor.LaborID, Labor.LaborName, Tool.ToolPK, Tool.ToolID, Tool.ToolName, WOtask.HoursEstimated, WOtask.HoursActual, WOtask.MeasurementInitial, WOtask.Meter1, WOtask.Meter2, WOtask.AssetPK " &_
	"FROM         WOtask WITH (NOLOCK) LEFT OUTER JOIN Task WITH (NOLOCK) ON Task.TaskPK = WOTask.TaskPK " &_
	"				    LEFT OUTER JOIN AssetSpecification WITH (NOLOCK) ON AssetSpecification.PK = WOTask.AssetSpecificationPK " &_
	"				    LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WOTask.AssetPK " &_
	"				    LEFT OUTER JOIN Labor WITH (NOLOCK) ON Labor.LaborPK = WOTask.CraftPK " &_
	"				    LEFT OUTER JOIN Tool WITH (NOLOCK) ON Tool.ToolPK = WOTask.ToolPK " &_
	"WHERE WOPK = " & WOPK & " " &_
	"ORDER BY TaskNo "
	End If

	'sql = _
	'"SELECT * " &_
	'"FROM WOTask WITH (NOLOCK) " &_
	'"WHERE WOPK = " & WOPK & " " &_
	'"ORDER BY TaskNo "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	Call StartMobileDocument(CardTitle)
		If rs.eof and False Then
			Call OutputWAPMsg("No Tasks Found")
		Else
		%>
			<% If Not FromAsset Then %>
				<% If IsPocketIE or IsBlackBerry Then %>
					<p align="center" mode="wrap"><%
					If Not IsBOF Then %>
					<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;pr=<% =WOPK %>">&lt;</a>&nbsp;<%
					End If %><b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">WO #<% =WOID %></a></b><%
					If Not IsEOF Then %>
					<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;nr=<% =WOPK %>">&gt;</a><%
					End If %>
					<br/>
				<% End If %>
				<% If Not AssetPK = "" Then %>
				<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
				<% End If %>
				<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
		    <% End If
		    If lang = "WML" and IsPPC Then
				If Not FromAsset Then
				Response.Write "<br/>"
				Response.Write HStyleBegin
				OutputBackButton
				OutputHomeButton
				Response.Write HStyleEnd
				End If
		    End If
			%>
			<% If Not FromAsset Then %>
			<% If IsPocketIE or IsBlackBerry Then %>
		    </p>
			<%
			End If
			End If %>
			<p mode="nowrap">
			<%

			Dim borderzero, wopkc
			wopkc="!"
			borderzero = ""
			If lang = "HTML" Then
			    borderzero="border=""0"" "
			End If
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				If FromAsset Then
					If Not wopkc = Trim(RS("WOID")) Then
						wopkc = Trim(RS("WOID"))
						If False and (IsPocketIE or IsBlackBerry) Then %>
						<p align="center" mode="wrap"><%
						End If %>
						WO #<% =wopkc %><br/><%
						If False and (IsPocketIE or IsBlackBerry) Then
						Response.Write "</p>"
						End If
					End If
				End If

				Dim Fail,Complete,LineTemplate,LineTemplate2,SpecHiOK,SpecLowOK, TaskText, TaskText2, TaskText3, SpecText, TaskTextFinal
				TaskText = JSEncode(RS("TaskAction"))
				TaskText2 = JSEncode(RS("TaskAction2"))
				TaskText3 = ""
				If Not TaskText2 = "" Then
					TaskText2 = TaskText2 & "<br/>"
				End If
				SpecText = ""
				SpecHiOK = False
				SpecLowOK = False
				If BitNullCheck(RS("Header")) Then %>
				    <% If lang = "HTML" Then %><nobr><% End If %><b><% =LStyleBegin %><% =WAPValidate(Shorten(Replace(NullCheck(rs("TaskAction")),"%0D%0A",CHR(13) & CHR(10)),100)) %><% =LStyleEnd %></b><% If lang = "HTML" Then %></nobr><% End If %><%
				Else
					If RS("Fail") Then
						Fail = "<img " & borderzero & "src=""images/failed.gif""/>"
					Else
						Fail = "<img " & borderzero & "src=""images/false.gif""/>"
					End If
					If RS("Complete") Then
						Complete = "<img " & borderzero & "src=""images/true.gif""/>"
					Else
						Complete = "<img " & borderzero & "src=""images/false.gif""/>"
					End If
					Select Case NullCheck(RS("LineStyle"))
						Case "I1"
							LineTemplate=7
						Case "I2"
							LineTemplate=8
						Case "I3"
							LineTemplate=9
						Case "B"
							LineTemplate=10
						Case "BR"
							LineTemplate=11
						Case "I"
							LineTemplate=12
						Case Else
							LineTemplate=4
					End Select
					LineTemplate2="0"
					If rs("Spec") and Not NullCheck(rs("ValueLow")) = "" and Not NullCheck(rs("Measurement")) = "" Then
						If IsNumeric(rs("Measurement")) and IsNumeric(rs("ValueLow")) Then
							If CLng(rs("Measurement")) < CLng(rs("ValueLow")) Then
								TaskText3 = " (Below minimum value of " & rs("ValueLow") & ")"
								LineTemplate2 = 11
							Else
								SpecLowOK = True
								SpecText = SpecText & "Minimum Value: " & rs("ValueLow") & " "
							End If
						End If
					End If
					If rs("Spec") and Not NullCheck(rs("ValueHi")) = "" and Not NullCheck(rs("Measurement")) = "" Then
						If IsNumeric(rs("Measurement")) and IsNumeric(rs("ValueHi")) Then
							If CLng(rs("Measurement")) > CLng(rs("ValueHi")) Then
								LineTemplate2 = 11
								TaskText3 = " (Above maximum value of " & rs("ValueHi") & ")"
							Else
								SpecHiOK = True
								If SpecText = "" Then
									SpecText = SpecText & "Maximum Value: " & rs("ValueLow")
								Else
									SpecText = "Range: " & rs("ValueLow") & " - " & rs("ValueHi")
								End If
							End If
						End If
					End If
					If rs("Spec") and (SpecLowOK or SpecHiOK) and Not LineTemplate2 = 11 Then
						LineTemplate2 = 14
						SpecText = " (" & SpecText & ")"
						TaskText3 = SpecText
					ElseIf RS("Spec") and Not LineTemplate2 = 11 and Not LineTemplate2 = 14 Then
						If IsNumeric(rs("ValueLow")) and IsNumeric(rs("ValueHi")) Then
							SpecText = " (Range: " & rs("ValueLow") & " - " & rs("ValueHi") & ")"
							TaskText3 = SpecText
						End If
						LineTemplate2 = 14
					End If
					If rs("TrackTask") Then
						If LineTemplate2 = 11 Then
							TaskTextFinal = "<img " & borderzero & "src=""images/tasks_g.gif""/> " & TaskText2 & TaskText & TaskText3
						ElseIf LineTemplate2 = 14 Then
							TaskTextFinal = "<img " & borderzero & "src=""images/tasks_g.gif""/> " & TaskText2 & TaskText & TaskText3
						ElseIf RS("Meter1") or RS("Meter2") Then
							If RS("Meter1") Then
								TaskTextFinal = "<img " & borderzero & "src=""images/tasks_g.gif""/> " & TaskText2 & TaskText
							Else
								TaskTextFinal = "<img " & borderzero & "src=""images/tasks_g.gif""/> " & TaskText2 & TaskText
							End If
						Else
							TaskTextFinal = "<img " & borderzero & "src=""images/tasks_g.gif""/> " & TaskText2 & TaskText
						End If
					Else
						If LineTemplate2 = 11 Then
							TaskTextFinal = TaskText2 & TaskText & TaskText3
						ElseIf LineTemplate2 = 14 Then
							TaskTextFinal = TaskText2 & TaskText & TaskText3
						ElseIf RS("Meter1") or RS("Meter2") Then
							If RS("Meter1") Then
								TaskTextFinal = TaskText2 & TaskText
							Else
								TaskTextFinal = TaskText2 & TaskText
							End IF
						Else
							TaskTextFinal = TaskText2 & TaskText
						End If
					End If
					If CheckBoxType = "GRAPHIC" Then
					If BitNullCheck(rs("Complete")) Then %>
					<% If lang = "HTML" Then %><nobr><% End If %><a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/true.gif"/></a> <% Else %>
					<% If lang = "HTML" Then %><nobr><% End If %><a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/false.gif"/></a> <% End If
					ElseIf CheckBoxType = "TEXT" Then
					If BitNullCheck(rs("Complete")) Then %>
					<% If lang = "HTML" Then %><nobr><% End If %><a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>">x </a><% Else %>
					<% If lang = "HTML" Then %><nobr><% End If %><a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>">_ </a><% End If
					Else
					If BitNullCheck(rs("Complete")) Then %><% If lang = "HTML" Then %><nobr><% End If %><% =LStyleBegin %><u>x</u><% =LStyleEnd %>&nbsp;<% Else %><% If lang = "HTML" Then %><nobr><% End If %>_&nbsp;<% End If %>
					<% End If %><% =LStyleBegin %><a href="default.asp?card=wotask&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>"><% =rs("TaskNo") %>: <% =WAPValidate(Shorten(Replace(NullCheck(TaskTextFinal),"%0D%0A","&nbsp;"),100)) %></a><% =LStyleEnd %><% If lang = "HTML" Then %></nobr><% End If %><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
			        Response.Write "<br/>"
			    End If
			End If
			Loop
			%>
			</p><%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOLaborRec()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec

    newcontextpage = True

	CardTitle = "WO Labor"
	CardCurrent = "WOLaborRec"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

	pk = Request("PK")
	If pk = "" Then
		pk = GetSession("WOLABORPK" & CardCurrentLevel)
		If pk = "" Then
		    pk = "-1"
		End If
	Else
		Call SetSession("WOLABORPK" & CardCurrentLevel,pk)
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
		isopen = True
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
	End If

	sql = _
	"SELECT     WOlabor.PK, WOlabor.LaborPK, lt.moduleid, WOlabor.LaborName, WOlabor.EstimatedHours, WOlabor.TotalHours, WOlabor.RegularHours, WOlabor.OvertimeHours, WOlabor.OtherHours, WOlabor.WorkDate, WOlabor.TimeIn, " &_
	"      WOlabor.TimeOut, WOlabor.AccountID, WOlabor.AccountName, WOlabor.CategoryID, WOlabor.CategoryName, WOlabor.TotalCost, " &_
	"                      WOlabor.TotalCharge, WOlabor.CostRegular, WOlabor.CostOvertime, WOlabor.CostOther, WOlabor.ChargeRate, WOlabor.ChargePercentage, WOlabor.RowVersionDate, Labor.Photo, WAS.Completed, WOlabor.Comments " &_
	"FROM         WOlabor WITH (NOLOCK)  LEFT OUTER JOIN WOAssignStatus WAS WITH (NOLOCK) ON WAS.WOPK = WOlabor.WOPK AND WAS.LaborPK = WOlabor.LaborPK " &_
	"                      LEFT OUTER JOIN Labor WITH (NOLOCK) ON WOlabor.LaborPK = Labor.LaborPK INNER JOIN " &_
	"                      LaborTypes lt WITH (NOLOCK) ON lt.LaborType = Labor.LaborType " &_
	"WHERE     (WOlabor.PK = " & PK & ") AND (WOlabor.RecordType = 2) " &_
	"ORDER BY WOlabor.LaborType, WOlabor.WorkDate, WOlabor.LaborName "

	Set rs = db.RunSQLReturnRS_RW(sql,"")
	Call CheckDB(db)

	If PK = "-1" Then
		Call BuildFields("LaborID","Labor ID","C",GlobalFieldLength,"*M","true",GetSession("UserID"),rs,PK)
	End If
	Call BuildFields("RegularHours","Reg Hrs","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("OvertimeHours","OT Hrs","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("OtherHours","Other Hrs","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("WorkDate","Date","C",GlobalFieldLength,"*M","true",CStr(DateNullCheck(Date())),rs,PK)
	If IsPocketIE or IsBlackBerry Then
		Call BuildFields("TimeIn","Time In","C",GlobalFieldLength,"*M","true","",rs,PK)
		Call BuildFields("TimeOut","Time Out","C",GlobalFieldLength,"*M","true","",rs,PK)
		Call BuildFields("AccountID","Account ID","C",GlobalFieldLength,"*M","true","",rs,PK)
		Call BuildFields("CategoryID","Category ID","C",GlobalFieldLength,"*M","true","",rs,PK)
	End If
	Call BuildFields("Comments","Comments","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("Completed","All Complete?","B","","","",True,rs,PK)
	If Not PK = "-1" and IsOpen Then
		Call BuildFields("Delete","Delete Record?","B","","","",False,rs,PK)
	End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("The Labor was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2">WO #<% =WOID %></a></b><br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
			</p>
			<% End If %>
			<p mode="wrap"><% =HStyleBegin %><b>
			<% If PK = "-1" Then %>
			New Labor<%
			Else %>
			<% =WAPValidate(NullCheck(rs("LaborName"))) %>: [<% =NullCheck(RS("EstimatedHours")) & " Est] [" & NullCheck(RS("TotalHours")) & " Actual" & "] " & DateNullCheck(RS("WorkDate")) %>
			<% End If %>
			</b><% =HStyleEnd %></p><%
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If IsOpen Then
			WOLaborRecSubmit rs
			End If
		Else
			If IsOpen Then
			WOLaborRecSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen Then
		%>
		</td><td align="right">
		<%
		WOLaborRecSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WOLaborRecSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="WOLabor"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="wopk" value="<% =WOPK %>"/>
			<postfield name="pk" value="<% =PK %>"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="WOLabor"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wopk" value="<% =WOPK %>"/>
	<input type="hidden" name="pk" value="<% =PK %>"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub WOLaborRecSave(ByRef db, FromClose, LaborID)

	'ASPDebug
	'Response.End

	Dim rs2, LaborPK, WorkDate

	If FromClose Then
		pk = "-1"
		WorkDate = NullCheck(Request("Date"))
	Else
		pk = Request("pk")
		LaborID = NullCheck(Request("LaborID"))
		WorkDate = NullCheck(Request("WorkDate"))
	End If

	If Not pk = "" Then
		sql = "SELECT * FROM WOLabor WITH (NOLOCK) WHERE PK = " & pk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If pk = "-1" Then
			rs.addnew()
		End If
		If Not rs.Eof Then
			If UCase(Request("Delete")) = "Y" or UCase(Request("Delete")) = "2" Then
				rs.delete()
				Call SetSession("WOLaborPK" & CardCurrentLevel,"")
                Call SetSession("WOLaborPK" & CardCurrentLevel+1,"")
			Else
				If pk = "-1" Then
					rs("RecordType") = 2
					rs("WOPK") = WOPK
					sql = "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborID = '" & LaborID & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Labor ID was not found."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					Else
						If pk = "-1" Then
							rs("LaborPK") = rs2("LaborPK")
							rs("LaborID") = rs2("LaborID")
							rs("LaborName") = rs2("LaborName")
							rs("CostRegular") = rs2("CostRegular")
							rs("CostOvertime") = rs2("CostOvertime")
							rs("CostOther") = rs2("CostOther")
							rs("ChargePercentage") = rs2("ChargePercentage")
							rs("ChargeRate") = rs2("ChargeRate")
						End If
						LaborPK = rs2("LaborPK")
					End If
				Else
					LaborPK = rs("LaborPK")
				End If
				On Error Resume Next
				If Not NullCheck(Request("RegularHours")) = "" Then
					rs("RegularHours") = NullCheck(Request("RegularHours"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Regular Hours is invalid."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					End If
				End If
				If Not NullCheck(Request("OvertimeHours")) = "" Then
					rs("OvertimeHours") =	NullCheck(Request("OvertimeHours"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Overtime Hours is invalid."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					End If
				End If
				If Not NullCheck(Request("OtherHours")) = "" Then
					rs("OtherHours") =	NullCheck(Request("OtherHours"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Other Hours is invalid."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					End If
				End If
				If Not WorkDate = "" Then
					rs("WorkDate") = WorkDate
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Work Date is invalid."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					End If
				End If
				If Not NullCheck(Request("TimeIn")) = "" Then
					rs("TimeIn") =	NullCheck(Request("TimeIn"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Time In is invalid."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					End If
				End If
				If Not NullCheck(Request("TimeOut")) = "" Then
					rs("TimeOut") =	NullCheck(Request("TimeOut"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Time Out is invalid."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					End If
				End If
				If Not NullCheck(Request("AccountID")) = "" Then
					sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Account ID was not found."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					Else
						rs("AccountPK") = rs2("AccountPK")
						rs("AccountID") = rs2("AccountID")
						rs("AccountName") = rs2("AccountName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Account ID is invalid."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					End If
				End If
				If Not NullCheck(Request("CategoryID")) = "" Then
					sql = "SELECT * FROM Category WITH (NOLOCK) WHERE CategoryID = '" & NullCheck(Request("CategoryID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Category ID was not found."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					Else
						rs("CategoryPK") = rs2("CategoryPK")
						rs("CategoryID") = rs2("CategoryID")
						rs("CategoryName") = rs2("CategoryName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Category ID is invalid."
						If Not FromClose Then
							WOLaborRec
						Else
							Exit Sub
						End If
					End If
				End If
			    If Not NullCheck(Request("Comments")) = "" Then
				    rs("Comments") = NullCheck(Request("Comments"))
			    End If
			    If Err.Number <> 0 Then
				    HeaderMSG = "The value provided for Comments is invalid."
					If Not FromClose Then
						WOLaborRec
					Else
						Exit Sub
					End If
			    End If
				On Error Goto 0
				If UCase(NullCheck(Request("Completed"))) = "Y" or UCase(NullCheck(Request("Completed"))) = "2" Then
					PostUpdateSQL = PostUpdateSQL + "UPDATE WOAssignStatus SET Completed = 1, CompletedDate = '" & SQLdatetimeADO(Request("WorkDate")) & "' WHERE WOPK = " & WOPK & " AND LaborPK = " & LaborPK & " "
				Else
					PostUpdateSQL = PostUpdateSQL + "UPDATE WOAssignStatus SET Completed = 0, CompletedDate = Null WHERE WOPK = " & WOPK & " AND LaborPK = " & LaborPK & " "
				End If
			End If
			db.dobatchupdate rs
			If db.dok Then
				'pl = rs("pk")
			Else
				Call OutputWAPError(db.derror)
			End If
			rs.close
			Call ProcessPostUpdateSQL(db)
		End If
	End If

End Sub

'====================================================================================================================================

Sub WOLabor()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo
	Dim CheckBoxType

	CardTitle = "WO Labor"
	CardCurrent = "WOLabor"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = 5
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = 5
    Else
	    pagesize = GlobalPageSize
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call WOLaborRecSave(db,False,"")

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
		isopen = True
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
	End If

	sql = _
	"SELECT     WOlabor.PK, WOlabor.LaborPK, lt.moduleid, WOlabor.LaborName, WOlabor.EstimatedHours, WOlabor.TotalHours, WOlabor.RegularHours, WOlabor.OvertimeHours, WOlabor.OtherHours, WOlabor.WorkDate, WOlabor.TimeIn, " &_
	"      WOlabor.TimeOut, WOlabor.AccountID, WOlabor.AccountName, WOlabor.CategoryID, WOlabor.CategoryName, WOlabor.TotalCost, " &_
	"                      WOlabor.TotalCharge, WOlabor.CostRegular, WOlabor.CostOvertime, WOlabor.CostOther, WOlabor.ChargeRate, WOlabor.ChargePercentage, WOlabor.RowVersionDate, Labor.Photo, WAS.Completed " &_
	"FROM         WOlabor WITH (NOLOCK)  LEFT OUTER JOIN WOAssignStatus WAS WITH (NOLOCK) ON WAS.WOPK = WOlabor.WOPK AND WAS.LaborPK = WOlabor.LaborPK " &_
	"                      LEFT OUTER JOIN Labor WITH (NOLOCK) ON WOlabor.LaborPK = Labor.LaborPK INNER JOIN " &_
	"                      LaborTypes lt WITH (NOLOCK) ON lt.LaborType = Labor.LaborType " &_
	"WHERE     (WOlabor.WOPK = " & WOPK & ") AND (WOlabor.RecordType = 2) " &_
	"ORDER BY WOlabor.LaborType, WOlabor.WorkDate, WOlabor.LaborName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

    If rs.Eof and IsOpen Then
		If Not UCase(CardFrom) = "WOLABORREC" Then
		    If cardlevelupdownamount <> 0 Then
        		CardSkipLevel = 1
                WOLaborRec
            End If
        Else
            CardSkipLevel = -1
            Call WOOptions("")
        End If
    End If

	Dim AccessToLARates, AccessToCRRates
	AccessToLARates = True
	AccessToCRRates = True
	If Not GetAccessRight(db,"LA_RATES_TABA",0) Then
		AccessToLARates = False
	End If
	If Not GetAccessRight(db,"CR_RATES_TABA",0) Then
		AccessToCRRates = False
	End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and False Then
			Call OutputWAPMsg("No Labor Found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap"><%
			If Not IsBOF Then %>
			<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;pr=<% =WOPK %>">&lt;</a>&nbsp;<%
			End If %><b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">WO #<% =WOID %></a></b><%
			If Not IsEOF Then %>
			<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;nr=<% =WOPK %>">&gt;</a><%
			End If %>
			<br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
            <%
   		    If lang = "WML" and IsPPC Then
		    Response.Write "<br/>"
		    Response.Write HStyleBegin
		    OutputBackButton
		    OutputHomeButton
		    Response.Write HStyleEnd
		    End If %>
			</p>
			<% End If %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				Dim AssignmentsCompleted
				If RS("Completed") Then
					AssignmentsCompleted = "<img src=""images/taskchecked.gif"">"
				Else
					AssignmentsCompleted = "<img src=""images/taskline.gif"">"
				End If
				If (Not AccessToLARates And RS("ModuleID") = "LA" And Not Trim(RS("LaborPK")) = Trim(GetSession("UserPK"))) or (Not AccessToCRRates And RS("ModuleID") = "CR" And Not Trim(RS("LaborPK")) = Trim(GetSession("CraftPK"))) Then
				Else %>
				<% =LStyleBegin %><a href="default.asp?card=wolaborrec&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>"><% =WAPValidate(NullCheck(rs("LaborName"))) %>: [<% =NullCheck(RS("EstimatedHours")) & " Est] [" & NullCheck(RS("TotalHours")) & " Actual" & "] " & DateNullCheck(RS("WorkDate")) %></a><% =LStyleEnd %><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
				    Response.Write "<br/>"
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Next</a></b>
		<%
		End If
		If IsOpen Then %>
		&nbsp;<b><a href="default.asp?card=WOLaborRec&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=-1">Add</a></b><%
		End If %>
		</p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOPartRec()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec

    newcontextpage = True

	CardTitle = "WO Material"
	CardCurrent = "WOPartRec"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

	pk = Request("PK")
	If pk = "" Then
		pk = GetSession("WOPARTPK" & CardCurrentLevel)
		If pk = "" Then
		    pk = "-1"
		End If
	Else
		Call SetSession("WOPARTPK" & CardCurrentLevel,pk)
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
		isopen = True
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
	End If

	sql = _
	"SELECT     WOpart.*, Part.Photo, PurchaseOrder.POID as PONumber, PurchaseOrder.InvoiceNumber, Part.PartDescription " &_
	"FROM         WOpart WITH (NOLOCK) " &_
	"			 LEFT OUTER JOIN Part WITH (NOLOCK) ON WOpart.PartPK = Part.PartPK " &_
	"			 LEFT OUTER JOIN PurchaseOrderReceive WITH (NOLOCK) ON PurchaseOrderReceive.PK = WOpart.POReceivedPK " &_
	"			 LEFT OUTER JOIN PurchaseOrder WITH (NOLOCK) ON PurchaseOrder.POPK = PurchaseOrderReceive.POPK " &_
	"WHERE     (WOpart.PK = " & PK & ") AND (WOpart.RecordType = 2) " &_
	"ORDER BY WOpart.LocationID, WOpart.PartID "

	Set rs = db.RunSQLReturnRS_RW(sql,"")
	Call CheckDB(db)

	If PK = "-1" Then
		Call BuildFields("PartID","Item #","C",GlobalFieldLength,"*M","true","",rs,PK)
		Call BuildFields("LocationID","Location ID","C",GlobalFieldLength,"*M","true",GetSession("SRID"),rs,PK)
	Else
		Call BuildFields("LocationID","Location ID","C",GlobalFieldLength,"*M","true","",rs,PK)
	End If
	Call BuildFields("QuantityActual","Actual Qty","C",GlobalFieldLength,"*M","true","1",rs,PK)
	Call BuildFields("OtherCost","Other Cost","C",GlobalFieldLength,"*M","true","",rs,PK)
	If IsPocketIE or IsBlackBerry Then
	Call BuildFields("AccountID","Account ID","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("CategoryID","Category ID","C",GlobalFieldLength,"*M","true","",rs,PK)

	Call BuildFields("Serial","Serial #","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("SerialReplaced","Existing Serial #","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("SerialReplaceToLocationID","Move Existing to","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("SerialReplaceExisting","Replace Existing?","B","","","",False,rs,PK)
	Call BuildFields("SerialReplacedOutOfService","Set Existing as Out of Service?","B","","","",True,rs,PK)

	End If
	If Not PK = "-1" and IsOpen Then
		Call BuildFields("Delete","Delete Record?","B","","","",False,rs,PK)
	End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("The Item was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2">WO #<% =WOID %></a></b><br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
			</p>
			<% End If %>
			<p mode="wrap"><% =HStyleBegin %><b>
			<% If PK = "-1" Then %>
			New Item<%
			Else %>
			[<% =WAPValidate(NullCheck(rs("PartID"))) %>] [<% =WAPValidate(NullCheck(rs("PartName"))) %>] [<% =WAPValidate(NullCheck(RS("LocationID"))) %>] [<% =WAPValidate(NullCheck(RS("QuantityEstimated"))) %>&nbsp;Est] [<% =WAPValidate(NullCheck(RS("QuantityActual"))) %>&nbsp;Actual]
			<% End If %>
			</b><% =HStyleEnd %></p><%
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If IsOpen Then
			WOPartRecSubmit rs
			End If
		Else
			WOPartRecSubmit rs
			If IsOpen Then
			OutputBackButton
			End If
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen Then
		%>
		</td><td align="right">
		<%
		WOPartRecSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WOPartRecSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="WOPart"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="wopk" value="<% =WOPK %>"/>
			<postfield name="pk" value="<% =PK %>"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="WOPart"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wopk" value="<% =WOPK %>"/>
	<input type="hidden" name="pk" value="<% =PK %>"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub WOPartRecSave(ByRef db)
	Dim rs2, PartPK, LocationPK
	pk = Request("pk")
	If Not pk = "" Then
		sql = "SELECT * FROM WOPart WITH (NOLOCK) WHERE PK = " & pk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If pk = "-1" Then
			rs.addnew()
		End If
		If Not rs.Eof Then
			If UCase(Request("Delete")) = "Y" or UCase(Request("Delete")) = "2" Then
				rs.delete()
				Call SetSession("WOPartPK" & CardCurrentLevel,"")
                Call SetSession("WOPartPK" & CardCurrentLevel+1,"")
			Else
				If pk = "-1" Then
					rs("RecordType") = 2
					rs("WOPK") = WOPK
					sql = "SELECT * FROM Part WITH (NOLOCK) WHERE PartID = '" & NullCheck(Request("PartID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Part ID was not found."
						WOPartRec
					Else
						If pk = "-1" Then
							rs("PartPK") = rs2("PartPK")
							rs("PartID") = rs2("PartID")
							rs("PartName") = rs2("PartName")
						End If
						PartPK = rs2("PartPK")
					End If
				Else
					PartPK = rs("PartPK")
				End If
				On Error Resume Next
				If Not NullCheck(Request("LocationID")) = "" Then
					sql = "SELECT * FROM Location WITH (NOLOCK) WHERE LocationID = '" & NullCheck(Request("LocationID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Location ID was not found."
						WOPartRec
					Else
						rs("LocationPK") = rs2("LocationPK")
						rs("LocationID") = rs2("LocationID")
						rs("LocationName") = rs2("LocationName")
						LocationPK = rs2("LocationPK")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Location ID is invalid."
						WOPartRec
					End If
					If Not PartPK = "" and Not LocationPK = "" Then
						sql = "SELECT * " &_
							  "	FROM PartLocation WITH (NOLOCK) " &_
							  "	WHERE PartPK = " & PartPK & " AND " &_
							  "	LocationPK = " & LocationPK
						Set rs2 = db.RunSQLReturnRS(sql,"")
						Call CheckDB(db)
						If Not rs2.eof Then
							rs("IssueUnitCost") = rs2("IssueUnitCost")
							rs("IssueUnitChargePrice") = rs2("IssueUnitChargePrice")
							rs("IssueUnitChargePercentage") = rs2("IssueUnitChargePercentage")
						End If
					End If
				Else
					If Not PartPK = "" Then
						sql = "SELECT * " &_
							  "	FROM Part WITH (NOLOCK) " &_
							  "	WHERE PartPK = " & PartPK
						Set rs2 = db.RunSQLReturnRS(sql,"")
						Call CheckDB(db)
						If Not rs2.eof Then
							rs("IssueUnitCost") = rs2("IssueUnitCost")
							rs("IssueUnitChargePrice") = rs2("IssueUnitChargePrice")
							rs("IssueUnitChargePercentage") = rs2("IssueUnitChargePercentage")
						End If
					End If
				End If
				If Not NullCheck(Request("QuantityActual")) = "" Then
					rs("QuantityActual") = NullCheck(Request("QuantityActual"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Actual Qty is invalid."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("OtherCost")) = "" Then
					rs("OtherCost") =	NullCheck(Request("OtherCost"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Other Cost is invalid."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("AccountID")) = "" Then
					sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Account ID was not found."
						WOPartRec
					Else
						rs("AccountPK") = rs2("AccountPK")
						rs("AccountID") = rs2("AccountID")
						rs("AccountName") = rs2("AccountName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Account ID is invalid."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("CategoryID")) = "" Then
					sql = "SELECT * FROM Category WITH (NOLOCK) WHERE CategoryID = '" & NullCheck(Request("CategoryID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Category ID was not found."
						WOPartRec
					Else
						rs("CategoryPK") = rs2("CategoryPK")
						rs("CategoryID") = rs2("CategoryID")
						rs("CategoryName") = rs2("CategoryName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Category ID is invalid."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("Serial")) = "" Then
					rs("Serial") =	NullCheck(Request("Serial"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Serial # is invalid."
						WOPartRec
					End If
				End If
			    If UCase(Request("SerialReplaceExisting")) = "Y" or UCase(Request("SerialReplaceExisting")) = "2" Then
			        rs("SerialReplaceExisting") = "S"
			    Else
			        rs("SerialReplaceExisting") = Null
			    End If
			    If UCase(Request("SerialReplacedOutOfService")) = "Y" or UCase(Request("SerialReplacedOutOfService")) = "2" Then
			        rs("SerialReplacedOutOfService") = 1
			    Else
			        rs("SerialReplacedOutOfService") = 0
			    End If
				If Not NullCheck(Request("SerialReplaced")) = "" Then
					rs("SerialReplaced") =	NullCheck(Request("SerialReplaced"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Existing Serial # is invalid."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("SerialReplaceToLocationID")) = "" Then
					sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetID = '" & NullCheck(Request("SerialReplaceToLocationID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Move Existing to Location was not found."
						WOPartRec
					Else
						rs("SerialReplaceToLocationID") = NullCheck(Request("SerialReplaceToLocationID"))
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Move Existing to Location is invalid."
						WOPartRec
					End If
				End If

				On Error Goto 0
			End If
			db.dobatchupdate rs
			If db.dok Then
				'pl = rs("pk")
			Else
				Call OutputWAPError(db.derror)
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub WOPart()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo
	Dim CheckBoxType

	CardTitle = "WO Materials"
	CardCurrent = "WOPart"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = 5
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = 5
    Else
	    pagesize = GlobalPageSize
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call WOPartRecSave(db)

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
		isopen = True
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
	End If

	sql = _
	"SELECT     WOpart.*, Part.Photo, PurchaseOrder.POID as PONumber, PurchaseOrder.InvoiceNumber, Part.PartDescription " &_
	"FROM         WOpart WITH (NOLOCK) " &_
	"			 LEFT OUTER JOIN Part WITH (NOLOCK) ON WOpart.PartPK = Part.PartPK " &_
	"			 LEFT OUTER JOIN PurchaseOrderReceive WITH (NOLOCK) ON PurchaseOrderReceive.PK = WOpart.POReceivedPK " &_
	"			 LEFT OUTER JOIN PurchaseOrder WITH (NOLOCK) ON PurchaseOrder.POPK = PurchaseOrderReceive.POPK " &_
	"WHERE     (WOpart.WOPK = " & WOPK & ") AND (WOpart.RecordType = 2) " &_
	"ORDER BY WOpart.LocationID, WOpart.PartID "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

    If rs.Eof and IsOpen Then
		If Not UCase(CardFrom) = "WOPARTREC" Then
		    If cardlevelupdownamount <> 0 Then
    		    CardSkipLevel = 1
                WOPartRec
            End If
        Else
            CardSkipLevel = -1
            Call WOOptions("")
        End If
    End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and False Then
			Call OutputWAPMsg("No Materials were Found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap"><%
			If Not IsBOF Then %>
			<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;pr=<% =WOPK %>">&lt;</a>&nbsp;<%
			End If %><b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">WO #<% =WOID %></a></b><%
			If Not IsEOF Then %>
			<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;nr=<% =WOPK %>">&gt;</a><%
			End If %>
			<br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
            <%
		    If lang = "WML" and IsPPC Then
		    Response.Write "<br/>"
		    Response.Write HStyleBegin
		    OutputBackButton
		    OutputHomeButton
		    Response.Write HStyleEnd
		    End If %>
			</p>
			<% End If %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else %>
				<% =LStyleBegin %><a href="default.asp?card=wopartrec&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>">[<% =WAPValidate(NullCheck(rs("PartID"))) %>] [<% =WAPValidate(NullCheck(rs("PartName"))) %>] [<% =WAPValidate(NullCheck(RS("LocationID"))) %>] [<% =WAPValidate(NullCheck(RS("QuantityEstimated"))) %>&nbsp;Est] [<% =WAPValidate(NullCheck(RS("QuantityActual"))) %>&nbsp;Actual]</a><% =LStyleEnd %><%
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
					Response.Write "<br/>"
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Next</a></b>
		<%
		End If
		If IsOpen Then %>
		&nbsp;<b><a href="default.asp?card=WOPartRec&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=-1">Add</a></b><%
		End If %>
		</p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOMiscCostRec()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec

    newcontextpage = True

	CardTitle = "WO Other Cost"
	CardCurrent = "WOMiscCostRec"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

	pk = Request("PK")
	If pk = "" Then
		pk = GetSession("WOMiscCostPK" & CardCurrentLevel)
		If pk = "" Then
		    pk = "-1"
		End If
	Else
		Call SetSession("WOMiscCostPK" & CardCurrentLevel,pk)
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
		isopen = True
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
	End If

	sql = _
	"SELECT     PK, MiscCostName, MiscCostDesc, InvoiceNumber, MiscCostDate, AccountID, AccountName, CategoryID, CategoryName, EstimatedCost, ActualCost, " &_
	"                      ChargePercentage, TotalCharge, RowVersionDate, CompanyPK, CompanyID, CompanyName, LaborPK, LaborID, LaborName, Comments " &_
	"FROM         WOmiscCost WITH (NOLOCK) " &_
	"WHERE     (WOMiscCost.PK = " & PK & ") AND (RecordType = 2) " &_
	"ORDER BY MiscCostDate, MiscCostName "

	Set rs = db.RunSQLReturnRS_RW(sql,"")
	Call CheckDB(db)

	Call BuildFields("MiscCostName","Name","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("MiscCostDesc","Description","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("CompanyID","Company ID","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("LaborID","Labor ID","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("InvoiceNumber","Invoice #","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("MiscCostDate","Date","C",GlobalFieldLength,"*M","true",CStr(DateNullCheck(Date())),rs,PK)
	If IsPocketIE or IsBlackBerry Then
	Call BuildFields("AccountID","Account ID","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("CategoryID","Category ID","C",GlobalFieldLength,"*M","true","",rs,PK)
	End If
	Call BuildFields("ActualCost","Actual Cost","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("Comments","Comments","C",GlobalFieldLength,"*M","true","",rs,PK)
	If Not PK = "-1" and IsOpen Then
		Call BuildFields("Delete","Delete Record?","B","","","",False,rs,PK)
	End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("The Other Cost was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2">WO #<% =WOID %></a></b><br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
			</p>
			<% End If %>
			<p mode="wrap"><% =HStyleBegin %><b>
			<% If PK = "-1" Then %>
			New Other Cost<%
			Else %>
			[<% =WAPValidate(NullCheck(rs("MiscCostName"))) %>] [<% =DateNullCheck(RS("MiscCostDate")) %>] [<% =FormatNumber(WAPValidate(NullCheck(RS("EstimatedCost"))),2,-2,0,0) %>&nbsp;Est] [<% =FormatNumber(WAPValidate(NullCheck(RS("ActualCost"))),2,-2,0,0) %>&nbsp;Actual]
			<% End If %>
			</b><% =HStyleEnd %></p><%
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If IsOpen Then
			WOMiscCostRecSubmit rs
			End If
		Else
			If IsOpen Then
			WOMiscCostRecSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen Then
		%>
		</td><td align="right">
		<%
		WOMiscCostRecSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WOMiscCostRecSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="WOMiscCost"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="wopk" value="<% =WOPK %>"/>
			<postfield name="pk" value="<% =PK %>"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="WOMiscCost"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wopk" value="<% =WOPK %>"/>
	<input type="hidden" name="pk" value="<% =PK %>"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub WOMiscCostRecSave(ByRef db)
	Dim rs2, MiscCostPK
	pk = Request("pk")
	If Not pk = "" Then
		sql = "SELECT * FROM WOMiscCost WITH (NOLOCK) WHERE PK = " & pk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If pk = "-1" Then
			rs.addnew()
		End If
		If Not rs.Eof Then
			If UCase(Request("Delete")) = "Y" or UCase(Request("Delete")) = "2" Then
				rs.delete()
				Call SetSession("WOMiscCostPK" & CardCurrentLevel,"")
                Call SetSession("WOMiscCostPK" & CardCurrentLevel+1,"")
			Else
				If pk = "-1" Then
					rs("RecordType") = 2
					rs("WOPK") = WOPK
				End If
				On Error Resume Next
				If Not NullCheck(Request("MiscCostName")) = "" Then
					rs("MiscCostName") = NullCheck(Request("MiscCostName"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Name is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("MiscCostDesc")) = "" Then
					rs("MiscCostDesc") = NullCheck(Request("MiscCostDesc"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Description is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("CompanyID")) = "" Then
					sql = "SELECT * FROM Company WITH (NOLOCK) WHERE CompanyID = '" & NullCheck(Request("CompanyID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Company ID was not found."
						WOMiscCostRec
					Else
						rs("CompanyPK") = rs2("CompanyPK")
						rs("CompanyID") = rs2("CompanyID")
						rs("CompanyName") = rs2("CompanyName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Company ID is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("LaborID")) = "" Then
					sql = "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborID = '" & NullCheck(Request("LaborID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Labor ID was not found."
						WOMiscCostRec
					Else
						rs("LaborPK") = rs2("LaborPK")
						rs("LaborID") = rs2("LaborID")
						rs("LaborName") = rs2("LaborName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Labor ID is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("InvoiceNumber")) = "" Then
					rs("InvoiceNumber") = NullCheck(Request("InvoiceNumber"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Invoice # is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("MiscCostDate")) = "" Then
					rs("MiscCostDate") = SQLdatetimeADO(Request("MiscCostDate"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Date is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("AccountID")) = "" Then
					sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Account ID was not found."
						WOMiscCostRec
					Else
						rs("AccountPK") = rs2("AccountPK")
						rs("AccountID") = rs2("AccountID")
						rs("AccountName") = rs2("AccountName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Account ID is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("CategoryID")) = "" Then
					sql = "SELECT * FROM Category WITH (NOLOCK) WHERE CategoryID = '" & NullCheck(Request("CategoryID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Category ID was not found."
						WOMiscCostRec
					Else
						rs("CategoryPK") = rs2("CategoryPK")
						rs("CategoryID") = rs2("CategoryID")
						rs("CategoryName") = rs2("CategoryName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Category ID is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("ActualCost")) = "" Then
					rs("ActualCost") = NullCheck(Request("ActualCost"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Actual Cost is invalid."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("Comments")) = "" Then
					rs("Comments") =	NullCheck(Request("Comments"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "The value provided for Comments is invalid."
					WOMiscCostRec
				End If
				On Error Goto 0
			End If
			db.dobatchupdate rs
			If db.dok Then
				'pl = rs("pk")
			Else
				If InStr(db.derror,"NULL") > 0 Then
					HeaderMSG = "A value must be provided for Name."
					WOMiscCostRec
				Else
					Call OutputWAPError(db.derror)
				End If
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub WOMiscCost()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo
	Dim CheckBoxType

	CardTitle = "WO Other Costs"
	CardCurrent = "WOMiscCost"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = 5
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = 5
    Else
	    pagesize = GlobalPageSize
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call WOMiscCostRecSave(db)

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
		isopen = True
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
	End If

	sql = _
	"SELECT     PK, MiscCostName, MiscCostDesc, InvoiceNumber, MiscCostDate, AccountID, AccountName, CategoryID, CategoryName, EstimatedCost, ActualCost, " &_
	"                      ChargePercentage, TotalCharge, RowVersionDate, CompanyPK, CompanyID, CompanyName, LaborPK, LaborID, LaborName, Comments " &_
	"FROM         WOmiscCost WITH (NOLOCK) " &_
	"WHERE     (WOPK = " & WOPK & ") AND (RecordType = 2) " &_
	"ORDER BY MiscCostDate, MiscCostName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

    If rs.Eof and IsOpen Then
		If Not UCase(CardFrom) = "WOMISCCOSTREC" Then
		    If cardlevelupdownamount <> 0 Then
        		CardSkipLevel = 1
                WOMiscCostRec
            End If
        Else
            CardSkipLevel = -1
            Call WOOptions("")
        End If
    End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and False Then
			Call OutputWAPMsg("No Other Costs were Found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap"><%
			If Not IsBOF Then %>
			<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;pr=<% =WOPK %>">&lt;</a>&nbsp;<%
			End If %><b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">WO #<% =WOID %></a></b><%
			If Not IsEOF Then %>
			<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;nr=<% =WOPK %>">&gt;</a><%
			End If %>
			<br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
		    <% If lang = "WML" and IsPPC Then
		    Response.Write "<br/>"
		    Response.Write HStyleBegin
		    OutputBackButton
		    OutputHomeButton
		    Response.Write HStyleEnd
		    End If %>
			</p>
			<% End If %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else %>
				<% =LStyleBegin %><a href="default.asp?card=womisccostrec&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>">[<% =WAPValidate(NullCheck(rs("MiscCostName"))) %>] [<% =DateNullCheck(RS("MiscCostDate")) %>] [<% =FormatNumber(WAPValidate(NullCheck(RS("EstimatedCost"))),2,-2,0,0) %>&nbsp;Est] [<% =FormatNumber(WAPValidate(NullCheck(RS("ActualCost"))),2,-2,0,0) %>&nbsp;Actual]</a><% =LStyleEnd %><%
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
			        Response.Write "<br/>"
			    End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Next</a></b>
		<%
		End If
		If IsOpen Then %>
		&nbsp;<b><a href="default.asp?card=WOMiscCostRec&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=-1">Add</a></b><%
		End If %>
		</p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASSpecsRec()

	Dim db, sql, rs, ASE, assetid, assetname, photo, parentlocationall, parentequipmentall, newrec, IsLocation, rseof

    newcontextpage = True

	CardTitle = "Asset Specification"
	CardCurrent = "ASSpecsRec"
	CardCurrentLevel = GetCardLevel()

	GetAssetPK

	pk = Request("PK")
	If pk = "" Then
		pk = GetSession("ASSpecPK" & CardCurrentLevel)
		If pk = "" Then
		    pk = "-1"
		End If
	Else
		Call SetSession("ASSpecPK" & CardCurrentLevel,pk)
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	ASE = GetAccessRight(db,"ASE",0)

    If Not AssetPK = "-1" Then
	    sql = _
	    "SELECT Asset.*, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!') " &_
	    "FROM Asset WITH (NOLOCK) INNER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
	    "WHERE Asset.AssetPK = " & AssetPK & " "

	    Set rs = db.RunSQLReturnRS(sql,"")
	    Call CheckDB(db)

	    If Not rs.eof Then
		    assetpk = NullCheck(rs("AssetPK"))
		    assetid = WAPValidate(NullCheck(RS("AssetID")))
		    assetname = WAPValidate(NullCheck(RS("AssetName")))
		    photo = NullCheck(rs("Photo"))
            parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!","<br/>")
            parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!","<br/>")
            If Not parentlocationall = "" Then
                parentlocationall = parentlocationall & "<br/>"
            End If
            If Not parentequipmentall = "" Then
                parentequipmentall = parentequipmentall & "<br/>"
            End If
	    Else
		    Call OutputWAPError("The Asset was not found.")
	    End If

	    sql = _
	    "SELECT     AssetSpecification.* " &_
	    "FROM       AssetSpecification WITH (NOLOCK)  " &_
	    "WHERE     (AssetSpecification.PK = " & PK & " )"

	    Set rs = db.RunSQLReturnRS_RW(sql,"")
	    Call CheckDB(db)

        If rs.eof Then
            rseof = True
        Else
            rseof = False
        End If

	End If

	Call BuildFields("ValueText","Text Value","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("ValueDate","Date Value","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("ValueNumeric","Numeric Value","C",GlobalFieldLength,"*M","true","",rs,AssetPK)

	Call StartMobileDocument(CardTitle)
		If rseof and Not AssetPK = "-1" Then
			Call OutputWAPMsg("The Asset Specification was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<% If AssetPK = "-1" Then %>
		    <b><% =HStyleBegin %>New Asset<% =HStyleEnd %></b>
			<% Else %>
		    <b><% =HStyleBegin %><% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)<% =HStyleEnd %></b>
		    <% End If %>
			</p>
			<% End If
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			' We do not need to check for ASN because they would not have AssetPK = -1 if they could not get here from the main menu
			If IsOpen and (ASE or AssetPK = "-1") Then
			ASSpecsSubmit rs
			End If
		Else
			If IsOpen and (ASE or AssetPK = "-1") Then
			ASSpecsSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen and (ASE or AssetPK = "-1") Then
		%>
		</td><td align="right">
		<%
		ASSpecsSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		Call CloseObj(rs)
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub ASSpecsSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="ASSpecs"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="Assetpk" value="<% =AssetPK %>"/>
			<postfield name="PK" value="<% =PK %>"/>
			<postfield name="POSTEDASSPECS" value="Y"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="ASSpecs"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="Assetpk" value="<% =AssetPK %>"/>
	<input type="hidden" name="PK" value="<% =PK %>"/>
	<input type="hidden" name="POSTEDASSPECS" value="Y"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub ASSpecsRecSave(ByRef db)
	Dim rs2, ASSpecPK
	pk = Request("pk")
	If Not pk = "" Then
		sql = "SELECT * FROM AssetSpecification WITH (NOLOCK) WHERE PK = " & pk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If pk = "-1" Then
			rs.addnew()
		End If
		If Not rs.Eof Then
			If UCase(Request("Delete")) = "Y" or UCase(Request("Delete")) = "2" Then
				rs.delete()
				Call SetSession("ASSpecPK" & CardCurrentLevel,"")
                Call SetSession("ASSpecPK" & CardCurrentLevel+1,"")
			Else
				If pk = "-1" Then
					rs("AssetPK") = AssetPK
				End If
				On Error Resume Next

                Call SaveField(db,"ValueText","Text Value","","","C","ASSpecsRec",False)
                Call SaveField(db,"ValueDate","Date Value","","","C","ASSpecsRec",False)
                Call SaveField(db,"ValueNumeric","Numeric Value","","","C","ASSpecsRec",False)

				On Error Goto 0
			End If
			db.dobatchupdate rs
			If db.dok Then
			Else
				Call OutputWAPError(db.derror)
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub ASSpecs()

	Dim db, sql, rs, ASE, assetid, assetname, photo, parentlocationall, parentequipmentall, newrec, IsLocation, rseof
	Dim CheckBoxType

	CardTitle = "Asset Specifications"
	CardCurrent = "ASSpecs"
	CardCurrentLevel = GetCardLevel()

	GetAssetPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = 7
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = 5
    Else
	    pagesize = GlobalPageSize
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call ASSpecsRecSave(db)

    sql = _
    "SELECT Asset.*, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!') " &_
    "FROM Asset WITH (NOLOCK) INNER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
    "WHERE Asset.AssetPK = " & AssetPK & " "

    Set rs = db.RunSQLReturnRS(sql,"")
    Call CheckDB(db)

    If Not rs.eof Then
	    assetpk = NullCheck(rs("AssetPK"))
	    assetid = WAPValidate(NullCheck(RS("AssetID")))
	    assetname = WAPValidate(NullCheck(RS("AssetName")))
	    photo = NullCheck(rs("Photo"))
        parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!","<br/>")
        parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!","<br/>")
        If Not parentlocationall = "" Then
            parentlocationall = parentlocationall & "<br/>"
        End If
        If Not parentequipmentall = "" Then
            parentequipmentall = parentequipmentall & "<br/>"
        End If
    Else
	    Call OutputWAPError("The Asset was not found.")
    End If

	sql = _
	"SELECT AssetSpecification.* FROM AssetSpecification INNER JOIN Specification ON Specification.SpecificationPK = AssetSpecification.SpecificationPK " &_
	"WHERE     (AssetSpecification.AssetPK = " & AssetPK & ") " &_
	"ORDER BY Specification.Categoryname, Specification.SpecificationName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

    ' Do not automatically go to add mode if there are no specs
    'If rs.Eof and IsOpen Then
	'	If Not UCase(CardFrom) = "ASSPECSREC" Then
	'	    If cardlevelupdownamount <> 0 Then
    '    		CardSkipLevel = 1
    '            ASSpecsRec
    '        End If
    '    Else
    '        CardSkipLevel = -1
    '        ASOptions
    '    End If
    'End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and False Then
			Call OutputWAPMsg("No Specifications were Found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			    <p align="center" mode="wrap">
		        <b><% =HStyleBegin %><% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)<% =HStyleEnd %></b>
			    <%
		        If lang = "WML" and IsPPC Then
		        Response.Write "<br/>"
		        Response.Write HStyleBegin
		        OutputBackButton
		        OutputHomeButton
		        Response.Write HStyleEnd
		        End If %>
		        </p>
		    <% End If %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
			    Dim ValueCombined
			    ValueCombined = ""
			    If Not NullCheck(rs("ValueText")) = "" Then
			        If ValueCombined = "" Then
			            ValueCombined = NullCheck(rs("ValueText"))
			        Else
			            ValueCombined = ValueCombined & " - " & NullCheck(rs("ValueText"))
			        End If
			    End If
			    If Not NullCheck(rs("ValueNumeric")) = "" Then
			        If ValueCombined = "" Then
			            ValueCombined = NullCheck(rs("ValueNumeric"))
			        Else
			            ValueCombined = ValueCombined & " - " & NullCheck(rs("ValueNumeric"))
			        End If
			    End If
			    If Not DateNullCheck(rs("ValueDate")) = "" Then
			        If ValueCombined = "" Then
			            ValueCombined = NullCheck(rs("ValueDate"))
			        Else
			            ValueCombined = ValueCombined & " - " & NullCheck(rs("ValueDate"))
			        End If
			    End If
			    If Not ValueCombined = "" Then
			        ValueCombined = "[" & ValueCombined & "]"
			    End If
			    %>
				<% If lang = "HTML" Then %><nobr><% End If %><% =LStyleBegin %><a href="default.asp?card=asspecsrec&amp;s=<% =SessionID %>&amp;assetpk=<% =assetpk %>&amp;pk=<% =rs("PK") %>">[<% If lang = "HTML" Then %><b><% End If %><% =WAPValidate(NullCheck(rs("SpecificationName"))) %><% If lang = "HTML" Then %></b><% End If %>] <% =WAPValidate(ValueCombined) %></a><% =LStyleEnd %><% If lang = "HTML" Then %></nobr><% End If %><%
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
			        Response.Write "<br/>"
			    End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;assetpk=<% =assetpk %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;assetpk=<% =assetpk %>">Next</a></b>
		<%
		End If
		If IsOpen and False Then %>
		&nbsp;<b><a href="default.asp?card=ASSpecsRec&amp;s=<% =SessionID %>&amp;assetpk=<% =assetpk %>&amp;pk=-1">Add</a></b><%
		End If %>
		</p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASLaborRec()

	Dim db, sql, rs, ASE, assetid, assetname, photo, parentlocationall, parentequipmentall, newrec, IsLocation, rseof

    newcontextpage = True

	CardTitle = "Asset Labor / Contacts"
	CardCurrent = "ASLaborRec"
	CardCurrentLevel = GetCardLevel()

	GetAssetPK

	pk = Request("PK")
	If pk = "" Then
		pk = GetSession("ASLaborPK" & CardCurrentLevel)
		If pk = "" Then
		    pk = "-1"
		End If
	Else
		Call SetSession("ASLaborPK" & CardCurrentLevel,pk)
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	ASE = GetAccessRight(db,"ASE",0)

    If Not AssetPK = "-1" Then
	    sql = _
	    "SELECT Asset.*, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!') " &_
	    "FROM Asset WITH (NOLOCK) INNER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
	    "WHERE Asset.AssetPK = " & AssetPK & " "

	    Set rs = db.RunSQLReturnRS(sql,"")
	    Call CheckDB(db)

	    If Not rs.eof Then
		    assetpk = NullCheck(rs("AssetPK"))
		    assetid = WAPValidate(NullCheck(RS("AssetID")))
		    assetname = WAPValidate(NullCheck(RS("AssetName")))
		    photo = NullCheck(rs("Photo"))
            parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!","<br/>")
            parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!","<br/>")
            If Not parentlocationall = "" Then
                parentlocationall = parentlocationall & "<br/>"
            End If
            If Not parentequipmentall = "" Then
                parentequipmentall = parentequipmentall & "<br/>"
            End If
	    Else
		    Call OutputWAPError("The Asset was not found.")
	    End If

	    sql = _
	    "SELECT     al.*, lt.moduleid, l.LaborName, l.LaborType, l.LaborTypeDesc, l.FirstName, l.MiddleName, l.LastName, l.JobTitle, l.RepairCenterID, l.RepairCenterName, l.PhoneHome, l.PhoneWork, l.PhoneMobile, l.Pager, " &_
	    "                      l.Fax, l.Email, l.Photo, l.LaborTypeDesc AS CategoryName " &_
	    "FROM         AssetLabor al WITH (NOLOCK) INNER JOIN " &_
	    "                      Labor l WITH (NOLOCK) ON al.LaborPK = l.LaborPK INNER JOIN " &_
	    "                      LaborTypes lt WITH (NOLOCK) ON lt.LaborType = l.LaborType " &_
	    "WHERE     (al.pk = " & pk & ") " &_
	    "ORDER BY CategoryName, al.Priority, l.LaborName "

	    Set rs = db.RunSQLReturnRS_RW(sql,"")
	    Call CheckDB(db)

        If rs.eof Then
            rseof = True
        Else
            rseof = False
        End If

	End If

	Call BuildFields("LaborName","Name","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("LaborType","Type","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("JobTitle","Job Title","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("PhoneWork","Work Phone","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("PhoneHome","Home Phone","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("PhoneMobile","Mobile Phone","C",GlobalFieldLength,"*M","true","",rs,AssetPK)
	Call BuildFields("Email","Email","C",GlobalFieldLength,"*M","true","",rs,AssetPK)

	Call StartMobileDocument(CardTitle)
		If rseof and Not AssetPK = "-1" Then
			Call OutputWAPMsg("The Asset Labor / Contact was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<% If AssetPK = "-1" Then %>
		    <b><% =HStyleBegin %>New Asset<% =HStyleEnd %></b>
			<% Else %>
		    <b><% =HStyleBegin %><% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)<% =HStyleEnd %></b>
		    <% End If %>
			</p>
			<% End If
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			' We do not need to check for ASN because they would not have AssetPK = -1 if they could not get here from the main menu
			If False and IsOpen and (ASE or AssetPK = "-1") Then
			ASLaborSubmit rs
			End If
		Else
			If False and IsOpen and (ASE or AssetPK = "-1") Then
			ASLaborSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If False and IsOpen and (ASE or AssetPK = "-1") Then
		%>
		</td><td align="right">
		<%
		ASLaborSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		Call CloseObj(rs)
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub ASLaborSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="ASSpecs"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="Assetpk" value="<% =AssetPK %>"/>
			<postfield name="PK" value="<% =PK %>"/>
			<postfield name="POSTEDASLABOR" value="Y"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="ASSpecs"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="Assetpk" value="<% =AssetPK %>"/>
	<input type="hidden" name="PK" value="<% =PK %>"/>
	<input type="hidden" name="POSTEDASLABOR" value="Y"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub ASLaborRecSave(ByRef db)
	Dim rs2, ASSpecPK
	pk = Request("pk")
	If Not pk = "" Then
		sql = "SELECT * FROM AssetLabor WITH (NOLOCK) WHERE PK = " & pk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If pk = "-1" Then
			rs.addnew()
		End If
		If Not rs.Eof Then
			If UCase(Request("Delete")) = "Y" or UCase(Request("Delete")) = "2" Then
				rs.delete()
				Call SetSession("ASLaborPK" & CardCurrentLevel,"")
                Call SetSession("ASLaborPK" & CardCurrentLevel+1,"")
			Else
				If pk = "-1" Then
					rs("AssetPK") = AssetPK
				End If
				On Error Resume Next

				On Error Goto 0
			End If
			db.dobatchupdate rs
			If db.dok Then
			Else
				Call OutputWAPError(db.derror)
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub ASLabor()

	Dim db, sql, rs, ASE, assetid, assetname, photo, parentlocationall, parentequipmentall, newrec, IsLocation, rseof
	Dim CheckBoxType

	CardTitle = "Asset Labor / Contacts"
	CardCurrent = "ASLabor"
	CardCurrentLevel = GetCardLevel()

	GetAssetPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = 5
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = 5
    Else
	    pagesize = GlobalPageSize
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call ASLaborRecSave(db)

    sql = _
    "SELECT Asset.*, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!') " &_
    "FROM Asset WITH (NOLOCK) INNER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
    "WHERE Asset.AssetPK = " & AssetPK & " "

    Set rs = db.RunSQLReturnRS(sql,"")
    Call CheckDB(db)

    If Not rs.eof Then
	    assetpk = NullCheck(rs("AssetPK"))
	    assetid = WAPValidate(NullCheck(RS("AssetID")))
	    assetname = WAPValidate(NullCheck(RS("AssetName")))
	    photo = NullCheck(rs("Photo"))
        parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!","<br/>")
        parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!","<br/>")
        If Not parentlocationall = "" Then
            parentlocationall = parentlocationall & "<br/>"
        End If
        If Not parentequipmentall = "" Then
            parentequipmentall = parentequipmentall & "<br/>"
        End If
    Else
	    Call OutputWAPError("The Asset was not found.")
    End If

	sql = _
    "SELECT     al.*, lt.moduleid, l.LaborName, l.LaborType, l.LaborTypeDesc, l.FirstName, l.MiddleName, l.LastName, l.JobTitle, l.RepairCenterID, l.RepairCenterName, l.PhoneHome, l.PhoneWork, l.PhoneMobile, l.Pager, " &_
    "                      l.Fax, l.Email, l.Photo, l.LaborTypeDesc AS CategoryName " &_
    "FROM         AssetLabor al WITH (NOLOCK) INNER JOIN " &_
    "                      Labor l WITH (NOLOCK) ON al.LaborPK = l.LaborPK INNER JOIN " &_
    "                      LaborTypes lt WITH (NOLOCK) ON lt.LaborType = l.LaborType " &_
    "WHERE     (al.assetpk = " & assetpk & ") " &_
    "ORDER BY CategoryName, al.Priority, l.LaborName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

    ' Do not automatically go to add mode if there are no specs
    'If rs.Eof and IsOpen Then
	'	If Not UCase(CardFrom) = "ASLABORREC" Then
	'	    If cardlevelupdownamount <> 0 Then
    '    		CardSkipLevel = 1
    '            ASLaborRec
    '        End If
    '    Else
    '        CardSkipLevel = -1
    '        ASOptions
    '    End If
    'End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and False Then
			Call OutputWAPMsg("No Specifications were Found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			    <p align="center" mode="wrap">
		        <b><% =HStyleBegin %><% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)<% =HStyleEnd %></b>
			    <%
		        If lang = "WML" and IsPPC Then
		        Response.Write "<br/>"
		        Response.Write HStyleBegin
		        OutputBackButton
		        OutputHomeButton
		        Response.Write HStyleEnd
		        End If %>
		        </p>
		    <% End If %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else %>
				<% =LStyleBegin %><a href="default.asp?card=aslaborrec&amp;s=<% =SessionID %>&amp;assetpk=<% =assetpk %>&amp;pk=<% =rs("PK") %>">[<% =WAPValidate(NullCheck(rs("LaborName"))) %>] [<% =WAPValidate(NullCheck(rs("PhoneWork"))) %>] [<% =WAPValidate(NullCheck(rs("PhoneMobile"))) %>]</a><% =LStyleEnd %><%
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
			        Response.Write "<br/>"
			    End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;assetpk=<% =assetpk %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;assetpk=<% =assetpk %>">Next</a></b>
		<%
		End If
		If IsOpen and False Then %>
		&nbsp;<b><a href="default.asp?card=ASLaborRec&amp;s=<% =SessionID %>&amp;assetpk=<% =assetpk %>&amp;pk=-1">Add</a></b><%
		End If %>
		</p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOAssignRec()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec, wostatus

    newcontextpage = True

	CardTitle = "WO Assign"
	CardCurrent = "WOAssignRec"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

	pk = Request("PK")
	If pk = "" Then
		pk = GetSession("WOAssignPK" & CardCurrentLevel)
		If pk = "" Then
		    pk = "-1"
		End If
	Else
		Call SetSession("WOAssignPK" & CardCurrentLevel,pk)
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
		isopen = True
		wostatus = ""
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
		wostatus = NullCheck(RS("Status"))
	End If

	sql = _
	"SELECT WOassign.PK, WOassign.IsAssigned, WOassign.LaborPK, WOassign.LaborID, " &_
	"       WOassign.LaborName, Labor.CraftID, Labor.CraftName, Labor.LaborType, Labor.LaborTypeDesc, WOassign.AssignedHours, WOassign.AssignedDate, Labor.Email, Labor.PagerEmail, Labor.Photo, " &_
	"       WOassign.RowVersionDate, WOassign.AssignedLead, WOassign.AssignedPDA " &_
	"FROM WOassign WITH (NOLOCK) " &_
	"     LEFT OUTER JOIN Labor Labor WITH (NOLOCK) ON WOassign.LaborPK = Labor.LaborPK " &_
	"WHERE (WOassign.PK = " & PK & ") " &_
	"ORDER BY WOassign.IsAssigned, Labor.LaborType DESC, WOassign.AssignedDate, WOassign.LaborName "

	Set rs = db.RunSQLReturnRS_RW(sql,"")
	Call CheckDB(db)

	Call BuildFields("LaborID","Labor ID","C",GlobalFieldLength,"*M","true","",rs,PK)
	Call BuildFields("AssignedDate","Date","C",GlobalFieldLength,"*M","true",CStr(DateNullCheck(Date())),rs,PK)
	Call BuildFields("AssignedHours","Hour(s)","C",GlobalFieldLength,"*M","true","1",rs,PK)
	Call BuildFields("AssignedLead","Assigned Lead?","B","","","",True,rs,PK)
	Call BuildFields("AssignedPDA","Assigned PDA User?","B","","","",True,rs,PK)

    ' if we implemented the below functionality we would need to check to
    ' see if autoissue on assign is turned on etc.
    'If UCase("WOStatus") = "REQUESTED" Then
    '    Call BuildFields("Issue","Issue WO?","B","","","",True,rs,PK)
    'End If

	If Not PK = "-1" and IsOpen Then
		Call BuildFields("Delete","Delete Record?","B","","","",False,rs,PK)
	End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("The Assignment was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2">WO #<% =WOID %></a></b><br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
			</p>
			<% End If %>
			<p mode="wrap"><% =HStyleBegin %><b>
			<% If PK = "-1" Then %>
			New Assignment<%
			Else %>
			[<% =WAPValidate(NullCheck(rs("LaborName"))) %>] [<% =DateNullCheck(RS("AssignedDate")) %>] [<% =NullCheck(RS("AssignedHours")) %>&nbsp;Hour(s)]
			<% End If %>
			</b><% =HStyleEnd %></p><%
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If IsOpen Then
			WOAssignRecSubmit rs
			End If
		Else
			If IsOpen Then
			WOAssignRecSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen Then
		%>
		</td><td align="right">
		<%
		WOAssignRecSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WOAssignRecSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="WOAssign"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="wopk" value="<% =WOPK %>"/>
			<postfield name="pk" value="<% =PK %>"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="WOAssign"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wopk" value="<% =WOPK %>"/>
	<input type="hidden" name="pk" value="<% =PK %>"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub WOAssignRecSave(ByRef db)
	Dim rs2, WOAssignPK
	pk = Request("pk")
	If Not pk = "" Then
		sql = "SELECT * FROM WOAssign WITH (NOLOCK) WHERE PK = " & pk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If pk = "-1" Then
			rs.addnew()
		End If
		If Not rs.Eof Then
			If UCase(Request("Delete")) = "Y" or UCase(Request("Delete")) = "2" Then
				rs.delete()
				Call SetSession("WOAssignPK" & CardCurrentLevel,"")
                Call SetSession("WOAssignPK" & CardCurrentLevel+1,"")
			Else
				If pk = "-1" Then
					rs("WOPK") = WOPK
        			rs("IsAssigned") = True
				End If
				On Error Resume Next
				If Not NullCheck(Request("LaborID")) = "" Then
					sql = "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborID = '" & NullCheck(Request("LaborID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "The Labor ID was not found."
						WOAssignRec
					Else
						rs("LaborPK") = rs2("LaborPK")
						rs("LaborID") = rs2("LaborID")
						rs("LaborName") = rs2("LaborName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Labor ID is invalid."
						WOAssignRec
					End If
				End If
				If Not NullCheck(Request("AssignedDate")) = "" Then
					rs("AssignedDate") = SQLdatetimeADO(Request("AssignedDate"))
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Assigned Date is invalid."
						WOAssignRec
					End If
				Else
					HeaderMSG = "A value must be provided for Assigned Date."
					WOAssignRec
				End If
				If Not NullCheck(Request("AssignedHours")) = "" Then
				    rs("AssignedHours") = Request("AssignedHours")
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Assigned Hours is invalid."
						WOAssignRec
					End If
				Else
					HeaderMSG = "A value must be provided for Assigned Hours."
					WOAssignRec
				End If
				If Not NullCheck(Request("AssignedLead")) = "" Then
				    If Request("AssignedLead") = "Y" or Request("AssignedLead") = "2" Then
				        rs("AssignedLead") = True
				    Else
				        rs("AssignedLead") = False
				    End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Assigned Lead is invalid."
						WOAssignRec
					End If
				End If
				If Not NullCheck(Request("AssignedPDA")) = "" Then
				    If Request("AssignedPDA") = "Y" or Request("AssignedPDA") = "2" Then
				        rs("AssignedPDA") = True
				    Else
				        rs("AssignedPDA") = False
				    End If
					If Err.Number <> 0 Then
						HeaderMSG = "The value provided for Assigned PDA is invalid."
						WOAssignRec
					End If
				End If

				On Error Goto 0
			End If
			db.dobatchupdate rs
			If db.dok Then
				'pl = rs("pk")
			Else
				If InStr(db.derror,"NULL") > 0 Then
					HeaderMSG = "A value must be provided for Labor ID."
					WOAssignRec
				Else
					Call OutputWAPError(db.derror)
				End If
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub WOAssign()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo
	Dim CheckBoxType

	CardTitle = "WO Assignments"
	CardCurrent = "WOAssign"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = 5
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = 5
    Else
	    pagesize = GlobalPageSize
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call WOAssignRecSave(db)

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		woid = "Unknown"
		reason = ""
		photo = ""
		assetpk = ""
		assetid = ""
		assetname = ""
		isopen = True
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
	End If

	sql = _
	"SELECT WOassign.PK, WOassign.IsAssigned, WOassign.LaborPK, WOassign.LaborID, " &_
	"       WOassign.LaborName, Labor.CraftID, Labor.CraftName, Labor.LaborType, Labor.LaborTypeDesc, WOassign.AssignedHours, WOassign.AssignedDate, Labor.Email, Labor.PagerEmail, Labor.Photo, " &_
	"       WOassign.RowVersionDate, WOassign.AssignedLead, WOassign.AssignedPDA " &_
	"FROM WOassign WITH (NOLOCK) " &_
	"     LEFT OUTER JOIN Labor Labor WITH (NOLOCK) ON WOassign.LaborPK = Labor.LaborPK " &_
	"WHERE (WOassign.WOPK = " & WOPK & ") AND (WOassign.IsAssigned = 1) AND (WOassign.Active = 1 or WOassign.Active Is Null) " &_
	"ORDER BY WOassign.IsAssigned, Labor.LaborType DESC, WOassign.AssignedDate, WOassign.LaborName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

    If rs.Eof and IsOpen Then
		If Not UCase(CardFrom) = "WOASSIGNREC" Then
		    If cardlevelupdownamount <> 0 Then
        		CardSkipLevel = 1
                WOAssignRec
            End If
        Else
            CardSkipLevel = -1
            Call WOOptions("")
        End If
    End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and False Then
			Call OutputWAPMsg("No Assignments were Found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap"><%
			If Not IsBOF Then %>
			<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;pr=<% =WOPK %>">&lt;</a>&nbsp;<%
			End If %><b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">WO #<% =WOID %></a></b><%
			If Not IsEOF Then %>
			<a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;nr=<% =WOPK %>">&gt;</a><%
			End If %>
			<br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
		    <% If lang = "WML" and IsPPC Then
		    Response.Write "<br/>"
		    Response.Write HStyleBegin
		    OutputBackButton
		    OutputHomeButton
		    Response.Write HStyleEnd
		    End If %>
			</p>
			<% End If %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else %>
				<% =LStyleBegin %><a href="default.asp?card=woassignrec&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>">[<% =WAPValidate(NullCheck(rs("LaborName"))) %>] [<% =DateNullCheck(RS("AssignedDate")) %>] [<% =NullCheck(RS("AssignedHours")) %>&nbsp;Hour(s)]</a><% =LStyleEnd %><%
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
			        Response.Write "<br/>"
			    End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>">Next</a></b>
		<%
		End If
		If IsOpen Then %>
		&nbsp;<b><a href="default.asp?card=WOAssignRec&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=-1">Add</a></b><%
		End If %>
		</p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOIssue()
	Dim HeaderTitle
	CardTitle = "WO Issue"
	HeaderTitle = "Issue WO"
	CardCurrent = "WOIssue"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WOOnHold()
	Dim HeaderTitle
	CardTitle = WAPValidate("WO On-Hold")
	HeaderTitle = WAPValidate("Place WO On-Hold")
	CardCurrent = "WOOnHold"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WORespond()
	Dim HeaderTitle
	CardTitle = "WO Respond"
	HeaderTitle = "Respond to WO"
	CardCurrent = "WORespond"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WOComplete()
	Dim HeaderTitle
	CardTitle = "WO Complete"
	HeaderTitle = "Complete WO"
	CardCurrent = "WOComplete"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WOClose()
	Dim HeaderTitle
	CardTitle = "WO Close"
	HeaderTitle = "Close WO"
	CardCurrent = "WOClose"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WOStatusProcess(headertitle)

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec

    newcontextpage = True

	CardCurrentLevel = GetCardLevel()

	GetWOPK

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	sql = _
	"SELECT WO.*, Asset.Photo as AssetPhoto, Asset.IsMeter, Asset.Meter1Reading As A_Meter1Reading, Asset.Meter2Reading As A_Meter2Reading " &_
	"FROM WO WITH (NOLOCK) LEFT OUTER JOIN ASSET WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " &_
	"WHERE WOPK = " & WOPK & " "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.eof Then
		Call OutputWAPMsg("The Work Order was not found")
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35))
		photo = NullCheck(rs("AssetPhoto"))
		assetpk = NullCheck(rs("AssetPK"))
		assetid = WAPValidate(NullCheck(RS("AssetID")))
		assetname = WAPValidate(NullCheck(RS("AssetName")))
		isopen = BitNullCheck(RS("IsOpen"))
	End If

	Dim RS_WOClosePref
	Dim WO_CLOSE_ACCOUNTSETALL
	Dim WO_CLOSE_ALLTASKSCOMPLETE
	Dim WO_CLOSE_CATEGORYSETALL
	Dim WO_CLOSE_CHARGEABLE
	Dim WO_CLOSE_CLOSE
	Dim WO_CLOSE_COMPLETE
	Dim WO_CLOSE_FINALIZE
	Dim WO_CLOSE_LABORHOURSASN
	Dim WO_CLOSE_LABORHOURSEST
	Dim WO_CLOSE_LABORREPORT
	Dim WO_CLOSE_MATERIALEST
	Dim WO_CLOSE_OTHEREST
	Dim WO_CLOSE_RESPOND
	Dim WO_CLOSE_RETURNTOSERVICE
	Dim WO_CLOSE_SETDOWNTIME

	Set RS_WOClosePref = db.runSPReturnRS("MC_GetWorkOrderClosePrefs",Array(Array("@LaborPK", adInteger, adParamInput, 4, GetSession("USERPK")),Array("@RepairCenterPK", adInteger, adParamInput, 4, GetSession("RCPK"))),"")
	Call CheckDB(db)

	If Not RS_WOClosePref.Eof Then
		WO_CLOSE_ACCOUNTSETALL = RS_WOClosePref("WO_CLOSE_ACCOUNTSETALL")
		WO_CLOSE_ALLTASKSCOMPLETE = RS_WOClosePref("WO_CLOSE_ALLTASKSCOMPLETE")
		WO_CLOSE_CATEGORYSETALL = RS_WOClosePref("WO_CLOSE_CATEGORYSETALL")
		WO_CLOSE_CHARGEABLE = RS_WOClosePref("WO_CLOSE_CHARGEABLE")
		WO_CLOSE_CLOSE = RS_WOClosePref("WO_CLOSE_CLOSE")
		WO_CLOSE_COMPLETE = RS_WOClosePref("WO_CLOSE_COMPLETE")
		WO_CLOSE_FINALIZE = RS_WOClosePref("WO_CLOSE_FINALIZE")
		WO_CLOSE_LABORHOURSASN = RS_WOClosePref("WO_CLOSE_LABORHOURSASN")
		WO_CLOSE_LABORHOURSEST = RS_WOClosePref("WO_CLOSE_LABORHOURSEST")
		WO_CLOSE_LABORREPORT = NullCheck(RS_WOClosePref("WO_CLOSE_LABORREPORT"))
		WO_CLOSE_MATERIALEST = RS_WOClosePref("WO_CLOSE_MATERIALEST")
		WO_CLOSE_OTHEREST = RS_WOClosePref("WO_CLOSE_OTHEREST")
		WO_CLOSE_RESPOND = RS_WOClosePref("WO_CLOSE_RESPOND")
		WO_CLOSE_RETURNTOSERVICE = RS_WOClosePref("WO_CLOSE_RETURNTOSERVICE")
		WO_CLOSE_SETDOWNTIME = RS_WOClosePref("WO_CLOSE_SETDOWNTIME")
	End If

	CloseObj RS_WOClosePref

	Dim LaborReport
	If NullCheck(rs("LaborReport")) = "" Then
		LaborReport = WO_CLOSE_LABORREPORT
	Else
		laborreport = Replace(NullCheck(rs("laborreport")),"%0D%0A"," ")
	End If

	Dim Chargeable
	If BitNullCheck(rs("Chargeable")) Then
		Chargeable = True
	Else
		Chargeable = WO_CLOSE_CHARGEABLE
	End If

	Dim FailedWO
	If BitNullCheck(rs("FailedWO")) Then
		FailedWO = True
	Else
		FailedWO = False
	End If

	Dim txtmeter1reading, txtmeter2reading
	If NullCheck(rs("meter1reading")) = "" or NullCheck(rs("meter1reading")) = "0" Then
		txtmeter1reading = NullCheck(rs("a_meter1reading"))
	Else
		txtmeter1reading = NullCheck(rs("meter1reading"))
	End If
	If NullCheck(rs("meter2reading")) = "" or NullCheck(rs("meter2reading")) = "0" Then
		txtmeter2reading = NullCheck(rs("a_meter2reading"))
	Else
		txtmeter2reading = NullCheck(rs("meter2reading"))
	End If

	If Not UCase(card) = "WOISSUE" Then
    	Call BuildFields("LaborReport","Report","C",GlobalFieldLength,"*M","true",WO_CLOSE_LABORREPORT,rs,"-1")
    End If
	If Not UCase(card) = "WOONHOLD" and Not UCase(card) = "WOISSUE" Then
		Call BuildFields("RegularHours","Hour(s)","C",GlobalFieldLength,"*M","true","",rs,"-1")
	End If
	Call BuildFields("Date","Date","C",GlobalFieldLength,"*M","true",CStr(DateNullCheck(Date())),rs,"-1")
	Call BuildFields("Time","Time","C",GlobalFieldLength,"*M","true",TimeNullCheck(Now()),rs,"-1")
	If Not UCase(card) = "WOONHOLD" and Not UCase(card) = "WOISSUE" Then
		Call BuildFields("AccountID","Account ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("AccountID")),rs,"-1")
		Call BuildFields("CategoryID","Category ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("CategoryID")),rs,"-1")
		If IsPocketIE or IsBlackBerry or IsWAP20 Then
			Call BuildFields("ProblemID","Problem ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("ProblemID")),rs,"-1")
		End If
		Call BuildFields("FailureID","Failure ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("FailureID")),rs,"-1")
		If IsPocketIE or IsBlackBerry or IsWAP20 Then
			Call BuildFields("SolutionID","Solution ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("SolutionID")),rs,"-1")
		End If
		If BitNullCheck(rs("IsMeter")) Then
			Call BuildFields("Meter1Reading","Meter 1","C",GlobalFieldLength,"*M","true",txtmeter1reading,rs,"-1")
			If IsPocketIE or IsBlackBerry or IsWAP20 Then
				Call BuildFields("Meter2Reading","Meter 2","C",GlobalFieldLength,"*M","true",txtmeter2reading,rs,"-1")
			End If
		End If
		Call BuildFields("Failed","WO Failed?","B","","","",FailedWO,rs,"-1")
		Call BuildFields("TasksComplete","All Tasks Complete?","B","","","",WO_CLOSE_ALLTASKSCOMPLETE,rs,"-1")
		Call BuildFields("Chargeable","Chargeable?","B","","","",Chargeable,rs,"-1")
	End If

	Call StartMobileDocument(CardTitle)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("The WO was not found")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2">WO #<% =WOID %></a></b><br/>
			<% If Not AssetPK = "" Then %>
			<% =HStyleBegin %><a href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a><% =HStyleEnd %><br/>
			<% End If %>
			<% =HStyleBegin %><% =Reason %><% =HStyleEnd %>
			<br/><% =HStyleBegin %><b><% =HeaderTitle %></b><% =HStyleEnd %>
			</p><%
			End If
			OutputFields
		End If
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			If IsOpen Then
			WOCloseSubmit rs
			End If
		Else
			If IsOpen Then
			WOCloseSubmit rs
			End If
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		If IsOpen Then
		%>
		</td><td align="right">
		<%
		WOCloseSubmit rs
		End If
		HTMLbuttonsEnd
		End If
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WOCloseSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<% If UCase(card) = "WOONHOLD" or UCase(card) = "WORESPOND" or UCase(card) = "WOISSUE" Then %>
			<postfield name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
			<% Else %>
			<postfield name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-2)) %>"/>
			<% End If %>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="wopk" value="<% =WOPK %>"/>
			<postfield name="woaction" value="<% =card %>"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<% If UCase(card) = "WOONHOLD" or UCase(card) = "WORESPOND" or UCase(card) = "WOISSUE" Then %>
	<input type="hidden" name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
	<% Else %>
	<input type="hidden" name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-2)) %>"/>
	<% End If %>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wopk" value="<% =WOPK %>"/>
	<input type="hidden" name="woaction" value="<% =card %>"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub AssetMenu()

	Dim db, sql, rs

	CardTitle = "AssetMenu"
	CardCurrent = "Asset Menu"
	CardCurrentLevel = GetCardLevel()

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call StartMobileDocument(CardTitle)
		If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap">
		<b>Asset Menu</b>
		</p><%
		End If
        %>
        <p align="center">
   		</p>
        <%
		If lang = "WML" Then
		    WAPbuttonsBegin
		    OutputBackButton
		    WAPbuttonsEnd
		Else
		    HTMLbuttonsBegin
		    OutputBackButton
		    %>
		    </td><td align="right">
		    <%
		    HTMLbuttonsEnd
		End If
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub InventoryMenu()

	Dim db, sql, rs

	CardTitle = "Inventory"
	CardCurrent = "Inventory"
	CardCurrentLevel = GetCardLevel()

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call StartMobileDocument(CardTitle)
		If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap">
		<b>Inventory Menu</b>
		</p><%
		End If
        %>
        <p align="center">
        <a href="default.asp?card=viewinventory&amp;s=<% =SessionID %>">View Inventory</a><br/>
        <a href="default.asp?card=newitem&amp;s=<% =SessionID %>">New Inventory Item</a><br/>
   		<a href="default.asp?card=adjustinventory&amp;s=<% =SessionID %>">Adjust Inventory</a><br/>
   		<a href="default.asp?card=countinventory&amp;s=<% =SessionID %>">Count Inventory</a><br/>
   		</p>
        <%
		If lang = "WML" Then
		    WAPbuttonsBegin
		    OutputBackButton
		    WAPbuttonsEnd
		Else
		    HTMLbuttonsBegin
		    OutputBackButton
		    %>
		    </td><td align="right">
		    <%
		    HTMLbuttonsEnd
		End If
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WONew()

	Dim db, sql, rs

    newcontextpage = True

	CardTitle = "New Work Order"
	CardCurrent = "WONew"
	CardCurrentLevel = GetCardLevel()

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Dim prefvalue, prefdesc, prefpk

	Call BuildFields("Reason","Reason","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Call BuildFields("AssetID","Asset ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Call BuildFields("ProblemID","Problem ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Call BuildFields("ProcedureID","Procedure ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Call BuildFields("TargetDate","Target Date","C",GlobalFieldLength,"*M","true",CStr(DateNullCheck(Date())),rs,"-1")
	If GetPreference(db,False,GetSession("RCPK"),"WO_DefaultTargetHours",prefvalue, prefdesc, prefpk) Then
	    Call BuildFields("TargetHours","Target Hours","C",GlobalFieldLength,"*M","true",WAPValidate(prefvalue),rs,"-1")
	Else
	    Call BuildFields("TargetHours","Target Hours","C",GlobalFieldLength,"*M","true","",rs,"-1")
	End If
	If GetPreference(db,False,GetSession("RCPK"),"WO_DefaultPriority",prefvalue, prefdesc, prefpk) Then
    	Call BuildFields("Priority","Priority","C",GlobalFieldLength,"*M","true",WAPValidate(prefvalue),rs,"-1")
	Else
    	Call BuildFields("Priority","Priority","C",GlobalFieldLength,"*M","true","",rs,"-1")
	End If
	If GetPreference(db,False,GetSession("RCPK"),"WO_DefaultType",prefvalue, prefdesc, prefpk) Then
    	Call BuildFields("Type","Type","C",GlobalFieldLength,"*M","true",WAPValidate(prefvalue),rs,"-1")
	Else
    	Call BuildFields("Type","Type","C",GlobalFieldLength,"*M","true","",rs,"-1")
	End If
	Call BuildFields("AccountID","Account ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Call BuildFields("CategoryID","Category ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Call BuildFields("DepartmentID","Department ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Call BuildFields("TenantID","Customer ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Call BuildFields("ProjectID","Project ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	If False Then
	' Do not include RC because the user would not be able to select a Shop in that RC
	' unless the Shop Lookup was redesigned.
	Call BuildFields("RepairCenterID","Repair Center ID","C",GlobalFieldLength,"*M","true",GetSession("RCID"),rs,"-1")
	End If
	If GetSession("SHPK") = "" Then
		Call BuildFields("ShopID","Shop ID","C",GlobalFieldLength,"*M","true","",rs,"-1")
	Else
		Call BuildFields("ShopID","Shop ID","C",GlobalFieldLength,"*M","true",GetSession("SHID"),rs,"-1")
	End If
	'Call BuildFields("RequesterID","Requested By","C",GlobalFieldLength,"*M","true",GetSession("UserID"),rs,"-1")
	'Call BuildFields("RequesterName","Requester Name","C",GlobalFieldLength,"*M","true",GetSession("UserName"),rs,"-1")
	'Call BuildFields("RequesterPhone","Requester Phone","C",GlobalFieldLength,"*M","true",GetSession("UserPhone"),rs,"-1")
	'Call BuildFields("RequesterEmail","Requester Email","C",GlobalFieldLength,"*M","true",GetSession("UserEmail"),rs,"-1")
	Call BuildFields("LaborID","Assign To","C",GlobalFieldLength,"*M","true",GetSession("UserID"),rs,"-1")

	Call BuildFields("Chargeable","Chargeable?","B","","","",False,rs,"-1")
	Call BuildFields("FollowupWork","Follow-up Work?","B","","","",False,rs,"-1")
	Call BuildFields("ShutdownBox","Shutdown Required?","B","","","",False,rs,"-1")
	Call BuildFields("LockoutTagoutBox","Lockout/Tagout?","B","","","",False,rs,"-1")
	'Call BuildFields("ViewWOAfterSave","View after Submit?","B","","","",False,rs,"-1")

	Call StartMobileDocument(CardTitle)
		If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap">
		<b>New Work Order</b>
		</p><%
		End If
		OutputFields
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			WONewSubmit rs
		Else
			WONewSubmit rs
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		WONewSubmit rs
		HTMLbuttonsEnd
		End If
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WONewSubmit(rs)
	If lang = "WML" Then
	Call OutputButtonOrLinkStart("","") %>
		<go href="default.asp" method="get">
			<postfield name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<postfield name="wonew" value="y"/>
			<postfield name="preventduplicatesubmit" value="<% =RandomString(7) %>"/>
			<% =PostFields %>
		</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wonew" value="y"/>
	<input type="hidden" name="preventduplicationsubmit" value="<% =RandomString(7) %>"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub CountInventory()

	Dim db, sql, rs, LocationID

    newcontextpage = True

	If Not Request("OnHandPending") = "" or Not Request("Bin") = "" Then
	    CardCurrent = "CountInventory2"
    ElseIf Not Request("Posted") = "" Then
	    CardCurrent = "CountInventory2"
    Else
	    CardCurrent = "CountInventory"
    End If
	CardTitle = "Count Inventory"
	CardCurrentLevel = GetCardLevel()

	Set db = New ADOHelper

	If Not Request("pk") = "" Then
        Call InventoryCountSave(db)
        If HeaderMSG = "" Then
            CardCurrent = "CountInventory"
            CardCurrentLevel = GetCardLevel()
	    End If
    Else
        If Not Request("PartID") = "" and Not Request("LocationID") = "" Then
            ' Validate PartID and LocationID - and if they are in PartLocation together
            ' Does PartID Exist?
            sql = "SELECT PartPK FROM Part WITH (NOLOCK) WHERE PartID = '" & Request("PartID") & "'"
            Set rs = db.RunSQLReturnRS(sql,"")
	        Call CheckDB(db)
	        If rs.Eof Then
	            HeaderMsg = "Part ID Does Not Exist."
	        End If
	        If HeaderMsg = "" Then
                sql = "SELECT LocationPK FROM Location WITH (NOLOCK) WHERE LocationID = '" & Request("LocationID") & "'"
                Set rs = db.RunSQLReturnRS(sql,"")
	            Call CheckDB(db)
	            If rs.Eof Then
    	            HeaderMsg = "Location ID Does Not Exist."
	            End If
	        End If
	        If HeaderMsg = "" Then
	            sql = _
	            "SELECT PartLocation.PK " &_
	            "FROM PartLocation WITH (NOLOCK) " & _
	            "INNER JOIN Part WITH (NOLOCK) ON Part.PartPK = PartLocation.PartPK " &_
	            "INNER JOIN Location WITH (NOLOCK) ON Location.LocationPK = PartLocation.LocationPK " &_
	            "WHERE Part.PartID = '" & Request("PartID") & "' AND Location.LocationID = '" & Request("LocationID") & "' "
                Set rs = db.RunSQLReturnRS(sql,"")
	            Call CheckDB(db)
	            If rs.Eof Then
    	            HeaderMsg = "The Part ID specified is not stocked at the Location ID specified."
	            End If
	        End If
            If Not HeaderMSG = "" Then
                CardCurrent = "CountInventory"
                CardCurrentLevel = GetCardLevel()
	        End If
	    Else
	        If CardCurrent = "CountInventory2" Then
	            HeaderMSG = "You must specify both Part ID and Location ID."
                CardCurrent = "CountInventory"
                CardCurrentLevel = GetCardLevel()
            End If
        End If
    End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

    If HeaderMsg = "" and Not Request("PartID") = "" and Not Request("LocationID") = "" Then

   	    Call SetSession("LocationID" & CardCurrentLevel - 1,Request("LocationID"))

    Else

        If CardCurrent = "CountInventory" Then

	        LocationID = GetSession("LocationID" & CardCurrentLevel)
	        If LocationID = "" Then
                LocationID = GetSession("SRID")
	        End If

	        Call BuildFields("PartID","Item #","C",GlobalFieldLength,"*M","true","",rs,-1)
	        Call BuildFields("LocationID","Location ID","C",GlobalFieldLength,"*M","true",LocationID,rs,-1)

        End If

    End If

    pk = ""

    If CardCurrent = "CountInventory2" Then
        If (Not HeaderMSG = "") or (Not Request("PartID") = "" and Not Request("LocationID") = "") Then

            If (Not HeaderMSG = "") Then
	            sql = _
	            "SELECT PartLocation.*, Part.PartID, Part.PartName, Location.LocationID, Location.LocationName " &_
	            "FROM PartLocation WITH (NOLOCK) " & _
	            "INNER JOIN Part WITH (NOLOCK) ON Part.PartPK = PartLocation.PartPK " &_
	            "INNER JOIN Location WITH (NOLOCK) ON Location.LocationPK = PartLocation.LocationPK " &_
	            "WHERE PartLocation.PK = " & Request("PK")
            Else
	            sql = _
	            "SELECT PartLocation.*, Part.PartID, Part.PartName, Location.LocationID, Location.LocationName " &_
	            "FROM PartLocation WITH (NOLOCK) " & _
	            "INNER JOIN Part WITH (NOLOCK) ON Part.PartPK = PartLocation.PartPK " &_
	            "INNER JOIN Location WITH (NOLOCK) ON Location.LocationPK = PartLocation.LocationPK " &_
	            "WHERE Part.PartID = '" & Request("PartID") & "' AND Location.LocationID = '" & Request("LocationID") & "' "
            End If

	        Set rs = db.RunSQLReturnRS(sql,"")
	        Call CheckDB(db)

            If rs.Eof Then
                ' The Part does not exist in this location or it is an invalid partid / locationid

            Else
    	        Call BuildFields("Bin","Bin","C",GlobalFieldLength,"*M","true","",rs,rs("PK"))
    	        Call BuildFields("OnHandPending","New Count","C",GlobalFieldLength,"*M","true","",rs,rs("PK"))

                pk = rs("pk")
            End If

        End If
    End If

	Call StartMobileDocument(CardTitle)
		If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap">
		<b>Count Inventory</b>
		</p><%
		End If

        If Not CardCurrent = "CountInventory" Then
            If (Not HeaderMSG = "") or (Not Request("PartID") = "" and Not Request("LocationID") = "") Then
            Response.Write "<p mode=""nowrap"">" & HStyleBegin & rs("PartID") & " (" & rs("PartName") & ")" & HStyleEnd & "<br/>"
            Response.Write HStyleBegin & rs("LocationID") & " (" & rs("LocationName") & ")" & HStyleEnd & "<br/><br/>"
            Response.Write "<b>On-Hand: " & rs("OnHand") & "</b>"
            Response.Write "</p>"
            End If
        End If

		OutputFields
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			CountInventorySubmit rs
		Else
			CountInventorySubmit rs
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		CountInventorySubmit rs
		HTMLbuttonsEnd
		End If
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub CountInventorySubmit(rs)
	If lang = "WML" Then
    Call OutputButtonOrLinkStart("","") %>
    <go href="default.asp" method="get">
			<postfield name="card" value="CountInventory"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<% If Not pk = "" Then %>
			<postfield name="pk" value="<% =pk %>"/>
			<% End If %>
			<postfield name="POSTED" value="Y"/>
			<% =PostFields %>
	</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="CountInventory"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="pk" value="<% =pk %>"/>
	<input type="hidden" name="posted" value="Y"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub InventoryCountSave(ByRef db)
	Dim rs2, WOAssignPK
	pk = Request("pk")
	If Not pk = "" Then
		sql = "SELECT * FROM PartLocation WITH (NOLOCK) WHERE PK = " & pk
		Set rs = db.RunSQLReturnRS_RW(sql,"")
		If Not rs.Eof Then
			On Error Resume Next

            If Not Request("OnHandPending") = "" Then
                rs("OnHandPending") = Request("OnHandPending")
                'RGJ 9/4/2008 START - WO# 16045
                rs("PhysicalLast") = Date
                'RGJ 9/4/2008 END
            Else
                rs("OnHandPending") = Null
            End If
			If Err.Number <> 0 Then
				HeaderMSG = "The value provided for New Count is invalid."
				Exit Sub
			End If

            If Not Request("Bin") = "" Then
                rs("Bin") = Request("Bin")
		    Else
		        rs("Bin") = Null
            End If
			If Err.Number <> 0 Then
				HeaderMSG = "The value provided for Bin is invalid."
				Exit Sub
			End If

			On Error Goto 0

			db.dobatchupdate rs
			If Not db.dok Then
				Call OutputWAPError(db.derror)
			End If
			rs.close
		End If
	End If

End Sub

'====================================================================================================================================

Sub AssetTasks()

	Dim db, sql, rs, AssetID

    newcontextpage = True

    CardCurrent = "AssetTasks"
	CardTitle = "Asset Tasks"
	CardCurrentLevel = GetCardLevel()

	Set db = New ADOHelper

	If Request("Posted") = "Y" AND Not Request("AssetID") = "" Then
		' Validate AssetID
		sql = "SELECT AssetPK FROM Asset WITH (NOLOCK) WHERE AssetID = '" & Request("AssetID") & "'"
		Set rs = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs.Eof Then
			HeaderMsg = "Asset ID Does Not Exist."
		Else
			Call SetSession("ParentCard2",CardCurrent)
			Call SetSession("CardFrom",CardCurrent)
			Call SetSession("CardFromLevel",CardCurrentLevel)
			Call SetSession("AssetPK2",rs("AssetPK"))
			CardFromLevel = CardFromLevel + 1
			WOTasks
		End If
	Else
		If Not Request("Posted") = "" Then
			HeaderMSG = "You must specify Asset ID."
		End If
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

    AssetID = ""
    Call BuildFields("AssetID","Asset ID","C",GlobalFieldLength,"*M","true","",rs,-1)

    pk = ""

	Call StartMobileDocument(CardTitle)
		If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap">
		<b>Asset Tasks</b>
		</p><%
		End If

		OutputFields
		If lang = "WML" Then
		WAPbuttonsBegin
		If IsPocketIE Then
			OutputBackButton
			AssetTasksSubmit rs
		Else
			AssetTasksSubmit rs
			OutputBackButton
		End If
		WAPbuttonsEnd
		Else
		HTMLbuttonsBegin
		OutputBackButton
		%>
		</td><td align="right">
		<%
		AssetTasksSubmit rs
		HTMLbuttonsEnd
		End If
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub AssetTasksSubmit(rs)
	If lang = "WML" Then
    Call OutputButtonOrLinkStart("","") %>
    <go href="default.asp" method="get">
			<postfield name="card" value="AssetTasks"/>
			<postfield name="s" value="<% =SessionID %>"/>
			<% If Not pk = "" Then %>
			<postfield name="pk" value="<% =pk %>"/>
			<% End If %>
			<postfield name="POSTED" value="Y"/>
			<% =PostFields %>
	</go><%
	Call OutputButtonOrLinkEnd("")
	Else
	%>
	<input style="width:100%;" type="submit" name="submit" value="Submit"/>
	<input type="hidden" name="card" value="AssetTasks"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="pk" value="<% =pk %>"/>
	<input type="hidden" name="posted" value="Y"/>
	<%
	End If
End Sub

'====================================================================================================================================

Sub WOCloseSave(ByRef db)

	Dim requested, requesteddate, requestedtime, requestedinitials
	Dim issued, issueddate, issuedtime, issuedinitials
	Dim responded, respondeddate, respondedtime, respondedinitials
	Dim completed, completeddate, completedtime, completedinitials
	Dim finalized, finalizeddate, finalizedtime, finalizedinitials
	Dim closed, closeddate, closedtime, closedinitials

	Dim laborreport
	Dim lri
	Dim txtwogrouptype
	Dim txtaccountpk
	Dim txtaccount
	Dim txtaccountdesc
	Dim txtchargeable
	Dim txtAccountAll
	Dim txtcategory
	Dim txtcategorydesc
	Dim txtcategorypk
	Dim txtCategoryAll
	Dim txtTasks
	Dim txtTaskInitials
	Dim txtLabor1
	Dim txtLabor3
	Dim txtMaterials
	Dim txtOtherCost
	Dim txtproblempk
	Dim txtproblem
	Dim txtproblemdesc
	Dim txtfailurepk
	Dim txtfailure
	Dim txtfailuredesc
	Dim txtsolutionpk
	Dim txtsolution
	Dim txtsolutiondesc
	Dim txtfailurewo
	Dim txtmeter1reading
	Dim txtmeter2reading
	Dim txtisup
	Dim txtDrawingUpdatesNeeded
	Dim rs2
	Dim woaction
	Dim mode
	Dim completeassignments
	Dim txtDate,txtTime

	woaction = Request("woaction")
	If Not woaction = "" Then

		sql = "SELECT * FROM WO WITH (NOLOCK) WHERE WOPK = " & WOPK
		Set rs = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)

		If Not rs.Eof Then

			laborreport = NullCheck(Request("LaborReport"))

			If Not NullCheck(Request("AccountID")) = "" Then
				sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
				Set rs2 = db.RunSQLReturnRS(sql,"")
				Call CheckDB(db)
				If rs2.eof Then
					HeaderMSG = "The Account ID was not found."
					If Not UCase(card) = "WOOPTIONS" Then
						CardSkipLevel = 1
					End If
					Card = WOAction
					WOClose
				Else
					txtaccountpk = rs2("AccountPK")
					txtaccount = rs2("AccountID")
					txtaccountdesc = rs2("AccountName")
				End If
			Else
				txtaccountpk = Null
				txtaccount = Null
				txtaccountdesc = Null
			End If
			If Not NullCheck(Request("CategoryID")) = "" Then
				sql = "SELECT * FROM Category WITH (NOLOCK) WHERE CategoryID = '" & NullCheck(Request("CategoryID")) & "' "
				Set rs2 = db.RunSQLReturnRS(sql,"")
				Call CheckDB(db)
				If rs2.eof Then
					HeaderMSG = "The Category ID was not found."
					If Not UCase(card) = "WOOPTIONS" Then
						CardSkipLevel = 1
					End If
					Card = WOAction
					WOClose
				Else
					txtcategorypk = rs2("CategoryPK")
					txtcategory = rs2("CategoryID")
					txtcategorydesc = rs2("CategoryName")
				End If
			Else
				txtcategorypk = Null
				txtcategory = Null
				txtcategorydesc = Null
			End If
			If Not NullCheck(Request("ProblemID")) = "" Then
				sql = "SELECT * FROM Failure WITH (NOLOCK) WHERE FailureID = '" & NullCheck(Request("ProblemID")) & "' "
				Set rs2 = db.RunSQLReturnRS(sql,"")
				Call CheckDB(db)
				If rs2.eof Then
					HeaderMSG = "The Problem ID was not found."
					If Not UCase(card) = "WOOPTIONS" Then
						CardSkipLevel = 1
					End If
					Card = WOAction
					WOClose
				Else
					txtProblempk = rs2("FailurePK")
					txtProblem = rs2("FailureID")
					txtProblemdesc = rs2("FailureName")
				End If
			Else
				txtProblempk = Null
				txtProblem = Null
				txtProblemdesc = Null
			End If
			If Not NullCheck(Request("FailureID")) = "" Then
				sql = "SELECT * FROM Failure WITH (NOLOCK) WHERE FailureID = '" & NullCheck(Request("FailureID")) & "' "
				Set rs2 = db.RunSQLReturnRS(sql,"")
				Call CheckDB(db)
				If rs2.eof Then
					HeaderMSG = "The Failure ID was not found."
					If Not UCase(card) = "WOOPTIONS" Then
						CardSkipLevel = 1
					End If
					Card = WOAction
					WOClose
				Else
					txtFailurepk = rs2("FailurePK")
					txtFailure = rs2("FailureID")
					txtFailuredesc = rs2("FailureName")
				End If
			Else
				txtFailurepk = Null
				txtFailure = Null
				txtFailuredesc = Null
			End If
			If Not NullCheck(Request("SolutionID")) = "" Then
				sql = "SELECT * FROM Failure WITH (NOLOCK) WHERE FailureID = '" & NullCheck(Request("SolutionID")) & "' "
				Set rs2 = db.RunSQLReturnRS(sql,"")
				Call CheckDB(db)
				If rs2.eof Then
					HeaderMSG = "The Solution ID was not found."
					If Not UCase(card) = "WOOPTIONS" Then
						CardSkipLevel = 1
					End If
					Card = WOAction
					WOClose
				Else
					txtSolutionpk = rs2("FailurePK")
					txtSolution = rs2("FailureID")
					txtSolutiondesc = rs2("FailureName")
				End If
			Else
				txtSolutionpk = Null
				txtSolution = Null
				txtSolutiondesc = Null
			End If
			If Request("date") = "" Then
				txtdate = DateNullCheck(Date())
			Else
				If Not IsDate(Request("date")) Then
					HeaderMSG = "The value for Date can not be blank."
					If Not UCase(card) = "WOOPTIONS" Then
						CardSkipLevel = 1
					End If
					Card = WOAction
					WOClose
				End If
				txtdate = Request("date")
			End If
			If Request("time") = "" Then
				txttime = TimeNullCheck(Now())
			Else
				txttime = Request("time")
			End If

			Dim actionwhere

			' Status changed to Issue?
			If UCase(woaction) = "WOISSUE" Then
				actionwhere = " WHERE WO.WOPK = " & WOPK
				sql = _
				"UPDATE WO" & _
				"	SET IsOpen = 1, Status = 'ISSUED', StatusDesc = 'Issued', ClosedInitials = Null, Closed = Null, StatusDate = '" & SQLdatetimeADO(txtdate & " " & txttime) & "', RowVersionUserPK = " & GetSession("UserPK") & ", RowVersionInitials = '" & GetSession("UserInitials") & "', RowVersionAction = 'STATUS', RowVersionDate = getdate() " & _
				actionwhere & " OR (WOGroupPK IN (SELECT WOGroupPK FROM WO WITH (NOLOCK) " & actionwhere & " AND WOGroupType = 'M')) AND WO.PDAWO = 0 "
				Call db.RunSQL(sql,"")
				Call CheckDB(db)
				Exit Sub
			End If

			' Status changed to On-Hold?
			If UCase(woaction) = "WOONHOLD" Then
				actionwhere = " WHERE WO.WOPK = " & WOPK
				sql = _
				"UPDATE WO" & _
				"	SET IsOpen = 1, LaborReport = LTRIM(IsNull(LaborReport,'') + ' " & SQLEncode(laborreport) & "'), Status = 'ONHOLD', StatusDesc = 'On-Hold', StatusDate = '" & SQLdatetimeADO(txtdate & " " & txttime) & "', RowVersionUserPK = " & GetSession("UserPK") & ", RowVersionInitials = '" & GetSession("UserInitials") & "', RowVersionAction = 'STATUS', RowVersionDate = getdate() " & _
				actionwhere & " OR (WOGroupPK IN (SELECT WOGroupPK FROM WO WITH (NOLOCK) " & actionwhere & " AND WOGroupType = 'M')) AND WO.PDAWO = 0 "
				Call db.RunSQL(sql,"")
				Call CheckDB(db)
				Exit Sub
			End If

			mode = "WO"
			completeassignments = 1

			requesteddate = DateNullCheck(rs("requested"))
			requestedtime = TimeNullCheck(rs("requested"))
			requestedinitials = NullCheck(rs("takenbyinitials"))
			If requesteddate = "" Then
				requesteddate = DateNullCheck(Date())
				requestedtime = FixTime(Time())
			End If

			issueddate = DateNullCheck(rs("issued"))
			issuedtime = TimeNullCheck(rs("issued"))
			issuedinitials = NullCheck(rs("issuedinitials"))
			If issueddate = "" Then
				issueddate = DateNullCheck(Date())
				issuedtime = FixTime(Time())
			End If
			If rs("status") = "REQUESTED" and issuedinitials = "" Then
				issuedinitials = GetSession("UserInitials")
			End If

			respondeddate = DateNullCheck(rs("responded"))
			respondedtime = TimeNullCheck(rs("responded"))
			respondedinitials = NullCheck(rs("respondedinitials"))
			If respondeddate = "" Then
				respondeddate = DateNullCheck(Date())
				respondedtime = FixTime(Time())
				responded = False
			Else
				responded = True
			End If

			completeddate = DateNullCheck(rs("complete"))
			completedtime = TimeNullCheck(rs("complete"))
			completedinitials = NullCheck(rs("completedinitials"))
			If completeddate = "" Then
				completeddate = DateNullCheck(Date())
				completedtime = FixTime(Time())
				completed = False
			Else
				completed = True
			End If

			finalizeddate = DateNullCheck(rs("finalized"))
			finalizedtime = TimeNullCheck(rs("finalized"))
			finalizedinitials = NullCheck(rs("finalizedinitials"))
			If finalizeddate = "" Then
				finalizeddate = DateNullCheck(Date())
				finalizedtime = FixTime(Time())
				finalized = False
			Else
				finalized = True
			End If

			closeddate = DateNullCheck(rs("closed"))
			closedtime = TimeNullCheck(rs("closed"))
			closedinitials = NullCheck(rs("closedinitials"))
			If closeddate = "" Then
				closeddate = DateNullCheck(Date())
				closedtime = FixTime(Time())
				closed = False
			Else
				closed = True
			End If

			Select Case UCase(woaction)
				Case "WORESPOND"
					respondeddate = DateNullCheck(txtdate)
					respondedtime = TimeNullCheck(txttime)
					respondedinitials = Trim(GetSession("UserInitials"))
					responded = True
				Case "WOCOMPLETE"
					completeddate = DateNullCheck(txtdate)
					completedtime = TimeNullCheck(txttime)
					completedinitials = Trim(GetSession("UserInitials"))
					respondeddate = DateNullCheck(rs("responded"))
					respondedtime = TimeNullCheck(rs("responded"))
					respondedinitials = NullCheck(rs("respondedinitials"))
					If respondeddate = "" Then
						respondeddate = completeddate
						respondedtime = completedtime
						respondedinitials = completedinitials
					End If
					sql = "Update WOAssignStatus SET completed = 1 WHERE WOPK = " & WOPK
					Call db.RunSQL(sql,"")
					Call CheckDB(db)
					responded = True
					completed = True
				Case "WOCLOSE"
					closeddate = DateNullCheck(txtdate)
					closedtime = TimeNullCheck(txttime)
					closedinitials = Trim(GetSession("UserInitials"))
					respondeddate = DateNullCheck(rs("responded"))
					respondedtime = TimeNullCheck(rs("responded"))
					respondedinitials = NullCheck(rs("respondedinitials"))
					If respondeddate = "" Then
						respondeddate = closeddate
						respondedtime = closedtime
						respondedinitials = closedinitials
					End If
					completeddate = DateNullCheck(rs("complete"))
					completedtime = TimeNullCheck(rs("complete"))
					completedinitials = NullCheck(rs("completedinitials"))
					If completeddate = "" Then
						completeddate = closeddate
						completedtime = closedtime
						completedinitials = closedinitials
					End If
					finalizeddate = DateNullCheck(rs("finalized"))
					finalizedtime = TimeNullCheck(rs("finalized"))
					finalizedinitials = NullCheck(rs("finalizedinitials"))
					If finalizeddate = "" Then
						finalizeddate = closeddate
						finalizedtime = closedtime
						finalizedinitials = closedinitials
					End If
					sql = "Update WOAssignStatus SET completed = 1 WHERE WOPK = " & WOPK
					Call db.RunSQL(sql,"")
					Call CheckDB(db)
					responded = True
					completed = True
					finalized = True
					closed = True
			End Select

			txtchargeable = UCase(Trim(Request("chargeable")))
			If txtchargeable = "Y" or txtchargeable = "2" Then
				txtchargeable = True
			Else
				txtchargeable = False
			End If

			txtTasks = UCase(Trim(Request("TasksComplete")))
			If txtTasks = "Y" or txtTasks = "2" Then
				txtTasks = True
			Else
				txtTasks = False
			End If

			txtTaskInitials = GetSession("UserInitials")

			txtfailurewo = UCase(Trim(Request("Failed")))
			If txtfailurewo = "Y" or txtfailurewo = "2" Then
				txtfailurewo = True
			Else
				txtfailurewo = False
			End If

			'txtLabor1 = Trim(Request("txtLabor1"))
			'If Not txtLabor1 = "" Then
			'	txtLabor1 = True
			'Else
				txtLabor1 = False
			'End If

			'txtLabor3 = Trim(Request("txtLabor3"))
			'If Not txtLabor3 = "" Then
			'	txtLabor3 = True
			'Else
				txtLabor3 = False
			'End If

			'txtMaterials = Trim(Request("txtMaterials"))
			'If Not txtMaterials = "" Then
			'	txtMaterials = True
			'Else
				txtMaterials = False
			'End If
			'txtOtherCost = Trim(Request("txtOtherCost"))
			'If Not txtOtherCost = "" Then
			'	txtOtherCost = True
			'Else
				txtOtherCost = False
			'End If

			txtmeter1reading = Trim(Request("meter1reading"))
			If txtmeter1reading = "" Then
				txtmeter1reading = 0
			End If
			txtmeter2reading = Trim(Request("meter2reading"))
			If txtmeter2reading = "" Then
				txtmeter2reading = 0
			End If

			'txtisup = Trim(Request("txtDownTime"))
			'If Not txtisup = "" Then
			'	txtisup = True
			'Else
				txtisup = False
			'End If

			'txtDrawingUpdatesNeeded = Trim(Request("txtDrawingUpdatesNeeded"))
			'If Not txtDrawingUpdatesNeeded = "" Then
			'	txtDrawingUpdatesNeeded = True
			'Else
				txtDrawingUpdatesNeeded = False
			'End If

			lri = 2

			' ************************************************************************************************************
			' CLOSE SINGLE WORK ORDER
			' ************************************************************************************************************

			If Not db.RunSP("MC_CloseWorkOrder", Array(_
				Array("@WOPK", adInteger, adParamInput, 4, WOPK),_
				Array("@requesteddate", adVarChar, adParamInput, 17, SQLdatetime(requesteddate & " " & requestedtime)),_
				Array("@requestedinitials", adChar, adParamInput, 5, requestedinitials),_
				Array("@issueddate", adVarChar, adParamInput, 17, SQLdatetime(issueddate & " " & issuedtime)),_
				Array("@issuedinitials", adChar, adParamInput, 5, issuedinitials),_
				Array("@responded", adBoolean, adParamInput, 1, responded),_
				Array("@respondeddate", adVarChar, adParamInput, 17, SQLdatetime(respondeddate & " " & respondedtime)),_
				Array("@respondedinitials", adChar, adParamInput, 5, respondedinitials),_
				Array("@completed", adBoolean, adParamInput, 1, completed),_
				Array("@completeddate", adVarChar, adParamInput, 17, SQLdatetime(completeddate & " " & completedtime)),_
				Array("@completedinitials", adChar, adParamInput, 5, completedinitials),_
				Array("@completeassignments", adBoolean, adParamInput, 1, completeassignments),_
				Array("@finalized", adBoolean, adParamInput, 1, finalized),_
				Array("@finalizeddate", adVarChar, adParamInput, 17, SQLdatetime(finalizeddate & " " & finalizedtime)),_
				Array("@finalizedinitials", adChar, adParamInput, 5, finalizedinitials),_
				Array("@closed", adBoolean, adParamInput, 1, closed),_
				Array("@closeddate", adVarChar, adParamInput, 17, SQLdatetime(closeddate & " " & closedtime)),_
				Array("@closedinitials", adChar, adParamInput, 5, closedinitials),_
				Array("@laborreport",  adVarchar, adParamInput, 3000, Trim(Mid(laborreport,1,3000))& " "),_
				Array("@lri", adSmallInt, adParamInput, 2, lri),_
				Array("@txtaccountpk", adInteger, adParamInput, 4, txtaccountpk),_
				Array("@txtaccount",  adVarchar, adParamInput, 25, Trim(Mid(txtaccount,1,25))),_
				Array("@txtaccountdesc",  adVarchar, adParamInput, 50, Trim(Mid(txtaccountdesc,1,50))),_
				Array("@txtAccountAll", adBoolean, adParamInput, 1, txtAccountAll),_
				Array("@txtchargeable", adBoolean, adParamInput, 1, txtchargeable),_
				Array("@txtcategorypk", adInteger, adParamInput, 4, txtcategorypk),_
				Array("@txtcategory",  adVarchar, adParamInput, 25, Trim(Mid(txtcategory,1,25))),_
				Array("@txtcategorydesc",  adVarchar, adParamInput, 50, Trim(Mid(txtcategorydesc,1,50))),_
				Array("@txtCategoryAll", adBoolean, adParamInput, 1, txtCategoryAll),_
				Array("@txtTasks", adBoolean, adParamInput, 1, txtTasks),_
				Array("@txtTaskInitials",  adVarchar, adParamInput, 5, Trim(Mid(txtTaskInitials,1,5))),_
				Array("@txtLabor1", adBoolean, adParamInput, 1, txtLabor1),_
				Array("@txtLabor3", adBoolean, adParamInput, 1, txtLabor3),_
				Array("@txtMaterials", adBoolean, adParamInput, 1, txtMaterials),_
				Array("@txtOtherCost", adBoolean, adParamInput, 1, txtOtherCost),_
				Array("@txtproblempk", adInteger, adParamInput, 4, txtproblempk),_
				Array("@txtproblem",  adVarchar, adParamInput, 25, Trim(Mid(txtproblem,1,25))),_
				Array("@txtproblemdesc",  adVarchar, adParamInput, 50, Trim(Mid(txtproblemdesc,1,50))),_
				Array("@txtfailurepk", adInteger, adParamInput, 4, txtfailurepk),_
				Array("@txtfailure",  adVarchar, adParamInput, 25, Trim(Mid(txtfailure,1,25))),_
				Array("@txtfailuredesc",  adVarchar, adParamInput, 50, Trim(Mid(txtfailuredesc,1,50))),_
				Array("@txtsolutionpk", adInteger, adParamInput, 4, txtsolutionpk),_
				Array("@txtsolution",  adVarchar, adParamInput, 25, Trim(Mid(txtsolution,1,25))),_
				Array("@txtsolutiondesc",  adVarchar, adParamInput, 50, Trim(Mid(txtsolutiondesc,1,50))),_
				Array("@txtfailurewo", adBoolean, adParamInput, 1, txtfailurewo),_
				Array("@txtmeter1reading", adInteger, adParamInput, 4, txtmeter1reading),_
				Array("@txtmeter2reading", adInteger, adParamInput, 4, txtmeter2reading),_
				Array("@txtisup", adBoolean, adParamInput, 1, txtisup),_
				Array("@txtDrawingUpdatesNeeded", adBoolean, adParamInput, 1, txtDrawingUpdatesNeeded),_
				Array("@mode",  adVarchar, adParamInput, 15, mode),_
				Array("@woauth", adChar, adParamInput, 1, GetSession("WOAuth")),_
				Array("@RowVersionUserPK", adInteger, adParamInput, 4, GetSession("UserPK")),_
				Array("@RowVersionInitials",  adVarchar, adParamInput, 5, GetSession("UserInitials")),_
				Array("@RowVersionIPAddress",  adVarchar, adParamInput, 25, GetSession("UserIPAddress"))_
				),"") Then

				HeaderMSG = db.derror
				If Not UCase(card) = "WOOPTIONS" Then
					CardSkipLevel = 1
				End If
				Card = WOAction
				WOClose

			Else
				If Not Request("RegularHours") = "" Then
					Call WOLaborRecSave(db,True,GetSession("UserID"))
					If Not HeaderMSG = "" Then
						If Not UCase(card) = "WOOPTIONS" Then
							CardSkipLevel = 1
						End If
						Card = WOAction
						WOClose
					End If
				End If
			End If

		End If

	End If

End Sub

'====================================================================================================================================

Sub MyWorkOrders()

	Dim db, sql, rs, parentlocationall, parentequipmentall, parentall

	CardTitle = "My WOs"
	CardCurrent = "MyWorkOrders"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetWOSQLWhere()

	Set db = New ADOHelper

	GetWOPK
	Call WOCloseSave(db)

	sql = _
	"SELECT DISTINCT " &_
	"WO.WOPK, " &_
	"WO.WOID, " &_
	"WO.RepairCenterID, " &_
	"WO.WOGroupPK, " &_
	"WO.TargetDate, " &_
	"WO.Reason, " &_
	"WO.ProcedureName, " &_
	"WO.AssetID, " &_
	"WO.AssetName, " &_
	"AssetHierarchy.ParentLocation, " &_
	"AssetHierarchy.ParentEquipment, " &_
	"AssetHierarchy.ParentLocationAll, " &_
	"AssetHierarchy.ParentEquipmentAll, " &_
	"Asset.IsLocation, " &_
	"WO.Type, " &_
	"WO.TypeDesc, " &_
	"WO.DemoLaborPK, " &_
	"WO.Status, " &_
	"WO.AuthStatus, " &_
	"WO.IsGenerated, " &_
	"WO.RouteOrder, " &_
	"WO.IsAssigned, " &_
	"WO.PrintedBox, " &_
	"WO.Priority, " &_
	"WO.AuthLevelsRequired, " &_
	"WO.IsApproved, " &_
	"WO.RouteOrder, " &_
	"WO.WOGroupType, " &_
	"WO.TargetHours, " &_
	"WO.PDAWO, " &_
	"WO.Complete, " &_
	"WO.Closed "

	sqlwhere = _
	"FROM WO WITH (NOLOCK) " &_
	"INNER JOIN WOassign WITH (NOLOCK) ON WO.WOPK = WOassign.WOPK " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.Complete Is Null " &_
	"AND WOAssign.LaborPK = '" & GetSession("UserPK") & "' " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) " &_
	sqlwhere &_
	"ORDER BY WO.TARGETDATE DESC,WO.RouteOrder,WO.WOPK DESC "

	sql = sql & sqlwhere

	Call SetSession("sqlwhere" & CardCurrentLevel,sqlwhere)

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.recordcount = 1 and UCase(CardFrom) = "WOSEARCH" Then
		Call SetSession("WOPK" & CardFromLevel,NullCheck(rs("WOPK")))
		CardFromLevel = CardFromLevel + 1
		Call WOOptions("")
	End If

	If UCase(CardFrom) = "WOSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
	If Not sqlewhere = "" Then %>
	<p>
	<b><% =sqlewhere %></b>
	</p>
	<% End If
	If rs.eof Then
		Call OutputWAPMsg("There were no Work Order Assignments found.")
	Else %>
		<p mode="nowrap">
		<%
		Do Until rs.Eof
		If Not ChoiceIndex > PagePos Then
			rs.MoveNext()
			ChoiceIndex = ChoiceIndex + 1
		Else
		    parentall = ""
            parentlocationall = WAPValidate(Replace(NullCheck(RS("parentlocationall")),"<br>"," - "))
            parentequipmentall = WAPValidate(Replace(NullCheck(RS("parentequipmentall")),"<br>"," - "))
            If Not parentlocationall = "" and Not parentequipmentall = "" Then
                parentall = parentlocationall & " - " & parentequipmentall & " - "
            ElseIf Not parentlocationall = "" Then
                parentall = parentlocationall & " - "
            ElseIf Not parentequipmentall = "" Then
                parentall = parentequipmentall & " - "
            End If
			%>
			<% If lang = "HTML" Then %><nobr><% End If %><% =LStyleBegin %><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =rs("WOPK") %>"><% =rs("WOPK") %>: <% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %> <% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) & " " %> <% If Not NullCheck(rs("AssetID")) = "" Then %><% =ParentAll %><% =WapValidate(NullCheck(rs("AssetName"))) %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write " (" & WAPValidate(NullCheck(rs("AssetID"))) & ")" End If %><% End If %></a><% =LStyleEnd %><% If lang = "HTML" Then %></nobr><% End If %><br/>
			<%
			rs.MoveNext()
			If Not rs.Eof Then
				TabIndex = TabIndex + 1
			End If
			If TabIndex = PageSize Then
				Exit Do
			End If
		End If
		Loop
		%>
		</p>
		<%
	End If
	%><p align="center">
	<%
	If PagePos >= PageSize Then
	%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
	<%
	End If
	If TabIndex = PageSize Then
	%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
	<%
	End If
	%></p>
	<%
	OutputButtons
	rs.Close()
	Set db = Nothing
	SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub AllWorkOrders()

	Dim db, sql, rs, parentlocationall, parentequipmentall, parentall

	CardTitle = "All WOs"
	CardCurrent = "AllWorkOrders"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetWOSQLWhere()

	Set db = New ADOHelper

	GetWOPK
	Call WOCloseSave(db)

	sql = _
	"SELECT DISTINCT " &_
	"WO.WOPK, " &_
	"WO.WOID, " &_
	"WO.RepairCenterID, " &_
	"WO.WOGroupPK, " &_
	"WO.TargetDate, " &_
	"WO.Reason, " &_
	"WO.ProcedureName, " &_
	"WO.AssetID, " &_
	"WO.AssetName, " &_
	"AssetHierarchy.ParentLocation, " &_
	"AssetHierarchy.ParentEquipment, " &_
	"AssetHierarchy.ParentLocationAll, " &_
	"AssetHierarchy.ParentEquipmentAll, " &_
	"Asset.IsLocation, " &_
	"WO.Type, " &_
	"WO.TypeDesc, " &_
	"WO.DemoLaborPK, " &_
	"WO.Status, " &_
	"WO.AuthStatus, " &_
	"WO.IsGenerated, " &_
	"WO.RouteOrder, " &_
	"WO.IsAssigned, " &_
	"WO.PrintedBox, " &_
	"WO.Priority, " &_
	"WO.AuthLevelsRequired, " &_
	"WO.IsApproved, " &_
	"WO.IsAssigned, " &_
	"WO.RouteOrder, " &_
	"WO.WOGroupType, " &_
	"WO.TargetHours, " &_
	"WO.PDAWO, " &_
	"WO.Complete, " &_
	"WO.Closed "

	sqlwhere = _
	"FROM WO WITH (NOLOCK) " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.Complete Is Null " &_
	"AND WO.RepairCenterPK = " & GetSession("RCPK") & " " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) " &_
	AddGeneralFilters("WO")	&_
	sqlwhere &_
	"ORDER BY WO.TARGETDATE DESC,WO.RouteOrder,WO.WOPK DESC "

	sql = sql & sqlwhere

	Call SetSession("sqlwhere" & CardCurrentLevel,sqlwhere)

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.recordcount = 1 and UCase(CardFrom) = "WOSEARCH" Then
		Call SetSession("WOPK" & CardFromLevel,NullCheck(rs("WOPK")))
		CardFromLevel = CardFromLevel + 1
		Call WOOptions("")
	End If

	If UCase(CardFrom) = "WOSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Work Orders found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
			    parentall = ""
                parentlocationall = WAPValidate(Replace(NullCheck(RS("parentlocationall")),"<br>"," - "))
                parentequipmentall = WAPValidate(Replace(NullCheck(RS("parentequipmentall")),"<br>"," - "))
                If Not parentlocationall = "" and Not parentequipmentall = "" Then
                    parentall = parentlocationall & " - " & parentequipmentall & " - "
                ElseIf Not parentlocationall = "" Then
                    parentall = parentlocationall & " - "
                ElseIf Not parentequipmentall = "" Then
                    parentall = parentequipmentall & " - "
                End If
				%>
				<% If lang = "HTML" Then %><nobr><% End If %><% =LStyleBegin %><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =rs("WOPK") %>"><% =rs("WOPK") %>: <% If BitNullCheck(rs("IsAssigned")) Then %>A <% Else %>U <% End If %> <% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %> <% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) & " " %> <% If Not NullCheck(rs("AssetID")) = "" Then %><% =ParentAll %><% =WapValidate(NullCheck(rs("AssetName"))) %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write " (" & WAPValidate(NullCheck(rs("AssetID"))) & ")" End If %><% End If %></a><% =LStyleEnd %><% If lang = "HTML" Then %></nobr><% End If %><br/>
				<% If False Then %>
				<select tabindex="2" name="S<% =rs("WOPK") %>" value="I">
				<option value="IS">Issued</option>
				<option value="OH">On-Hold</option>
				<option value="CO">Completed</option>
				<option value="CL">Closed</option>
				</select>
				Hours: <input tabindex="1" size="2" name="H<% =rs("WOPK") %>" value=""/>
				<br/>
				Report: <input tabindex="4" size="10" name="R<% =rs("WOPK") %>" value=""/>
				<br/>
				<% End If %>
				<%
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub


'====================================================================================================================================

Sub AllWorkOrdersU()

	Dim db, sql, rs, parentlocationall, parentequipmentall, parentall

	CardTitle = "All Unassigned WOs"
	CardCurrent = "AllWorkOrdersU"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetWOSQLWhere()

	Set db = New ADOHelper

	GetWOPK
	Call WOCloseSave(db)

	sql = _
	"SELECT DISTINCT " &_
	"WO.WOPK, " &_
	"WO.WOID, " &_
	"WO.RepairCenterID, " &_
	"WO.WOGroupPK, " &_
	"WO.TargetDate, " &_
	"WO.Reason, " &_
	"WO.ProcedureName, " &_
	"WO.AssetID, " &_
	"WO.AssetName, " &_
	"AssetHierarchy.ParentLocation, " &_
	"AssetHierarchy.ParentEquipment, " &_
	"AssetHierarchy.ParentLocationAll, " &_
	"AssetHierarchy.ParentEquipmentAll, " &_
	"Asset.IsLocation, " &_
	"WO.Type, " &_
	"WO.TypeDesc, " &_
	"WO.DemoLaborPK, " &_
	"WO.Status, " &_
	"WO.AuthStatus, " &_
	"WO.IsGenerated, " &_
	"WO.RouteOrder, " &_
	"WO.IsAssigned, " &_
	"WO.PrintedBox, " &_
	"WO.Priority, " &_
	"WO.AuthLevelsRequired, " &_
	"WO.IsApproved, " &_
	"WO.IsAssigned, " &_
	"WO.RouteOrder, " &_
	"WO.WOGroupType, " &_
	"WO.TargetHours, " &_
	"WO.PDAWO, " &_
	"WO.Complete, " &_
	"WO.Closed "

	sqlwhere = _
	"FROM WO WITH (NOLOCK) " &_
	"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
	"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
	"WHERE WO.IsOpen = 1 " &_
	"AND WO.IsAssigned = 0 " &_
	"AND WO.Complete Is Null " &_
	"AND WO.RepairCenterPK = " & GetSession("RCPK") & " " &_
	"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) " &_
	AddGeneralFilters("WO")	&_
	sqlwhere &_
	"ORDER BY WO.TARGETDATE DESC,WO.RouteOrder,WO.WOPK DESC "

	sql = sql & sqlwhere

	Call SetSession("sqlwhere" & CardCurrentLevel,sqlwhere)

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If rs.recordcount = 1 and UCase(CardFrom) = "WOSEARCH" Then
		Call SetSession("WOPK" & CardFromLevel,NullCheck(rs("WOPK")))
		CardFromLevel = CardFromLevel + 1
		Call WOOptions("")
	End If

	If UCase(CardFrom) = "WOSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Work Orders found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
			    parentall = ""
                parentlocationall = WAPValidate(Replace(NullCheck(RS("parentlocationall")),"<br>"," - "))
                parentequipmentall = WAPValidate(Replace(NullCheck(RS("parentequipmentall")),"<br>"," - "))
                If Not parentlocationall = "" and Not parentequipmentall = "" Then
                    parentall = parentlocationall & " - " & parentequipmentall & " - "
                ElseIf Not parentlocationall = "" Then
                    parentall = parentlocationall & " - "
                ElseIf Not parentequipmentall = "" Then
                    parentall = parentequipmentall & " - "
                End If
				%>
				<% If lang = "HTML" Then %><nobr><% End If %><% =LStyleBegin %><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =rs("WOPK") %>"><% =rs("WOPK") %>: <% If BitNullCheck(rs("IsAssigned")) Then %>A <% Else %>U <% End If %> <% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %> <% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) & " " %> <% If Not NullCheck(rs("AssetID")) = "" Then %><% =ParentAll %><% =WapValidate(NullCheck(rs("AssetName"))) %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write " (" & WAPValidate(NullCheck(rs("AssetID"))) & ")" End If %><% End If %></a><% =LStyleEnd %><% If lang = "HTML" Then %></nobr><% End If %><br/>
				<% If False Then %>
				<select tabindex="2" name="S<% =rs("WOPK") %>" value="I">
				<option value="IS">Issued</option>
				<option value="OH">On-Hold</option>
				<option value="CO">Completed</option>
				<option value="CL">Closed</option>
				</select>
				Hours: <input tabindex="1" size="2" name="H<% =rs("WOPK") %>" value=""/>
				<br/>
				Report: <input tabindex="4" size="10" name="R<% =rs("WOPK") %>" value=""/>
				<br/>
				<% End If %>
				<%
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CMLookup()

	Dim db, sql, rs

	CardTitle = "Company Lookup"
	CardCurrent = "CMLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetCMSQLWhere()

    Dim ft
    ft = GetSession("CMFT")

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Company WITH (NOLOCK) WHERE Active = 1 " &_
	sqlwhere &_
	"ORDER BY CompanyName, CompanyID "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "CMSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Companies found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;" & ft & "=" & WAPEncode(rs("CompanyID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform." & ft & ".value='" & JSEncode(rs("CompanyID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(NullCheck(rs("CompanyName"))) %> (<% =WAPValidate(NullCheck(rs("CompanyID"))) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(NullCheck(rs("CompanyName"))) %> (<% =WAPValidate(NullCheck(rs("CompanyID"))) %>)<go href="<% =returnurl %>"><setvar name="<% =ft %>" value="<% =WAPValidate(NullCheck(rs("CompanyID"))) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub FALookup()

	Dim db, sql, rs, desc, desc2

	CardCurrent = Card
	Select Case UCase(Card)
		Case "FAPRLOOKUP"
			CardTitle = "Problem Lookup"
			desc = "Problems"
			desc2 = "Problem"
		Case "FAFALOOKUP"
			CardTitle = "Failure Lookup"
			desc = "Failures"
			desc2 = "Failure"
		Case "FASOLOOKUP"
			CardTitle = "Solution Lookup"
			desc = "Solutions"
			desc2 = "Solution"
	End Select
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetFASQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Failure WITH (NOLOCK) WHERE Active = 1 " &_
	sqlwhere &_
	"ORDER BY FailureName, FailureID "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "FASEARCH" OR _
	   UCase(CardFrom) = "FAPRSEARCH" OR _
	   UCase(CardFrom) = "FAFASEARCH" OR _
	   UCase(CardFrom) = "FASOSEARCH" Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no " & desc & " found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;" & desc2 & "ID=" & WAPEncode(rs("FailureID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform." & desc2 & "ID.value='" & JSEncode(rs("FailureID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("FailureName")) %> (<% =WAPValidate(rs("FailureID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("FailureName")) %> (<% =WAPValidate(rs("FailureID")) %>)<go href="<% =returnurl %>"><setvar name="<% =desc2 %>ID" value="<% =WAPValidate(rs("FailureID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub Assets()

	Dim db, sql, rs, rsp, assetpk, regnode, parentpk

	LStyleBegin = ""
	LStyleEnd = ""

	CardTitle = "Assets"
	CardCurrent = "Assets"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize

    If lang = "HTML" Then
	    pagesize = 500
	Else
	    If IsPocketIE Then
		    pagesize = 50
	    Else
	        pagesize = 25
	    End If
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetASSQLWhere2()

	Set db = New ADOHelper

	sql = _
	"SELECT Asset.*, AssetHierarchy.AssetLevel, AssetHierarchy.HasChildren FROM Asset WITH (NOLOCK) " &_
	"INNER JOIN AssetHierarchy ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
	"WHERE 1=1 " &_
	sqlwhere &_
	"ORDER BY AssetName, AssetID "

	If UCase(CardFrom) = "ASSEARCH"	Then
		Pagepos = 0
	End If

    AssetPK = Request("AssetPK")
    ParentPK = ""

    If AssetPK = "" Then
        regnode = True
        AssetPK = "0"
    Else
        If Pagepos = 0 Then
            regnode = False
        Else
            regnode = True
        End If
    End If

    If Not sqlwhere = "" Then
	    Set rs = db.RunSQLReturnRS(sql,"")
    	Call CheckDB(db)
	    regnode = True
	Else
        Set rs = db.RunSPReturnRS("MC_GetAssetTree",Array(Array("@assetnode", adInteger, AdParamInput, 4, CLng(AssetPK)),Array("@excludecurrent", adBoolean, AdParamInput,1,0),Array("@autoopt", adBoolean, AdParamInput,1,1),Array("@autooptrecs", adInteger, AdParamInput,4,5),Array("@depth", adInteger, AdParamInput, 4,10000),Array("@system", adChar, AdParamInput,15,Null),Array("@repaircenterpk", adInteger, AdParamInput,4,""),Array("@locationonly", adBoolean, AdParamInput,1,0),Array("@demolaborpk", adInteger, AdParamInput,4,GetSession("UserPK")),Array("@laborpk", adInteger, AdParamInput,4,GetSession("UserPK"))),"")
    	Call CheckDB(db)
    	If Not rs.Eof Then
    	    If Not rs("AssetLevel") = "1" and Not AssetPK = "0" Then
                sql = "SELECT ParentPK FROM AssetHierarchy WITH (NOLOCK) WHERE AssetPK = " & AssetPK
	            Set rsp = db.RunSQLReturnRS(sql,"")
    	        Call CheckDB(db)
	            If Not rsp.Eof Then
    	            ParentPK = rsp("ParentPK")
	            End If
	            Call CloseObj(rsp)
	        End If
	    End If
    End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Assets found.")
		Else %>
			<p mode="nowrap">
			<%

            Dim NoIconFile_Location, NoIconFile_Asset, IconFile, ParentAll, parentlocationall, parentequipmentall

			NoIconFile_Location = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/facility_g.gif"
			NoIconFile_Asset = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/gearsxp_g.gif"

            If Not ParentPK = "" Then %>
            <% =LStyleBegin %><a href="default.asp?card=assets&amp;s=<% =SessionID %>&amp;assetpk=<% =ParentPK %>&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><img <% If lang="HTML" Then %>border="0" <% End If %>src="images/gobackup.gif"/><% If Not IsBlackBerry Then %> Up<% End If %></a><% If Not IsBlackBerry Then %><br/><% End If %>
            <% If False Then %><% If lang="HTML" Then %>&nbsp;<% End If %> <a href="default.asp?card=assets&amp;s=<% =SessionID %>&amp;assetpk=&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>">Top</a><% End If %><% =LStyleEnd %><%
            End If

			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
			    If NullCheck(rs("AssetLevel")) > 1 and Not RegNode Then
                    parentlocationall = WAPValidate(Replace(NullCheck(RS("parentlocationall")),"<br>"," - "))
                    parentequipmentall = WAPValidate(Replace(NullCheck(RS("parentequipmentall")),"<br>"," - "))
        		    parentall = ""
                    If Not parentlocationall = "" and Not parentequipmentall = "" Then
                        parentall = parentlocationall & " - " & parentequipmentall & " - "
                    ElseIf Not parentlocationall = "" Then
                        parentall = parentlocationall & " - "
                    ElseIf Not parentequipmentall = "" Then
                        parentall = parentequipmentall & " - "
                    End If
                End If
			    If NullCheck(rs("icon")) = "" Then
				      If BitNullCheck(rs("islocation")) Then
					    IconFile = NoIconFile_Location
				      Else
					    IconFile = NoIconFile_Asset
				      End If
			    Else
				      IconFile = Application("MCVirtualDirectory") & Application("mapp_path") & rs("icon")
			    End If
				returnurl = "default.asp?card=asoptions&amp;s=" & SessionID & "&amp;AssetPK=" & WAPEncode(rs("AssetPK"))
				If lang="HTML" Then
					'returnurl = "" & JSEncode(rs("AssetID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr>
    			    <% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %>
					<% If BitNullCheck(rs("HasChildren")) Then %>
					<a href="default.asp?card=assets&amp;s=<% =SessionID %>&amp;assetpk=<% =rs("AssetPK") %>&amp;pagepos=<% =pagepos %>&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><img border="0" src="images/closed.gif"/></a><% Else %>
					<img border="0" src="images/nochildren.gif"/><%
					End If
					End If %>
					<% If Not Application("MCVirtualDirectory") = "" Then %>
					<img border="0" src="<% =IconFile %>" alt="Asset Icon"/>
					<% End If %><% =LStyleBegin %><a href="<% =returnurl %>"><% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %><% Else %><% =ParentAll %><% End If %><% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "(" & WAPValidate(rs("AssetID")) & ")" End If %></a><% =LStyleEnd %>
					</nobr><%
				Else %>
				    <% Dim OldIsPPC %>
				    <% OldIsPPC = IsPPC %>
				    <% IsPPC = False %>
				    <% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %>
					<% If BitNullCheck(rs("HasChildren")) Then %>
					<a href="default.asp?card=assets&amp;s=<% =SessionID %>&amp;assetpk=<% =rs("AssetPK") %>&amp;pagepos=<% =pagepos %>&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><% If IsPPC Then %><img src="images/closed.gif"/><% Else %>_+_<% End If %></a><% If Not IsPPC Then %>&nbsp; <% End If %><% Else %>
					<% If IsPPC Then %><img src="images/nochildren.gif"/><% Else %>___ <% End If %><%
					End If
					End If %>
                    <% If IsPPC and Not Application("MCVirtualDirectory") = "" Then %><img src="<% =IconFile %>"/> <% End If %><% =LStyleBegin %><anchor><% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %><% Else %><% =ParentAll %><% End If %><% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "(" & WAPValidate(rs("AssetID")) & ")" End If %><go href="<% =returnurl %>"><setvar name="AssetID" value="<% =WAPValidate(rs("AssetID")) %>"/></go></anchor><% =LStyleEnd %><%
                    IsPPC = OldIsPPC
				End If
				RegNode = True
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
				    Response.Write "<br/>"
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASLookup()

	Dim db, sql, rs, rsp, assetpk, regnode, parentpk

	CardTitle = "Asset Lookup"
	CardCurrent = "ASLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize

    If lang = "HTML" Then
	    pagesize = 500
	Else
	    If IsPocketIE Then
		    pagesize = 50
	    Else
	        pagesize = 25
	    End If
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetASSQLWhere()

    Dim ft
    ft = GetSession("ASFT")

	Set db = New ADOHelper

	sql = _
	"SELECT Asset.*, AssetHierarchy.AssetLevel, AssetHierarchy.HasChildren FROM Asset WITH (NOLOCK) " &_
	"INNER JOIN AssetHierarchy ON AssetHierarchy.AssetPK = Asset.AssetPK " &_
	"WHERE 1=1 " &_
	sqlwhere &_
	"ORDER BY AssetName, AssetID "

	If UCase(CardFrom) = "ASSEARCH"	Then
		Pagepos = 0
	End If

    AssetPK = Request("AssetPK")
    ParentPK = ""

    If AssetPK = "" Then
        regnode = True
        AssetPK = "0"
    Else
        If Pagepos = 0 Then
            regnode = False
        Else
            regnode = True
        End If
    End If

    If Not sqlwhere = "" Then
	    Set rs = db.RunSQLReturnRS(sql,"")
    	Call CheckDB(db)
	    regnode = True
	Else
        Set rs = db.RunSPReturnRS("MC_GetAssetTree",Array(Array("@assetnode", adInteger, AdParamInput, 4, CLng(AssetPK)),Array("@excludecurrent", adBoolean, AdParamInput,1,0),Array("@autoopt", adBoolean, AdParamInput,1,1),Array("@autooptrecs", adInteger, AdParamInput,4,5),Array("@depth", adInteger, AdParamInput, 4,10000),Array("@system", adChar, AdParamInput,15,Null),Array("@repaircenterpk", adInteger, AdParamInput,4,""),Array("@locationonly", adBoolean, AdParamInput,1,0),Array("@demolaborpk", adInteger, AdParamInput,4,GetSession("UserPK")),Array("@laborpk", adInteger, AdParamInput,4,GetSession("UserPK"))),"")
    	Call CheckDB(db)
    	If Not rs.Eof Then
    	    If Not rs("AssetLevel") = "1" and Not AssetPK = "0" Then
                sql = "SELECT ParentPK FROM AssetHierarchy WITH (NOLOCK) WHERE AssetPK = " & AssetPK
	            Set rsp = db.RunSQLReturnRS(sql,"")
    	        Call CheckDB(db)
	            If Not rsp.Eof Then
    	            ParentPK = rsp("ParentPK")
	            End If
	            Call CloseObj(rsp)
	        End If
	    End If
    End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Assets found.")
		Else %>
			<p mode="nowrap">
			<%

            Dim NoIconFile_Location, NoIconFile_Asset, IconFile, ParentAll, parentlocationall, parentequipmentall

			NoIconFile_Location = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/facility_g.gif"
			NoIconFile_Asset = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/gearsxp_g.gif"

            If Not ParentPK = "" Then %>
            <% =LStyleBegin %><a href="default.asp?card=aslookup&amp;s=<% =SessionID %>&amp;assetpk=<% =ParentPK %>&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><img <% If lang="HTML" Then %>border="0" <% End If %>src="images/gobackup.gif"/><% If Not IsBlackBerry Then %> Up<% End If %></a><% If Not IsBlackBerry Then %><br/><% End If %>
            <% If False Then %><% If lang="HTML" Then %>&nbsp;<% End If %> <a href="default.asp?card=aslookup&amp;s=<% =SessionID %>&amp;assetpk=&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>">Top</a><% End If %><% =LStyleEnd %><%
            End If

			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
			    If NullCheck(rs("AssetLevel")) > 1 and Not RegNode Then
                    parentlocationall = WAPValidate(Replace(NullCheck(RS("parentlocationall")),"<br>"," - "))
                    parentequipmentall = WAPValidate(Replace(NullCheck(RS("parentequipmentall")),"<br>"," - "))
        		    parentall = ""
                    If Not parentlocationall = "" and Not parentequipmentall = "" Then
                        parentall = parentlocationall & " - " & parentequipmentall & " - "
                    ElseIf Not parentlocationall = "" Then
                        parentall = parentlocationall & " - "
                    ElseIf Not parentequipmentall = "" Then
                        parentall = parentequipmentall & " - "
                    End If
                End If
			    If NullCheck(rs("icon")) = "" Then
				      If BitNullCheck(rs("islocation")) Then
					    IconFile = NoIconFile_Location
				      Else
					    IconFile = NoIconFile_Asset
				      End If
			    Else
				      IconFile = Application("MCVirtualDirectory") & Application("mapp_path") & rs("icon")
			    End If
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;" & ft & "=" & WAPEncode(rs("AssetID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform." & ft & ".value='" & JSEncode(rs("AssetID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr>
    			    <% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %>
					<% If BitNullCheck(rs("HasChildren")) Then %>
					<a href="default.asp?card=aslookup&amp;s=<% =SessionID %>&amp;assetpk=<% =rs("AssetPK") %>&amp;pagepos=<% =pagepos %>&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><img border="0" src="images/closed.gif"/></a><% Else %>
					<img border="0" src="images/nochildren.gif"/><%
					End If
					End If %>
					<% If Not Application("MCVirtualDirectory") = "" Then %>
					<img border="0" src="<% =IconFile %>" alt="Asset Icon"/>
					<% End If %><% =LStyleBegin %><a href="<% =returnurl %>"><% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %><% Else %><% =ParentAll %><% End If %><% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "(" & WAPValidate(rs("AssetID")) & ")" End If %></a><% =LStyleEnd %>
					</nobr><%
				Else %>
				    <% Dim OldIsPPC %>
				    <% OldIsPPC = IsPPC %>
				    <% IsPPC = False %>
				    <% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %>
					<% If BitNullCheck(rs("HasChildren")) Then %>
					<a href="default.asp?card=aslookup&amp;s=<% =SessionID %>&amp;assetpk=<% =rs("AssetPK") %>&amp;pagepos=<% =pagepos %>&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><% If IsPPC Then %><img src="images/closed.gif"/><% Else %>+<% End If %></a><% If Not IsPPC Then %>&nbsp; <% End If %><% Else %>
					<% If IsPPC Then %><img src="images/nochildren.gif"/><% Else %>_ <% End If %><%
					End If
					End If %>
                    <% If IsPPC and Not Application("MCVirtualDirectory") = "" Then %><img src="<% =IconFile %>"/> <% End If %><% =LStyleBegin %><anchor><% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %><% Else %><% =ParentAll %><% End If %><% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "(" & WAPValidate(rs("AssetID")) & ")" End If %><go href="<% =returnurl %>"><setvar name="<% =ft %>" value="<% =WAPValidate(rs("AssetID")) %>"/></go></anchor><% =LStyleEnd %><%
                    IsPPC = OldIsPPC
				End If
				RegNode = True
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
				    Response.Write "<br/>"
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CLLookup()

	Dim db, sql, rs, rsp, classificationpk, regnode, parentpk

	CardTitle = "Classification Lookup"
	CardCurrent = "CLLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize

    If lang = "HTML" Then
	    pagesize = 500
	Else
	    If IsPocketIE Then
		    pagesize = 50
	    Else
	        pagesize = 25
	    End If
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetCLSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT Classification.*, ClassificationHierarchy.ClassificationLevel, ClassificationHierarchy.HasChildren FROM Classification WITH (NOLOCK) " &_
	"INNER JOIN ClassificationHierarchy ON ClassificationHierarchy.ClassificationPK = Classification.ClassificationPK " &_
	"WHERE System = 'AS' " &_
	sqlwhere &_
	"ORDER BY ClassificationName, ClassificationID "

	If UCase(CardFrom) = "CLSEARCH"	Then
		Pagepos = 0
	End If

    ClassificationPK = Request("ClassificationPK")
    ParentPK = ""

    If ClassificationPK = "" Then
        regnode = True
        ClassificationPK = "0"
    Else
        If Pagepos = 0 Then
            regnode = False
        Else
            regnode = True
        End If
    End If

    If Not sqlwhere = "" Then
	    Set rs = db.RunSQLReturnRS(sql,"")
    	Call CheckDB(db)
	    regnode = True
	Else
        Set rs = db.RunSPReturnRS("MC_GetClassificationTree",Array(Array("@Classificationnode", adInteger, AdParamInput, 4, CLng(ClassificationPK)),Array("@excludecurrent", adBoolean, AdParamInput,1,0),Array("@autoopt", adBoolean, AdParamInput,1,1),Array("@autooptrecs", adInteger, AdParamInput,4,5),Array("@depth", adInteger, AdParamInput, 4,1),Array("@system", adChar, AdParamInput,15,"AS"),Array("@repaircenterpk", adInteger, AdParamInput,4,""),Array("@locationonly", adBoolean, AdParamInput,1,0),Array("@demolaborpk", adInteger, AdParamInput,4,GetSession("UserPK"))),"")
    	Call CheckDB(db)
    	If Not rs.Eof Then
    	    If Not rs("ClassificationLevel") = "1" and Not ClassificationPK = "0" Then
                sql = "SELECT ParentPK FROM ClassificationHierarchy WITH (NOLOCK) WHERE ClassificationPK = " & ClassificationPK
	            Set rsp = db.RunSQLReturnRS(sql,"")
    	        Call CheckDB(db)
	            If Not rsp.Eof Then
    	            ParentPK = rsp("ParentPK")
	            End If
	            Call CloseObj(rsp)
	        End If
	    End If
    End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Classifications found.")
		Else %>
			<p mode="nowrap">
			<%

            Dim NoIconFile_Location, NoIconFile_Asset, IconFile, ParentAll, parentlocationall, parentequipmentall

			NoIconFile_Location = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/facility_g.gif"
			NoIconFile_Asset = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/gearsxp_g.gif"

            If Not ParentPK = "" Then %>
            <% =LStyleBegin %><a href="default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;classificationpk=<% =ParentPK %>&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><img <% If lang="HTML" Then %>border="0" <% End If %>src="images/gobackup.gif"/><% If Not IsBlackBerry Then %> Up<% End If %></a><% If Not IsBlackBerry Then %><br/><% End If %>
            <% If False Then %><% If lang="HTML" Then %>&nbsp;<% End If %> <a href="default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;classificationpk=&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>">Top</a><% End If %><% =LStyleEnd %><%
            End If

			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
			    If NullCheck(rs("ClassificationLevel")) > 1 and Not RegNode Then
                    parentlocationall = WAPValidate(Replace(NullCheck(RS("parentlocationall")),"<br>"," - "))
                    'parentequipmentall = WAPValidate(Replace(NullCheck(RS("parentequipmentall")),"<br>"," - "))
        		    parentall = ""
                    If Not parentlocationall = "" Then
                        parentall = parentlocationall & " - "
                    End If
                End If
			    If NullCheck(rs("icon")) = "" Then
				      'If BitNullCheck(rs("islocation")) Then
					  '  IconFile = NoIconFile_Location
				      'Else
					    IconFile = NoIconFile_Asset
				      'End If
			    Else
				      IconFile = Application("MCVirtualDirectory") & Application("mapp_path") & rs("icon")
			    End If
			    Dim IsLocation
			    IsLocation = True
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;ClassificationID=" & WAPEncode(rs("ClassificationID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.ClassificationID.value='" & JSEncode(rs("ClassificationID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr>
    			    <% If NullCheck(rs("ClassificationLevel")) > 1 and RegNode Then %>
					<% If BitNullCheck(rs("HasChildren")) Then %>
					<a href="default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;classificationpk=<% =rs("ClassificationPK") %>&amp;pagepos=<% =pagepos %>&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><img border="0" src="images/closed.gif"/></a><% Else %>
					<img border="0" src="images/nochildren.gif"/><%
					End If
					End If %>
					<% If Not Application("MCVirtualDirectory") = "" Then %>
					<img border="0" src="<% =IconFile %>" alt="Classification Icon"/>
					<% End If %><% =LStyleBegin %><a href="<% =returnurl %>"><% If NullCheck(rs("ClassificationLevel")) > 1 and RegNode Then %><% Else %><% =ParentAll %><% End If %><% =WAPValidate(rs("ClassificationName")) & " " %><% If Not IsLocation Then Response.Write "(" & WAPValidate(rs("ClassificationID")) & ")" End If %></a><% =LStyleEnd %>
					</nobr><%
				Else %>
				    <% Dim OldIsPPC %>
				    <% OldIsPPC = IsPPC %>
				    <% IsPPC = False %>
				    <% If NullCheck(rs("ClassificationLevel")) > 1 and RegNode Then %>
					<% If BitNullCheck(rs("HasChildren")) Then %>
					<a href="default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;classificationpk=<% =rs("ClassificationPK") %>&amp;pagepos=<% =pagepos %>&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>"><% If IsPPC Then %><img src="images/closed.gif"/><% Else %>+<% End If %></a><% If Not IsPPC Then %>&nbsp; <% End If %><% Else %>
					<% If IsPPC Then %><img src="images/nochildren.gif"/><% Else %>_ <% End If %><%
					End If
					End If %>
                    <% If IsPPC and Not Application("MCVirtualDirectory") = "" Then %><img src="<% =IconFile %>"/> <% End If %><% =LStyleBegin %><anchor><% If NullCheck(rs("ClassificationLevel")) > 1 and RegNode Then %><% Else %><% =ParentAll %><% End If %><% =WAPValidate(rs("ClassificationName")) & " " %><% If Not IsLocation Then Response.Write "(" & WAPValidate(rs("ClassificationID")) & ")" End If %><go href="<% =returnurl %>"><setvar name="ClassificationID" value="<% =WAPValidate(rs("ClassificationID")) %>"/></go></anchor><% =LStyleEnd %><%
                    IsPPC = OldIsPPC
				End If
				RegNode = True
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
				If Not rs.Eof Then
				    Response.Write "<br/>"
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;Classificationpk=<% =ClassificationPK %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;Classificationpk=<% =ClassificationPK %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASForSerialPartLookup()

	Dim db, sql, rs

	CardTitle = "Asset Lookup"
	CardCurrent = "ASForSerialPartLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetASSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Asset WITH (NOLOCK) " &_
	"WHERE 1=1 " &_
	sqlwhere &_
	"ORDER BY AssetName, AssetID "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "ASSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Assets found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;AssetID=" & WAPEncode(rs("AssetID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.SerialReplaceToLocationID.value='" & JSEncode(rs("AssetID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "(" & WAPValidate(rs("AssetID")) & ")" End If %></a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "(" & WAPValidate(rs("AssetID")) & ")" End If %><go href="<% =returnurl %>"><setvar name="SerialReplaceToLocationID" value="<% =WAPValidate(rs("AssetID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub LTLookup()

	Dim db, sql, rs

	CardTitle = "Lookup"
	CardCurrent = "LTLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug
	sqlwhere = " LookupTable = '" & Request("lt") & "' "

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM LookupTableValues WITH (NOLOCK) WHERE " &_
	sqlwhere &_
	"ORDER BY CodeValue, CodeDesc "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no values found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;" & Request("lf") & "=" & WAPEncode(rs(Trim(Request("lr"))))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform." & Request("lf") & ".value='" & JSEncode(rs(Trim(Request("lr")))) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% If rs("CodeDesc") = "" Then %>(Not Specified)<% Else %><% =WAPValidate(rs("CodeDesc")) & " " %>(<% =WAPValidate(rs("CodeName")) & ")" %><% End If %></a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% If rs("CodeDesc") = "" Then %>(Not Specified)<% Else %><% =WAPValidate(rs("CodeDesc")) & " " %>(<% =WAPValidate(rs("CodeName")) %>)<% End If %><go href="<% =returnurl %>"><setvar name="<% =Request("lf") %>" value="<% =WAPValidate(rs(Trim(Request("lr")))) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ACLookup()

	Dim db, sql, rs

	CardTitle = "Account Lookup"
	CardCurrent = "ACLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetACSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Account WITH (NOLOCK) WHERE Active = 1 " &_
	sqlwhere &_
	"ORDER BY AccountID, AccountName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "ACSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Accounts found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;AccountID=" & WAPEncode(rs("AccountID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.AccountID.value='" & JSEncode(rs("AccountID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("AccountID")) %> <% =WAPValidate(rs("AccountName")) %></a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("AccountID")) %> <% =WAPValidate(rs("AccountName")) %><go href="<% =returnurl %>"><setvar name="AccountID" value="<% =WAPValidate(rs("AccountID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub RCLookup()

	Dim db, sql, rs, rs2, anchoroutput

	CardTitle = "Repair Center Lookup"
	CardCurrent = "RCLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetRCSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM RepairCenter WITH (NOLOCK) WHERE Active = 1 " &_
	sqlwhere &_
	"ORDER BY RepairCenterName, RepairCenterID "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "RCSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Repair Centers found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;RepairCenterID=" & WAPEncode(rs("RepairCenterID"))
				anchoroutput = WAPValidate(rs("RepairCenterName")) & " - " & WAPValidate(rs("RepairCenterID"))
				If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
					sql = _
					"SELECT WOCount = COUNT(WO.WOPK) " &_
					"FROM WO WITH (NOLOCK) " &_
					"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
					"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
					"WHERE WO.IsOpen = 1 " &_
					"AND WO.Complete Is Null " &_
					"AND WO.RepairCenterPK = " & rs("RepairCenterPK") & " " &_
					"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs.eof Then
						anchoroutput = anchoroutput & " (0)"
					Else
						anchoroutput = anchoroutput & " (" & rs2("WOCount") & ")"
					End If
				End If
				If lang="HTML" Then
					If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
						returnurl = "javascript:parent.location.replace('" & JSEncode(returnurl) & "');closeiframe('" & JSEncode(returnurl) & "');"
					Else
						returnurl = "javascript:parent.document.mcform.RepairCenterID.value='" & JSEncode(rs("RepairCenterID")) & "';closeiframe('" & JSEncode(returnurl) & "');"
					End If %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =anchoroutput %></a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =anchoroutput %><go href="<% =returnurl %>"><setvar name="RepairCenterID" value="<% =WAPValidate(rs("RepairCenterID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub SHLookup()

	Dim db, sql, rs, rs2, anchoroutput, addall

	CardTitle = "Shop Lookup"
	CardCurrent = "SHLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
		AddAll = True
	Else
		AddAll = False
	End If

	'Call ASPDebug

	Call GetSHSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Shop WITH (NOLOCK) WHERE Active = 1 " &_
	"AND Shop.RepairCenterPK = " & GetSession("RCPK") & " " &_
	sqlwhere &_
	"ORDER BY ShopName, ShopID "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "SHSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof and Not AddAll Then
			Call OutputWAPMsg("There were no Shops found.")
		Else %>
			<p mode="nowrap">
			<%
			If AddAll and PagePos = 0 Then
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;ShopID=ALL"
				anchoroutput = "All Shops"
				If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
					sql = _
					"SELECT WOCount = COUNT(WO.WOPK) " &_
					"FROM WO WITH (NOLOCK) " &_
					"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
					"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
					"WHERE WO.IsOpen = 1 " &_
					"AND WO.Complete Is Null " &_
					"AND WO.RepairCenterPK = " & GetSession("RCPK") & " " &_
					"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					' This is actually ambiguous because if shop is not specified on a WO
					' then it would look like the All Shops count is higher than it really is
					' and then the sum of shop totals would not equal the "All Shops" count.
					If rs.eof Then
						'anchoroutput = anchoroutput & " (0)"
					Else
						'anchoroutput = anchoroutput & " (" & rs2("WOCount") & ")"
					End If
				End If
				If lang="HTML" Then
					If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
						returnurl = "javascript:parent.location.replace('" & JSEncode(returnurl) & "');closeiframe('" & JSEncode(returnurl) & "');"
					Else
						returnurl = "javascript:parent.document.mcform.ShopID.value='ALL';closeiframe('" & JSEncode(returnurl) & "');"
					End If %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =anchoroutput %></a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =anchoroutput %><go href="<% =returnurl %>"><setvar name="ShopID" value="ALL"/></go></anchor><% =LStyleEnd %><br/><%
				End If
			End If
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;ShopID=" & WAPEncode(rs("ShopID"))
				anchoroutput = WAPValidate(rs("ShopName")) & " - " & WAPValidate(rs("ShopID"))
				If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
					sql = _
					"SELECT WOCount = COUNT(WO.WOPK) " &_
					"FROM WO WITH (NOLOCK) " &_
					"LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK " &_
					"LEFT OUTER JOIN Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " &_
					"WHERE WO.IsOpen = 1 " &_
					"AND WO.Complete Is Null " &_
					"AND WO.RepairCenterPK = " & GetSession("RCPK") & " " &_
					"AND WO.ShopPK = " & RS("ShopPK") & " " &_
					"AND (WO.WOGroupType <> 'C' or WO.WOGroupType Is Null) "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs.eof Then
						anchoroutput = anchoroutput & " (0)"
					Else
						anchoroutput = anchoroutput & " (" & rs2("WOCount") & ")"
					End If
				End If
				If lang="HTML" Then
					If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
						returnurl = "javascript:parent.location.replace('" & JSEncode(returnurl) & "');closeiframe('" & JSEncode(returnurl) & "');"
					Else
						returnurl = "javascript:parent.document.mcform.ShopID.value='" & JSEncode(rs("ShopID")) & "';closeiframe('" & JSEncode(returnurl) & "');"
					End If %>
					<% =LStyleBegin %><a href="<% =returnurl %>"><% =anchoroutput %></a><% =LStyleEnd %><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =anchoroutput %><go href="<% =returnurl %>"><setvar name="ShopID" value="<% =WAPValidate(rs("ShopID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CALookup()

	Dim db, sql, rs

	CardTitle = "Category Lookup"
	CardCurrent = "CALookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetCASQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Category WITH (NOLOCK) WHERE Active = 1 " &_
	" AND (mcmodule = 'WO' or mcmodule Is Null) " &_
	sqlwhere &_
	"ORDER BY CategoryName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "CASEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Categories found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;CategoryID=" & WAPEncode(rs("CategoryID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.CategoryID.value='" & JSEncode(rs("CategoryID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("CategoryName")) %> (<% =WAPValidate(rs("CategoryID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("CategoryName")) %> (<% =WAPValidate(rs("CategoryID")) %>)<go href="<% =returnurl %>"><setvar name="CategoryID" value="<% =WAPValidate(rs("CategoryID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ZNLookup()

	Dim db, sql, rs

	CardTitle = "Zone Lookup"
	CardCurrent = "ZNLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetZNSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Zone WITH (NOLOCK) WHERE Active = 1 " &_
	sqlwhere &_
	"ORDER BY ZoneName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "ZNSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Zones found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;ZoneID=" & WAPEncode(rs("ZoneID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.ZoneID.value='" & JSEncode(rs("ZoneID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("ZoneName")) %> (<% =WAPValidate(rs("ZoneID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("ZoneName")) %> (<% =WAPValidate(rs("ZoneID")) %>)<go href="<% =returnurl %>"><setvar name="ZoneID" value="<% =WAPValidate(rs("ZoneID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub PRLookup()

	Dim db, sql, rs

	CardTitle = "Procedure Lookup"
	CardCurrent = "PRLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetPRSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM ProcedureLibrary WITH (NOLOCK) WHERE Active = 1 " &_
	" AND (RepairCenterPK Is Null or RepairCenterPK = " & GetSession("RCPK") & ") " &_
	sqlwhere &_
	"ORDER BY ProcedureName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "PRSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Procedures found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;ProcedureID=" & WAPEncode(rs("ProcedureID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.ProcedureID.value='" & JSEncode(rs("ProcedureID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("ProcedureName")) %> (<% =WAPValidate(rs("ProcedureID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("ProcedureName")) %> (<% =WAPValidate(rs("ProcedureID")) %>)<go href="<% =returnurl %>"><setvar name="ProcedureID" value="<% =WAPValidate(rs("ProcedureID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub DPLookup()

	Dim db, sql, rs

	CardTitle = "Department Lookup"
	CardCurrent = "DPLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetDPSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Department WITH (NOLOCK) WHERE Active = 1 " &_
	sqlwhere &_
	"ORDER BY DepartmentName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "DPSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Departments found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;DepartmentID=" & WAPEncode(rs("DepartmentID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.DepartmentID.value='" & JSEncode(rs("DepartmentID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("DepartmentName")) %> (<% =WAPValidate(rs("DepartmentID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("DepartmentName")) %> (<% =WAPValidate(rs("DepartmentID")) %>)<go href="<% =returnurl %>"><setvar name="DepartmentID" value="<% =WAPValidate(rs("DepartmentID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub TNLookup()

	Dim db, sql, rs

	CardTitle = "Customer Lookup"
	CardCurrent = "TNLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetTNSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Tenant WITH (NOLOCK) WHERE Active = 1 " &_
	sqlwhere &_
	"ORDER BY TenantName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "TNSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Customers found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;TenantID=" & WAPEncode(rs("TenantID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.TenantID.value='" & JSEncode(rs("TenantID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("TenantName")) %> (<% =WAPValidate(rs("TenantID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("TenantName")) %> (<% =WAPValidate(rs("TenantID")) %>)<go href="<% =returnurl %>"><setvar name="TenantID" value="<% =WAPValidate(rs("TenantID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub PJLookup()

	Dim db, sql, rs

	CardTitle = "Project Lookup"
	CardCurrent = "PJLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetPJSQLWhere()

	Set db = New ADOHelper

	sql = _
	"SELECT * FROM Project WITH (NOLOCK) WHERE IsOpen = 1 " &_
	" AND (RepairCenterPK Is Null or RepairCenterPK = " & GetSession("RCPK") & ") " &_
	sqlwhere &_
	"ORDER BY ProjectName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "PJSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Projects found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;ProjectID=" & WAPEncode(rs("ProjectID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.ProjectID.value='" & JSEncode(rs("ProjectID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("ProjectName")) %> (<% =WAPValidate(rs("ProjectID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("ProjectName")) %> (<% =WAPValidate(rs("ProjectID")) %>)<go href="<% =returnurl %>"><setvar name="ProjectID" value="<% =WAPValidate(rs("ProjectID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub LALookup()

	Dim db, sql, rs

	CardTitle = "Labor Lookup"
	CardCurrent = "LALookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetLASQLWhere()

    Dim ft
    ft = GetSession("LAFT")

	Set db = New ADOHelper

    Select Case UCase(ft)
        Case "LABORID"
	        sql = _
	        "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborType IN ('EMP','CON') AND Active = 1 " &_
	        "AND Labor.RepairCenterPK = " & GetSession("RCPK") & " " &_
	        sqlwhere &_
	        "ORDER BY LaborName "
        Case "OPERATORID"
	        sql = _
	        "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborType IN ('EMP','CON') AND Active = 1 " &_
	        "AND Labor.RepairCenterPK = " & GetSession("RCPK") & " " &_
	        sqlwhere &_
	        "ORDER BY LaborName "
        Case "CONTACTID"
	        sql = _
	        "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborType IN ('EMP','CON','CONTACT','REQ') AND Active = 1 " &_
	        "AND Labor.RepairCenterPK = " & GetSession("RCPK") & " " &_
	        sqlwhere &_
	        "ORDER BY LaborName "
    End Select

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "LASEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There was no Labor found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;" & ft & "=" & WAPEncode(rs("LaborID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform." & ft & ".value='" & JSEncode(rs("LaborID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("LaborName")) %> (<% =WAPValidate(rs("LaborID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("LaborName")) %> (<% =WAPValidate(rs("LaborID")) %>)<go href="<% =returnurl %>"><setvar name="<% =ft %>" value="<% =WAPValidate(rs("LaborID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub INLookup()

	Dim db, sql, rs

	CardTitle = "Item Lookup"
	CardCurrent = "INLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetINSQLWhere()

	Set db = New ADOHelper

    Dim ft
    ft = GetSession("INFT")

	sql = _
	"SELECT * FROM Part WITH (NOLOCK) WHERE Active = 1 " &_
	sqlwhere &_
	"ORDER BY PartID, PartName "

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "INSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Items found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else

				Dim vpnstring
				vpnstring = ""
				If True Then
					Dim sql2,rs2
					sql2 = "Select Distinct VendorPartNumber FROM PartVendor WHERE PartPK = " & rs("PartPK")
					Set rs2 = db.RunSQLReturnRS(sql2,"")
					Call CheckDB(db)
					Do While Not rs2.Eof
						vpnstring = vpnstring & NullCheck(rs2("VendorPartNumber"))
						rs2.MoveNext
						If Not rs2.Eof Then
							vpnstring = vpnstring & ", "
						End If
					Loop
				End If
			    If Not vpnstring = "" Then
					vpnstring = "[" & vpnstring & "]"
				End If
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;" & ft & "=" & WAPEncode(rs("PartID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform." & ft & ".value='" & JSEncode(rs("PartID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>">[<% =rs("PartID") %>] [<% =WAPValidate(rs("PartName")) %>] <% =vpnstring %></a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor>[<% =WAPValidate(rs("PartID")) %>] [<% =WAPValidate(rs("PartName")) %>] <% =vpnstring %><go href="<% =returnurl %>"><setvar name="<% =ft %> %>" value="<% =WAPValidate(rs("PartID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub SRLookup()

	Dim db, sql, rs

	CardTitle = "Location Lookup"
	CardCurrent = "SRLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetSRSQLWhere()

	Set db = New ADOHelper

	'sql = _
	'"SELECT * FROM Location WITH (NOLOCK) WHERE Type = 'SR' AND Active = 1 " &_
	'"AND Location.RepairCenterPK = " & GetSession("RCPK") & " " &_
	'sqlwhere &_
	'"ORDER BY LocationName "

	sql = _
	"SELECT * FROM Location WITH (NOLOCK) WHERE Type = 'SR' AND Active = 1 " &_
	sqlwhere &_
	"ORDER BY LocationName "

	'Call OutputWAPError(sql)

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	If UCase(CardFrom) = "SRSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Locations found.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;LocationID=" & WAPEncode(rs("LocationID"))
				If lang="HTML" Then
					returnurl = "javascript:parent.document.mcform.LocationID.value='" & JSEncode(rs("LocationID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% =WAPValidate(rs("LocationName")) %> (<% =WAPValidate(rs("LocationID")) %>)</a><% =LStyleEnd %></nobr><br/><%
				Else %>
					<% =LStyleBegin %><anchor><% =WAPValidate(rs("LocationName")) %> (<% =WAPValidate(rs("LocationID")) %>)<go href="<% =returnurl %>"><setvar name="LocationID" value="<% =WAPValidate(rs("LocationID")) %>"/></go></anchor><% =LStyleEnd %><br/><%
				End If
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASHistory()

	Dim db, sql, rs

	CardTitle = "Asset History"
	CardCurrent = "ASHistory"
	CardCurrentLevel = GetCardLevel()

	GetWOPK
	GetAssetPK
	If assetpk = "" Then
		assetpk = "-1"
	End If

	pagesize = GlobalPageSize
	If IsPocketIE or IsBlackBerry Then
		pagesize = 14
	End If

	GetPagePos

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Call GetWOSQLWhere()

	Set db = New ADOHelper

	Call WOCloseSave(db)

    sqlwhere = " WHERE WO.AssetPK =" & AssetPK & " " & sqlwhere & " "

	sql = _
	"SELECT     WO.WOPK, WO.WOID, WO.RepairCenterID, WO.WOGroupPK, WO.TargetDate, WO.Reason, WO.ProcedureName, WO.AssetID, WO.AssetName, " + nl + _
	"                      AssetHierarchy.ParentLocation, AssetHierarchy.ParentEquipment, ltp.CodeIcon AS PriorityIcon, lts.CodeIcon AS StatusIcon, " + nl + _
	"                      ltas.CodeIcon AS AuthStatusIcon, Asset.IsLocation, WO.Type, WO.TypeDesc, WO.DemoLaborPK, WO.Status, WO.AuthStatus, WO.IsGenerated, WO.RouteOrder, WO.IsAssigned, WO.PrintedBox, WO.Priority, WO.AuthLevelsRequired, WO.IsApproved, WO.RouteOrder, WO.WOGroupType, WO.TargetHours, WO.PDAWO, WO.Complete, WO.Closed " + nl

    sqlwhere = _
	"FROM         WO WITH (NOLOCK) LEFT OUTER JOIN " + nl + _
	"                      LookupTableValues ltp WITH (NOLOCK) ON WO.Priority = ltp.CodeName AND ltp.LookupTable = 'PRIORITY' INNER JOIN " + nl + _
	"                      LookupTableValues lts WITH (NOLOCK) ON WO.Status = lts.CodeName AND lts.LookupTable = 'WOSTATUS' INNER JOIN " + nl + _
	"                      LookupTableValues ltas WITH (NOLOCK) ON WO.AuthStatus = ltas.CodeName AND ltas.LookupTable = 'WOAUTHSTATUS' LEFT OUTER JOIN " + nl + _
	"                      AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK LEFT OUTER JOIN " + nl + _
	"					   Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " + nl + _
	sqlwhere + " " + nl + _
	"UNION " + nl + _
	"SELECT     WO.WOPK, WO.WOID, WO.RepairCenterID, WO.WOGroupPK, WO.TargetDate, WO.Reason, WO.ProcedureName, WO.AssetID, WO.AssetName, " + nl + _
	"                      AssetHierarchy.ParentLocation, AssetHierarchy.ParentEquipment, ltp.CodeIcon AS PriorityIcon, lts.CodeIcon AS StatusIcon, " + nl + _
	"                      ltas.CodeIcon AS AuthStatusIcon, Asset.IsLocation, WO.Type, WO.TypeDesc, WO.DemoLaborPK, WO.Status, WO.AuthStatus, WO.IsGenerated, WO.RouteOrder, WO.IsAssigned, WO.PrintedBox, WO.Priority, WO.AuthLevelsRequired, WO.IsApproved, WO.RouteOrder, WO.WOGroupType, WO.TargetHours, WO.PDAWO, WO.Complete, WO.Closed " + nl + _
	"FROM         WOTask WITH (NOLOCK) INNER JOIN " + nl + _
    "                      WO WITH (NOLOCK) ON WOTask.WOPK = WO.WOPK LEFT OUTER JOIN " + nl + _
	"                      LookupTableValues ltp WITH (NOLOCK) ON WO.Priority = ltp.CodeName AND ltp.LookupTable = 'PRIORITY' INNER JOIN " + nl + _
	"                      LookupTableValues lts WITH (NOLOCK) ON WO.Status = lts.CodeName AND lts.LookupTable = 'WOSTATUS' INNER JOIN " + nl + _
	"                      LookupTableValues ltas WITH (NOLOCK) ON WO.AuthStatus = ltas.CodeName AND ltas.LookupTable = 'WOAUTHSTATUS' LEFT OUTER JOIN " + nl + _
	"                      AssetHierarchy WITH (NOLOCK) ON WO.AssetPK = AssetHierarchy.AssetPK LEFT OUTER JOIN " + nl + _
	"					   Asset WITH (NOLOCK) ON WO.AssetPK = Asset.AssetPK " + nl + _
	REPLACE(sqlwhere,"WO.AssetPK =","WOTask.AssetPK = ") + " " + nl

	sql = sql & sqlwhere & "ORDER BY WO.TargetDate Desc,WO.RouteOrder,WO.WOPK DESC"

	Call SetSession("sqlwhere" & CardCurrentLevel,sql)

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)

	'If rs.recordcount = 1 and UCase(CardFrom) = "WOSEARCH" Then
	'	Call SetSession("WOPK" & CardFromLevel,NullCheck(rs("WOPK")))
	'	CardFromLevel = CardFromLevel + 1
	'	Call WOOptions("")
	'End If

	If UCase(CardFrom) = "WOSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("There were no Work Orders found for the specified Asset.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				%>
				<% If lang = "HTML" Then %><nobr><% End If %><% =LStyleBegin %><a href="default.asp?card=wooptionsh&amp;s=<% =SessionID %>&amp;wopk=<% =rs("WOPK") %>"><% =rs("WOPK") %>: <% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %> <% =WAPValidate(DateNullCheck(rs("CLOSED"))) & " " %> <% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) %></a><% =LStyleEnd %><% If lang = "HTML" Then %></nobr><% End If %><br/>
				<%
				rs.MoveNext()
				If Not rs.Eof Then
					TabIndex = TabIndex + 1
				End If
				If TabIndex = PageSize Then
					Exit Do
				End If
			End If
			Loop
			%>
			</p>
			<%
		End If
		%><p align="center">
		<%
		If PagePos >= PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>">Prev</a></b>&nbsp;
		<%
		End If
		If TabIndex = PageSize Then
		%><b><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>">Next</a></b>
		<%
		End If
		%></p>
		<%
		OutputButtons
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub CalendarLookup()

	Dim db, sql, rs
	Dim datToday, intThisMonth, intThisYear, strMonthName, datFirstDay, intFirstWeekDay, intLastDay, intPrevMonth, intPrintDay, LastMonthDate, NextMonthDate, dFirstDay, dLastDay, EndRows, intLoopDay, intPrevYear, intNextMonth, intNextYear, intLastMonth, dToday, bEvents

	CardTitle = "Calendar"
	CardCurrent = "CalendarLookup"
	CardCurrentLevel = GetCardLevel()

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("PagePos" & CardCurrentLevel,PagePos)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	'Call ASPDebug

	Set db = New ADOHelper

    Dim ft
    ft = Trim(Request("ft"))
    If Not ft = "" Then
        Call SetSession("CALFT",ft)
    Else
        ft = GetSession("CALFT")
    End If
    If ft = "" Then
        ft = "TargetDate"
    End If

	' Constants for the days of the week
	Const cSUN = 1, cMON = 2, cTUE = 3, cWED = 4, cTHU = 5, cFRI = 6, cSAT = 7

	' Check for valid month input
	If IsEmpty(Request("MONTH")) OR NOT IsNumeric(Request("MONTH")) Then
	  datToday = DateNullCheck(Date())
	  intThisMonth = Month(datToday)
	ElseIf CInt(Request("MONTH")) < 1 OR CInt(Request("MONTH")) > 12 Then
	  datToday = DateNullCheck(Date())
	  intThisMonth = Month(datToday)
	Else
	  intThisMonth = CInt(Request("MONTH"))
	End If

	' Check for valid year input
	If IsEmpty(Request("YEAR")) OR NOT IsNumeric(Request("YEAR")) Then
	  datToday = DateNullCheck(Date())
	  intThisYear = Year(datToday)
	Else
	  intThisYear = CInt(Request("YEAR"))
	End If

	strMonthName = MonthName(intThisMonth)
	datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
	intFirstWeekDay = WeekDay(datFirstDay, vbSunday)
	intLastDay = GetLastDay(intThisMonth, intThisYear)

	' Get the previous month and year
	intPrevMonth = intThisMonth - 1
	If intPrevMonth = 0 Then
		intPrevMonth = 12
		intPrevYear = intThisYear - 1
	Else
		intPrevYear = intThisYear
	End If

	' Get the next month and year
	intNextMonth = intThisMonth + 1
	If intNextMonth > 12 Then
		intNextMonth = 1
		intNextYear = intThisYear + 1
	Else
		intNextYear = intThisYear
	End If

	' Get the last day of previous month. Using this, find the sunday of
	' last week of last month
	LastMonthDate = GetLastDay(intLastMonth, intPrevYear) - intFirstWeekDay + 2
	NextMonthDate = 1

	' Initialize the print day to 1
	intPrintDay = 1

	' Open a record set of schedules
	Set Rs = Server.CreateObject("ADODB.RecordSet")

	' These dates are used in the SQL
	dFirstDay = intThisMonth & "/1/" & intThisYear
	dLastDay 	= intThisMonth & "/" & intLastDay & "/" & intThisYear

	'sSQL = 	"SELECT DISTINCT Start_Date, End_Date FROM tEvents WHERE " & _
	'				"(Start_Date >=" & dFirstDay & " AND Start_Date <= " & dLastDay & ") " & _
	'				"OR " & _
	'				"(End_Date >=" & dFirstDay & " AND End_Date <= " & dLastDay & ") " & _
	'				"OR " & _
	'				"(Start_Date < " & dFirstDay & " AND End_Date > " & dLastDay & " )"  & _
	'				"ORDER BY Start_Date"
	'Response.Write sSQL

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
	'Rs.Open sSQL, sDSN, adOpenStatic, adLockReadOnly, adCmdText

	Call StartMobileDocument(CardTitle)
    %>
	<p align="center" mode="nowrap">
    <a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;month=<% =IntPrevMonth %>&amp;year=<% =IntPrevYear %>"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/prev.gif" alt="Prev"/></a> <b><% = strMonthName & " " & intThisYear %></b> <a href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;month=<% =IntNextMonth %>&amp;year=<% =IntNextYear %>"><img<% If lang="HTML" Then %> border="0"<% End If %> src="images/next.gif" alt="Next"/></a><br/>
    <% =LStyleBegin %>Su Mo Tu We Th Fr Sa <% =LStyleEnd %>
	<%
			' Initialize the end of rows flag to false
			EndRows = False
			'Response.Write vbCrLf

			' Loop until all the rows are exhausted
		 	Do While EndRows = False
				' Start a table row
				Response.Write "	<br/>" & vbCrLf
				' This is the loop for the days in the week
				For intLoopDay = cSUN To cSAT
					' If the first day is not sunday then print the last days of previous month in grayed font
					If intFirstWeekDay > cSUN Then
						Write_CalDay LStyleBegin & CheckDigits(LastMonthDate) & " " & LStyleEnd, "NON"
						LastMonthDate = LastMonthDate + 1
						intFirstWeekDay = intFirstWeekDay - 1
					' The month starts on a sunday
					Else
						' If the dates for the month are exhausted, start printing next month's dates
						' in grayed font
						If intPrintDay > intLastDay Then
							Write_CalDay LStyleBegin & CheckDigits(NextMonthDate) & " " & LStyleEnd, "NON"
							NextMonthDate = NextMonthDate + 1
							EndRows = True
						Else
							' If last day of the month, flag the end of the row
							If intPrintDay = intLastDay Then
								EndRows = True
							End If

							dToday = CDate(intThisMonth & "/" & intPrintDay & "/" & intThisYear)
							'If NOT Rs.EOF Then
							If False Then
								' Set events flag to false. This means the day has no event in it
								bEvents = False
							  Do While NOT Rs.EOF AND bEvents = False
									' If the date falls within the range of dates in the recordset, then
									' the day has an event. Make the events flag True
							    If dToday >= Rs("Start_Date") AND dToday <= Rs("End_Date") Then
										' Print the date in a highlighted font
							      Write_CalDay "<a href=events.asp?date="& Server.URLEncode(dToday) & " CLASS='EVENT' TARGET='rightframe'> " & intPrintDay & "</a>", "HL"
										bEvents = True
									' If the Start date is greater than the date itself, there is no point
									' checking other records. Exit the loop
							    ElseIf dToday < Rs("Start_Date") Then
										Exit Do
									' Move to the next record
									Else
								    Rs.MoveNext
									End If
							  Loop
								' Checks for that day
								Rs.MoveFirst
							End If

							' If the event flag is not raise for that day, print it in a plain font
							If bEvents = False Then

				                returnurl = "default.asp?card=" & GetSession("ParentCard" & (CardCurrentLevel-1)) & "&amp;s=" & SessionID & "&amp;back=1&amp;" & ft & "=" & WAPEncode(dToday)
				                If lang="HTML" Then
					                returnurl = "javascript:parent.document.mcform." & ft & ".value='" & JSEncode(dToday) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					                <nobr><% =LStyleBegin %><a href="<% =returnurl %>"><% If dToday = Date Then %><b><% End If %><% =CheckDigits(intPrintDay) %><% If dToday = Date Then %></b><% End If %></a><% =LStyleEnd %> </nobr> <%
				                Else %>
					                <% =LStyleBegin %><% If dToday = Date Then %><b><% End If %><anchor><% =CheckDigits(intPrintDay) %><go href="<% =returnurl %>"><setvar name="<% =ft %>" value="<% =WAPEncode(dToday) %>"/></go></anchor> <% If dToday = Date Then %></b><% End If %><% =LStyleEnd %> <%
				                End If

							End If
						End If

						' Increment the date. Done once in the loop.
						intPrintDay = intPrintDay + 1
					End If

				' Move to the next day in the week
				Next

			Loop
	%>
	</p>
    <%
		OutputButtons
		'rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub
'====================================================================================================================================

Sub WOSave(db)

    'If InStr(GetSession("pd_wosave"),"," & Request("preventduplicationsubmit") & ",") > 0 Then
    '    ' User hit Submit 2x or more - so Exit
    '    Response.Write "HERE"
    '    Response.End
    '    Exit Sub
    'Else
    '    Call SetSession("pd_wosave",GetSession("wosave")&","&Trim(Request("preventduplicationsubmit"))&",")
    'End If

	Dim iswonew,rs2
	iswonew = Request("wonew")
	If iswonew = "" Then
		Exit Sub
	End If

	'Call ASPDebug
	'Response.End

	Dim sql, rs, rsTasks, OutArray
	Dim RCPK,RCID,RCNM
	Dim ShopPK, ShopID, ShopName
	Dim SupervisorPK, SupervisorID, SupervisorName
	Dim DepartmentPK, DepartmentID, DepartmentName
	Dim TenantPK, TenantID, TenantName
	Dim AccountPK, AccountID, AccountName
	Dim AssetPK, AssetID, AssetName
	Dim LaborPK, LaborID, LaborName
	Dim CategoryPK, CategoryID, CategoryName
	Dim ProblemPK, ProblemID, ProblemName
	Dim ProcedurePK, ProcedureID, ProcedureName
	Dim Reference, ReferenceDesc, TargetHours, ProjectPK, ProjectID, ProjectName, Instructions, LockoutTagoutBox, AttachmentsBox, ChargeableBox, ShutdownBox, FollowupWork
	Dim prefvalue, prefdesc, prefpk
	Dim RequesterPhone, RequesterEmail
	Dim WarrantyBox
	Dim txtType, txtTypeDesc, txtPriority, txtPriorityDesc
	Dim txtTaskPKs

	AssetPK = ""
	AssetID = ""
	AssetName = ""

	RCPK = ""
	RCID = ""
	RCNM = ""

	ShopPK = ""
	ShopID = ""
	ShopName = ""

	SupervisorPK = ""
	SupervisorID = ""
	SupervisorName = ""

	DepartmentPK = ""
	DepartmentID = ""
	DepartmentName = ""

	TenantPK = ""
	TenantID = ""
	TenantName = ""

	AccountPK = ""
	AccountID = ""
	AccountName = ""

	CategoryPK = ""
	CategoryID = ""
	CategoryName = ""

	ProblemPK = ""
	ProblemID = ""
	ProblemName = ""

	ProcedurePK = ""
	ProcedureID = ""
	ProcedureName = ""

	Reference = ""
	ReferenceDesc = ""

	TargetHours = "1"

	ProjectPK = ""
	ProjectID = ""
	ProjectName = ""

	Instructions = ""

	RequesterPhone = ""
	RequesterEmail = ""

	LockoutTagoutBox = 0
	AttachmentsBox = 0
	ChargeableBox = 0
	ShutdownBox = 0

	WarrantyBox = 0
	txtTaskPKs = ""

	If Request("Reason").Item = "" Then
		HeaderMSG = "The Reason can not be left blank."
		Call WONew()
	End If

	If Not NullCheck(Request("AssetID")) = "" Then
		sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetID = '" & NullCheck(Request("AssetID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Asset ID was not found."
			Call WONew()
		Else
			AssetPK = rs2("AssetPK")
			AssetID = rs2("AssetID")
			AssetName = rs2("AssetName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Asset ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("ProblemID")) = "" Then
		sql = "SELECT * FROM Failure WITH (NOLOCK) WHERE FailureID = '" & NullCheck(Request("ProblemID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Problem ID was not found."
			Call WONew()
		Else
			ProblemPK = rs2("FailurePK")
			ProblemID = rs2("FailureID")
			ProblemName = rs2("FailureName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Problem ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("ProcedureID")) = "" Then
		sql = "SELECT * FROM ProcedureLibrary WITH (NOLOCK) WHERE ProcedureID = '" & NullCheck(Request("ProcedureID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Procedure ID was not found."
			Call WONew()
		Else
			ProcedurePK = rs2("ProcedurePK")
			ProcedureID = rs2("ProcedureID")
			ProcedureName = rs2("ProcedureName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Procedure ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("CategoryID")) = "" Then
		sql = "SELECT * FROM Category WITH (NOLOCK) WHERE CategoryID = '" & NullCheck(Request("CategoryID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Category ID was not found."
			Call WONew()
		Else
			CategoryPK = rs2("CategoryPK")
			CategoryID = rs2("CategoryID")
			CategoryName = rs2("CategoryName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Category ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("AccountID")) = "" Then
		sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Account ID was not found."
			Call WONew()
		Else
			AccountPK = rs2("AccountPK")
			AccountID = rs2("AccountID")
			AccountName = rs2("AccountName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Account ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("Priority")) = "" Then
		sql = "SELECT * FROM LookupTableValues WITH (NOLOCK) WHERE LookupTable = 'WOPriority' AND CodeName = '" & NullCheck(Request("Priority")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Priority was not found."
			Call WONew()
		Else
			txtPriority = rs2("CodeName")
			txtPriorityDesc = rs2("CodeDesc")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Priority is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("Type")) = "" Then
		sql = "SELECT * FROM LookupTableValues WITH (NOLOCK) WHERE LookupTable = 'WOType' AND CodeName = '" & NullCheck(Request("Type")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Type was not found."
			Call WONew()
		Else
			txtType = rs2("CodeName")
			txtTypeDesc = rs2("CodeDesc")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Type is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("ProjectID")) = "" Then
		sql = "SELECT * FROM Project WITH (NOLOCK) WHERE ProjectID = '" & NullCheck(Request("ProjectID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Project ID was not found."
			Call WONew()
		Else
			ProjectPK = rs2("ProjectPK")
			ProjectID = rs2("ProjectID")
			ProjectName = rs2("ProjectName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Project ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("DepartmentID")) = "" Then
		sql = "SELECT * FROM Department WITH (NOLOCK) WHERE DepartmentID = '" & NullCheck(Request("DepartmentID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Department ID was not found."
			Call WONew()
		Else
			DepartmentPK = rs2("DepartmentPK")
			DepartmentID = rs2("DepartmentID")
			DepartmentName = rs2("DepartmentName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Department ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("TenantID")) = "" Then
		sql = "SELECT * FROM Tenant WITH (NOLOCK) WHERE TenantID = '" & NullCheck(Request("TenantID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Customer ID was not found."
			Call WONew()
		Else
			TenantPK = rs2("TenantPK")
			TenantID = rs2("TenantID")
			TenantName = rs2("TenantName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Customer ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("ShopID")) = "" Then
		sql = "SELECT * FROM Shop WITH (NOLOCK) WHERE ShopID = '" & NullCheck(Request("ShopID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Shop ID was not found."
			Call WONew()
		Else
			ShopPK = rs2("ShopPK")
			ShopID = rs2("ShopID")
			ShopName = rs2("ShopName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Shop ID is invalid."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("LaborID")) = "" Then
		sql = "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborID = '" & NullCheck(Request("LaborID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "The Assigned ID was not found."
			Call WONew()
		Else
			LaborPK = rs2("LaborPK")
			LaborID = rs2("LaborID")
			LaborName = rs2("LaborName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "The value provided for Assigned ID is invalid."
			Call WONew()
		End If
	End If

	If Not AssetPK = "" Then

		Set rs = db.RunSPReturnMultiRS("MC_ValAsset",Array(Array("@AssetPK", adInteger, adParamInput, 4, AssetPK)),"")
		Call CheckDB(db)
		If Not rs.Eof Then

			If isDate(NullCheck(rs("WarrantyExpire"))) Then
				If DateDiff("d",rs("WarrantyExpire"),Date()) <=1 Then
					WarrantyBox=1
				Else
					WarrantyBox=0
				End If
			Else
				WarrantyBox=0
			End If

			Set rs = rs.NextRecordset()

			' Update Repair Center (if not null)
			If Not NullCheck(rs("RepairCenterPK")) = "" Then
				RCPK = NullCheck(rs("RepairCenterPK"))
				RCID = NullCheck(rs("RepairCenterID"))
				RCNM = NullCheck(rs("RepairCenterName"))
			End If

			' Update Shop (if not null)
			If Not NullCheck(rs("ShopPK")) = "" and ShopPK = "" Then
				ShopPK = NullCheck(rs("ShopPK"))
				ShopID = NullCheck(rs("ShopID"))
				ShopName = NullCheck(rs("ShopName"))
			End If

			' Update Supervisor (if not null)
			If Not NullCheck(rs("SupervisorPK")) = "" Then
				SupervisorPK = NullCheck(rs("SupervisorPK"))
				SupervisorID = NullCheck(rs("SupervisorID"))
				SupervisorName = NullCheck(rs("SupervisorName"))
			End If

			' Update Department (if not null)
			If Not NullCheck(rs("DepartmentPK")) = "" and DepartmentPK = "" Then
				DepartmentPK = NullCheck(rs("DepartmentPK"))
				DepartmentID = NullCheck(rs("DepartmentID"))
				DepartmentName = NullCheck(rs("DepartmentName"))
			End If

			' Update Tenant (if not null)
			If Not NullCheck(rs("TenantPK")) = "" and TenantPK = "" Then
				TenantPK = NullCheck(rs("TenantPK"))
				TenantID = NullCheck(rs("TenantID"))
				TenantName = NullCheck(rs("TenantName"))
			End If

			' Update Account (if not null)
			If Not NullCheck(rs("AccountPK")) = "" and AccountPK = "" Then
				AccountPK = NullCheck(rs("AccountPK"))
				AccountID = NullCheck(rs("AccountID"))
				AccountName = NullCheck(rs("AccountName"))
			End If

		End If

	End If

	If Not ProblemPK = "" Then

		sql="Select f.FailurePK,f.FailureID,f.FailureName,f.ProcedurePK,f.ProcedureID,f.ProcedureName " & _
		    "FROM Failure f WITH (NOLOCK) " & _
		    "WHERE FailurePK = " & ProblemPK

		Set rs = db.RunSqlReturnRS(sql,"")
		Call CheckDB(db)

		If Not rs.Eof Then

            If ProcedurePK = "" Then
			    ProcedurePK = rs("ProcedurePK")
			    ProcedureID = rs("ProcedureID")
			    ProcedureName = rs("ProcedureName")
			End If

	    End If

	End If

	If UCase(NullCheck(Request("Chargeable"))) = "Y" or UCase(NullCheck(Request("Chargeable"))) = "2" Then
		ChargeableBox = True
	Else
		ChargeableBox = False
	End If
	If UCase(NullCheck(Request("ShutdownBox"))) = "Y" or UCase(NullCheck(Request("ShutdownBox"))) = "2" Then
		ShutdownBox = True
	Else
		ShutdownBox = False
	End If
	If UCase(NullCheck(Request("FollowupWork"))) = "Y" or UCase(NullCheck(Request("FollowupWork"))) = "2" Then
		FollowupWork = True
	Else
		FollowupWork = False
	End If
	If UCase(NullCheck(Request("LockoutTagoutBox"))) = "Y" or UCase(NullCheck(Request("LockoutTagoutBox"))) = "2" Then
		LockoutTagoutBox = True
	Else
		LockoutTagoutBox = False
	End If

    Dim RCPreference
	If RCPK = "" Then
		RCPreference = GetSession("RCPK")	' Nullable: YES Type: int (from session)
	Else
		RCPreference = RCPK	' Nullable: YES Type: int (from session)
	End If

	If Not ProcedurePK = "" Then

		sql="Select " & _
		    "p.Reference,p.ReferenceDesc,p.TargetHours,p.ProjectPK,p.ProjectID,p.ProjectName,p.Instructions," & _
		    "p.LockoutTagoutBox,p.AttachmentsBox,p.Chargeable,p.ShutdownBox, p.CategoryPK, p.CategoryID, p.CategoryName " & _
		    "FROM ProcedureLibrary p WITH (NOLOCK) " & _
		    "WHERE ProcedurePK = " & ProcedurePK

		Set rs = db.RunSqlReturnRS(sql,"")
		Call CheckDB(db)

		If Not rs.Eof Then

			Reference = rs("Reference")
			ReferenceDesc = rs("ReferenceDesc")

			TargetHours = rs("TargetHours")

            If Not NullCheck(rs("ProjectPK")) = "" and ProjectPK = "" Then
			    ProjectPK = rs("ProjectPK")
			    ProjectID = rs("ProjectID")
			    ProjectName = rs("ProjectName")
			End If

			If Not NullCheck(rs("CategoryPK")) = "" and CategoryPK = "" Then
				CategoryPK = rs("CategoryPK")
				CategoryID = rs("CategoryID")
				CategoryName = rs("CategoryName")
			End If

			Instructions = rs("Instructions")

			LockoutTagoutBox = BitNullCheck(rs("LockoutTagoutBox"))
			AttachmentsBox = BitNullCheck(rs("AttachmentsBox"))
			ChargeableBox = BitNullCheck(rs("Chargeable"))
			ShutdownBox = BitNullCheck(rs("ShutdownBox"))

		End If

	End If

    If Request("TargetHours") = "" Then
	    If GetPreference(db,False,RCPreference,"WO_DefaultTargetHours",prefvalue, prefdesc, prefpk) Then
    		TargetHours = prefvalue
    	End If
    Else
        TargetHours = Request("TargetHours")
    End If

    If Request("TargetHours") = "" Then
	    If GetPreference(db,False,RCPreference,"WO_DefaultTargetHours",prefvalue, prefdesc, prefpk) Then
    		TargetHours = prefvalue
    	End If
    Else
        TargetHours = Request("TargetHours")
    End If

    If txtPriority = "" Then
        If GetPreference(db,False,RCPreference,"WO_DEFAULTPRIORITY",prefvalue, prefdesc, prefpk) Then
	        txtPriority = prefvalue
        Else
	        txtPriority = "2"
        End If
    End If

    If txtType = "" Then
        If GetPreference(db,False,RCPreference,"WO_DEFAULTTYPE",prefvalue, prefdesc, prefpk) Then
	        txtType = prefvalue
        Else
	        txtType = "S"
        End If
    End If

	sql="Select CodeDesc from LookUpTableValues WITH (NOLOCK) WHERE LookUpTable = 'WOTYPE' AND CodeName='" + txtType +"'"
	Set rs = db.RunSqlReturnRS(sql,"")
	Call CheckDB(db)

	If Not rs.Eof Then
		txtTypeDesc=NullCheck(rs("CodeDesc"))
	Else
		txtTypeDesc=""
	End If

	sql="Select CodeDesc from LookUpTableValues WITH (NOLOCK) WHERE LookUpTable = 'WOPRIORITY' AND CodeName='" + txtPriority +"'"
	Set rs = db.RunSqlReturnRS(sql,"")
	Call CheckDB(db)

	If Not rs.Eof Then
		txtPriorityDesc=NullCheck(rs("CodeDesc"))
	Else
		txtPriorityDesc=""
	End If

	sql = "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborPK = " & GetSession("UserPK") & " "
	Set rs2 = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)
	If Not rs2.eof Then
		RequesterPhone = NullCheck(rs2("PhoneWork"))
		RequesterEmail = NullCheck(rs2("Email"))
	End If

	' See if the Requester is associated with a Department
	sql = "Select DepartmentPK, DepartmentID, DepartmentName From Labor WITH (NOLOCK) Where LaborPK = " & GetSession("UserPK")
	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)
	If Not rs.eof Then
		If Not NullCheck(rs("DepartmentPK")) = "" Then
			If DepartmentPK = "" Then
				DepartmentPK = NullCheck(rs("DepartmentPK"))
				DepartmentID = NullCheck(rs("DepartmentID"))
				DepartmentName = NullCheck(rs("DepartmentName"))
			End If
		End If
	End If

	'On Error Resume Next

	'Set rs = db.RunSqlReturnRS("Select GETDATE() AS ServerDate","")
	'Call CheckDB(db)
	'ServerDate = rs("ServerDate")

	Set rs = db.RunSQLReturnRS_RW("SELECT TOP 0 * FROM WO","")
	Call CheckDB(db)

	On Error Resume Next

	rs.AddNew

	rs("WOID") = "" ' Nullable: No Type: nvarchar

	If Not Request("Reason").Item = "" Then
		rs("Reason") = SEZ(Trim(Mid(Replace(Request("Reason").Item,Chr(13)+Chr(10),"%0D%0A"),1,2000)))	' Nullable: YES Type: nvarchar
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Reason is invalid."
	    Call WONew()
    End If

	rs("Status") = Trim(Mid("REQUESTED",1,15))	' Nullable: No Type: nvarchar
	rs("StatusDesc") = Trim(Mid("Requested",1,50))	' Nullable: YES Type: nvarchar
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Status is invalid."
	    Call WONew()
    End If

	If GetSession("WOAuthReq") = "0" Then
		rs("AuthStatus") = Trim(Mid("NOTREQUIRED",1,15))	' Nullable: No Type: nvarchar
		rs("AuthStatusDesc") = Trim(Mid("(not required)",1,50))	' Nullable: YES Type: nvarchar
	Else
		rs("AuthStatus") = Trim(Mid("REQUIRED" & GetSession("WOAuthReq"),1,15))	' Nullable: No Type: nvarchar
		rs("AuthStatusDesc") = Trim(Mid("Required - Level " & GetSession("WOAuthReq"),1,50))	' Nullable: YES Type: nvarchar
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Auth Status is invalid."
	    Call WONew()
    End If

	rs("StatusDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))	' Nullable: YES Type: datetime
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Status Date is invalid."
	    Call WONew()
    End If

	rs("TakenByPK") = GetSession("UserPK")	' Nullable: YES Type: int
	rs("TakenByID") = GetSession("UserID")	' Nullable: YES Type: nvarchar
	rs("TakenByName") = GetSession("UserName")	' Nullable: YES Type: nvarchar

	rs("RequesterPK") = GetSession("UserPK")	' Nullable: YES Type: int
	rs("RequesterID") = GetSession("UserID")	' Nullable: YES Type: nvarchar
	rs("RequesterName") = GetSession("UserName")	' Nullable: YES Type: nvarchar
	rs("RequesterEmail") = RequesterEmail	' Nullable: YES Type: nvarchar
	rs("RequesterPhone") = RequesterPhone	' Nullable: YES Type: nvarchar
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Requester ID is invalid."
	    Call WONew()
    End If

	If Not AssetPK = "" Then
		rs("AssetPK") = NullCheck(AssetPK)		' Nullable: YES Type: int
		rs("AssetID") = NullCheck(AssetID)		' Nullable: YES Type: nvarchar
		rs("AssetName") = NullCheck(AssetName)	' Nullable: YES Type: nvarchar
	Else
		' Put the typed in Location in the Instructions field
		'rs("Instructions") = Trim(Mid(Replace(txtLocationManual,Chr(13)+Chr(10),"%0D%0A"),1,1000))
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Asset ID is invalid."
	    Call WONew()
    End If

    If Request("TargetDate") = "" Then
	    rs("TargetDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))
	Else
	    rs("TargetDate") = SQLdatetimeADO(Request("TargetDate"))
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Target Date is invalid."
	    Call WONew()
    End If

	rs("Type") = Trim(Mid(txtType,1,25))	' Nullable: YES Type: nvarchar
	rs("TypeDesc") = NullCheck(txtTypeDesc)	' Nullable: YES Type: nvarchar
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Type is invalid."
	    Call WONew()
    End If

	rs("Priority") = Trim(Mid(txtPriority,1,25))	' Nullable: No Type: nvarchar
	rs("PriorityDesc") = NullCheck(txtPriorityDesc)	' Nullable: YES Type: nvarchar
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Priority is invalid."
	    Call WONew()
    End If

	If RCPK = "" Then
		rs("RepairCenterPK") = GetSession("RCPK")	' Nullable: YES Type: int (from session)
		rs("RepairCenterID") = GetSession("RCID")	' Nullable: YES Type: nvarchar (from session)
		rs("RepairCenterName") = GetSession("RCNM")	' Nullable: YES Type: nvarchar
	Else
		rs("RepairCenterPK") = RCPK	' Nullable: YES Type: int (from session)
		rs("RepairCenterID") = RCID	' Nullable: YES Type: nvarchar (from session)
		rs("RepairCenterName") = RCNM	' Nullable: YES Type: nvarchar
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Repair Center ID is invalid."
	    Call WONew()
    End If

	If ShopPK = "" Then
		If GetPreference(db,True,RCPreference,"WO_DefaultShop",prefvalue, prefdesc, prefpk) Then
			rs("ShopPK") = prefpk
			rs("ShopID") = prefvalue
			rs("ShopName") = prefdesc
		End If
	Else
		rs("ShopPK") = ShopPK
		rs("ShopID") = ShopID
		rs("ShopName") = ShopName
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Shop ID is invalid."
	    Call WONew()
    End If

	If SupervisorPK = "" Then
		If GetPreference(db,True,RCPreference,"WO_DefaultSupervisor",prefvalue, prefdesc, prefpk) Then
			rs("SupervisorPK") = prefpk
			rs("SupervisorID") = prefvalue
			rs("SupervisorName") = prefdesc
		End If
	Else
		rs("SupervisorPK") = SupervisorPK
		rs("SupervisorID") = SupervisorID
		rs("SupervisorName") = SupervisorName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "The value provided for Supervisor ID is invalid."
		Call WONew()
	End If

	If Not DepartmentPK = "" Then
		rs("DepartmentPK") = DepartmentPK
		rs("DepartmentID") = DepartmentID
		rs("DepartmentName") = DepartmentName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "The value provided for Department ID is invalid."
		Call WONew()
	End If

	If Not TenantPK = "" Then
		rs("TenantPK") = TenantPK
		rs("TenantID") = TenantID
		rs("TenantName") = TenantName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "The value provided for Customer ID is invalid."
		Call WONew()
	End If

	If Not AccountPK = "" Then
		rs("AccountPK") = AccountPK
		rs("AccountID") = AccountID
		rs("AccountName") = AccountName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "The value provided for Account ID is invalid."
		Call WONew()
	End If

	If Not CategoryPK = "" Then
		rs("CategoryPK") = CategoryPK
		rs("CategoryID") = CategoryID
		rs("CategoryName") = CategoryName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "The value provided for Category ID is invalid."
		Call WONew()
	End If

	If CLng(TargetHours) > 0 Then
		rs("TargetHours") = TargetHours
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "The value provided for Target Hours is invalid."
		Call WONew()
	End If

	If GetPreference(db,True,RCPreference,"WO_DefaultSurvey",prefvalue, prefdesc, prefpk) Then
		rs("SurveyBox") = 1
		rs("Survey_ID") = prefpk
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Survey is invalid."
	    Call WONew()
    End If

	rs("Requested") = SQLdatetimeADO(DateTimeNullCheck(Now()))
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Requested is invalid."
	    Call WONew()
    End If

	If Not ProblemPK = "" Then

		rs("ProblemPK") = ProblemPK
		rs("ProblemID") = ProblemID
		rs("ProblemName") = ProblemName

	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Problem ID is invalid."
	    Call WONew()
    End If

	If Not ProcedurePK = "" Then

		rs("ProcedurePK") = ProcedurePK
		rs("ProcedureID") = ProcedureID
		rs("ProcedureName") = ProcedureName

    End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Procedure ID is invalid."
	    Call WONew()
    End If

	If Not Reference = "" Then
		rs("Reference") = Reference
		rs("ReferenceDesc") = ReferenceDesc
	End If

	If Not ProjectPK = "" Then
		rs("ProjectPK") = ProcedurePK
		rs("ProjectID") = ProcedureID
		rs("ProjectName") = ProcedureName
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Project ID is invalid."
	    Call WONew()
    End If

	' Special Instructions on the Asset Record do NOT get copied to the
	' Special Instrucitons Field - but still show up when printing WOs.
	' Special Instructions on the Procedure DO get copied over.

	If Not Instructions = "" Then
		If IsNull(rs("Instructions")) or rs("Instructions") = "" Then
			rs("Instructions") = Trim(Instructions)
		Else
			rs("Instructions") = Trim(rs("Instructions")) & " / " & Trim(Instructions)
		End If
	End If

    rs("ShutdownBox") = ShutdownBox
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Shutdown is invalid."
	    Call WONew()
    End If
    rs("WarrantyBox") = WarrantyBox
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Warranty is invalid."
	    Call WONew()
    End If
    rs("FollowupWork") = FollowupWork
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Follow-up Work is invalid."
	    Call WONew()
    End If
    rs("LockoutTagoutBox") = LockoutTagoutBox
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Lockout/Tagout is invalid."
	    Call WONew()
    End If
    rs("Chargeable") = ChargeableBox
    If Err.Number <> 0 Then
	    HeaderMSG = "The value provided for Chargeable is invalid."
	    Call WONew()
    End If

	rs("RowVersionAction") = "SR"

	If DemoMode() Then
		' This makes it so that only the records the demo user adds
		' will show up in their Explorer Lists. It also prevents them from
		' seeing other demo user's records - which may contain
		' private email addresses and / or other data.
		rs("DemoLaborPK") = GetSession("UserPK")
	End If

	rs("RowVersionIPAddress") = GetSession("UserIPAddress")	' Nullable: YES Type: int
	rs("RowVersionUserPK") = GetSession("UserPK")	' Nullable: YES Type: int
	rs("RowVersionInitials") = Trim(Mid(GetSession("UserInitials"),1,5))									' Nullable: No Type: nvarchar
	rs("RowVersionDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))

    On Error Goto 0

	db.dobatchupdate rs

	Call CheckDB(db)

	WOPK = rs("WOPK")

	' Assignments
	If Not LaborPK = "" Then

		sql = _
		  "INSERT INTO WOAssign " &_
          "(IsAssigned, WOPK, WOlaborPK, LaborPK, AssignedHours, AssignedDate, AssignedLead, AssignedPDA, DemoLaborPK, RowVersionIPAddress, RowVersionUserPK,  " &_
          "RowVersionInitials, RowVersionDate) " &_
          "SELECT 1,WO.WOPK,null," & LaborPK & "," & "1" & ",getDate()," & "1" & "," & "1" & ",null,'" & GetSession("UserIPAddress") & "'," & GetSession("UserPK") & ",'" & Trim(Mid(GetSession("UserInitials"),1,5)) & "',getDate() " + _
          "FROM WO WITH (NOLOCK) WHERE WO.WOPK = " & WOPK & " OR " + _
					"(WO.WOGroupPK IN " + _
					"(SELECT WO.WOGroupPK FROM WO WITH (NOLOCK) " + _
					"WHERE WO.WOPK = " & WOPK & " AND WO.WOGroupType = 'M')) "

		Call db.RunSQL(sql,"")
		Call CheckDB(db)

	End If

	' Auto-Assignments
	Call db.RunSP("MC_WOAutoAssign",Array(Array("@WOPK", adInteger, adParamInput, 4, WOPK)),"")
	Call CheckDB(db)

	If Not txtTaskPKs = "" Then

		txtTaskPKs = Left(txtTaskPKs,Len(txtTaskPKs)-1)
		txtTaskPKs = Replace(txtTaskPKs,"TASK_","")

		sql = "SELECT TaskPK, CategoryName, TaskGroup, TaskAction " & _
		      "FROM Task WITH (NOLOCK) " & _
			  "WHERE TaskPK IN (" & txtTaskPKs & ") " & _
			  "ORDER BY CategoryName, TaskGroup, DisplayOrder, TaskAction "

		Set rsTasks = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)

		If Not rsTasks.eof Then

			Set rs = db.RunSQLReturnRS_RW("SELECT TOP 0 * FROM WOTask","")
			Dim TaskNo
			TaskNo = 1000
			Call CheckDB(db)

			Dim LastHeader, IsHeader
			LastHeader = "!"

			Dim LastGroup, IsGroup
			LastGroup = "!"

			Do While Not rsTasks.eof

				IsHeader = False
				IsGroup = False

				rs.AddNew

				If Not NullCheck(rsTasks("CategoryName")) = LastHeader and _
				   Not NullCheck(rsTasks("CategoryName")) = "" Then

				   LastHeader = Trim(rsTasks("CategoryName"))
				   IsHeader = True
				   LastGroup = "!"

				Else

					If Not NullCheck(rsTasks("TaskGroup")) = LastGroup and _
					   Not NullCheck(rsTasks("TaskGroup")) = "" Then

					   LastGroup = Trim(rsTasks("TaskGroup"))
					   IsGroup = True

					End If

				End If

				rs("WOPK") = WOPK

				rs("TaskNo") = TaskNo		' Nullable: YES Type: int

				If IsHeader Then
					rs("TaskAction") = Trim(LastHeader)	' Nullable: YES Type: nvarchar
				ElseIf IsGroup Then
					rs("TaskAction") = Proper(Trim(LastGroup))	' Nullable: YES Type: nvarchar
				Else
					If Not NullCheck(rsTasks("TaskAction")) = "" Then
						rs("TaskAction") = Trim(rsTasks("TaskAction"))	' Nullable: YES Type: nvarchar
					End If
				End If

				rs("Complete") = False					' Nullable: YES Type: bit
				rs("Fail") = False						' Nullable: YES Type: bit
				rs("Header") = IsHeader or IsGroup		' Nullable: No Type: bit

				If IsGroup and False Then
					rs("LineStyle") = "B"			' Nullable: YES Type: nvarchar
					rs("LineStyleDesc") = "Bold"	' Nullable: YES Type: nvarchar
				End If

				rs("TaskPK") = rsTasks("TaskPK").Value

				rs("RowVersionIPAddress") = GetSession("UserIPAddress")	' Nullable: YES Type: int
				rs("RowVersionUserPK") = GetSession("UserPK")	' Nullable: YES Type: int
				rs("RowVersionInitials") = Trim(Mid(GetSession("UserInitials"),1,5))									' Nullable: No Type: nvarchar
				rs("RowVersionDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))

				If Not IsHeader and Not IsGroup Then
					rsTasks.MoveNext()
				End If

				TaskNo = TaskNo + 10

			Loop

			db.dobatchupdate rs
			Call CheckDB(db)

		End If

	End If

End Sub
%>