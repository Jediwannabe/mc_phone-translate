<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>
<!--#INCLUDE FILE="includes/mc_pda_init.asp" -->
<!--#INCLUDE FILE="includes/mc_pda_authenticate.asp" -->
<!--#INCLUDE FILE="includes/mc_pda_support.asp" -->
<%
'====================================================================================================
' SET DEBUG MODE
'====================================================================================================
debug = True
dim headerString
If GetSession("logintime") = "" Then
	Call SetSession("logintime",Now())
End If

If LCase(card) <> "login" Then
	On Error Resume Next
	'If Clng(DateDiff("n", CDate(GetSession("logintime")), Now())) > CLng(Application("SessionTimeout")) Then
	If Clng(DateDiff("n", CDate(GetSession("logintime")), Now())) > 60 Then
		Response.Redirect("default.asp")
	End If

	If Err.Number <> 0 Then
		On Error Goto 0
		Response.Redirect("default.asp")
	End If
	On Error Goto 0
End If

Select Case LCase(card)

case "login"

	Dim inputfieldtype
	Call StartMobileDocument("Maintenance Connection")
	txtmembername = Request.Cookies("m")
	inputfieldtype = "password"
		%>
		<script type="text/javascript">
			$(document).ready(function(){
			});
		</script>

			<div class="Font1" style='wposition:relative; text-align:center;'>
				<table cellpadding='2' cellspacing='0'>
					<tr>
						<td>
							<img src="images/logo_color_png.png" alt="" title="" />
						</td>
					</tr>
					<tr>
						<td>
							<div class="Font2" style='font-size:14pt; padding-top:40px; text-align:left;'>Was ist Ihre Benutzer ID?</div>
						</td>
						
					</tr>
					<tr>
						<td>
							<input style='width:300px;' class="Textbox_Normal" value="<% =txtmembername %>" tabindex="1" type="text" id="membername" name="membername" format="*M"/><br/>
						</td>
					</tr>
					<tr>
						<td>
							<div class="Font2" style='font-size:14pt; width;300px;'>Was ist Ihre Passwort?</div>
						</td>
					</tr>
					<tr>
						<td>
							<input style='width:300px;' class="Textbox_Normal" value="<% =txtpassword %>" tabindex="2" type="<% =inputfieldtype %>" id="password" name="password" format="*M"/>
						</td>
					</tr>
					<tr>
						<td>
							<div style='padding-top:20px;'>
								<input tabindex="-1" style="width:300px;" type="submit" name="login" value="Login"/>
								<input type="hidden" name="card" value="authenticate"/>
							</div>
						</td>
					</tr>

				</table>
			</div>


	<%
	EndWMLDocument

Case "wotaskcomplete"
	WOPK = Request("wopk")
	If WOPK <> "" Then
		Call db.RunSQL("UPDATE WOTASK WITH ( ROWLOCK ) SET Complete = 1 WHERE PK = " & Request("taskid") & " AND WOPK = " & WOPK,"")
	End If
	wotasks
Case "wotaskuncomplete"
	WOPK = Request("wopk")
	If WOPK <> "" Then
		Call db.RunSQL("UPDATE WOTASK WITH ( ROWLOCK ) SET Complete = 0 WHERE PK = " & Request("taskid") & " AND WOPK = " & WOPK,"")
	End If
	wotasks
case "wodetails"
		wodetails

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
%>
	<div class="Font1" style='wposition:relative; text-align:center;'>
		<table cellpadding='2' cellspacing='0'>
			<tr>
				<td>
					<img src="images/logo_color_png.png" alt="" title="" />
				</td>
			</tr>
			<tr>
				<td>
					<div class="Font2" style='font-size:14pt; padding-top:40px; text-align:center;'>Danke f&#252;r die Benutzung<br />Maintenance Connection Mobile.</div>
				</td>
			</tr>
			<tr>
				<td>
					<div style='padding-top:20px;'>
						<input onclick="location.href = 'default.asp?card=login';" tabindex="-1" style="width:300px;" type="submit" name="login" value="Login"/>
					</div>
				</td>
			</tr>

		</table>
	</div>
<%
Case Else

	Call StartMobileDocument("")
		OutputWAPMsg("Die Karte wurde nicht gefunden: " & card)
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
	CardTitle = "Startseite"
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
	headerString ="<div onclick=""self.location.href='default.asp?card=mainmenu&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:1px;'><img src='images/icons/48/Refresh.png' alt='Refresh' title='Refresh' style='border:none; cursor:pointer;' /></div>"
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		%>
		<div class='Font1' style='float:left; font-size:12pt;'>
		<%
		If GetSession("ed") = "Y" Then
			If Not GetSession("searchby") = "" Then %>
				<a class='Font2' style='float:left; font-size:12pt;' href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=<% =GetSession("SearchBy") %>">[Ändern]</a>
			<%
			Else %>
				<a class='Font2' style='float:left; font-size:12pt;' href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=NONE">[Ändern]</a>
			<%
			End If %>
		<% End If %>
		</div>
		<div class='Font1 RowElisp' style='float:right; font-size:12pt; max-width:60%;'>
			<% = GetSession("en") %>
		</div>
		<div style='clear:both;padding-bottom:10px;'>&nbsp;</div>

		<table cellpadding='2' cellspacing='0' style='width:100%;'>
			<tr align='center' valign='top'>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=myworkorders&amp;s=<% =SessionID %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/User Labor Male.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Meine Arbeitsauftr&#228;ge (<% =wocount %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=allworkorders&amp;s=<% =SessionID %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/User Group Labor.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Alle Arbeitsauftr&#228;ge (<% =wocount2 %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=allworkordersu&amp;s=<% =SessionID %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/User Labor Female.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Nicht zugeordnete Arbeitsauftr&#228;ge (<% =wocountu %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=assettasks&amp;s=<% =SessionID %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Hospital 2 Check.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Ausr&#252;stungs-Aufgabe
					</div>
				</td>
			</tr>
			<tr><td colspan='5'><div style='height:20px;'>&nbsp;</div></td></tr>
			<tr align='center' valign='top'>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=assets&amp;s=<% =SessionID %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Hospital 2.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Ausr&#252;stungs-Liste
					</div>
				</td>

				<%If WON Then%>
				<td style='width:25%; cursor:pointer;' <%If WON Then %>onclick="location.href = 'default.asp?card=wonew&amp;s=<% =SessionID %>';"<%End If%>>
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Configuration Tools.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Neuer Arbeitsauftrag
					</div>
				</td>
				<%End If%>
				<%If ASN Then%>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=asdetails&amp;s=<% =SessionID %>&amp;assetpk=-1';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Hospital 2 Add.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Neue Ausr&#252;stung
					</div>
				</td>
				<%End If%>
				<td style='width:25%; cursor:pointer;' <%If SYADJUSTIN Then%>onclick="location.href = 'default.asp?card=countinventory&amp;s=<% =SessionID %>';"<%End If%>>
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Coins.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Inventar z&#228;hlen
					</div>
				</td>
			</tr>
			<tr align='center' valign='top'>
				<%If False Then%>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=inventorymenu&amp;s=<% =SessionID %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Medical Invoice Flat.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Inventar Menu
					</div>
				</td>
				<%End If%>
				<td style='width:25%; cursor:pointer;'>
					&nbsp;
				</td>
				<td style='width:25%; cursor:pointer;'>
					&nbsp;
				</td>
				<td style='width:25%; cursor:pointer;'>
					&nbsp;
				</td>
			</tr>
			<tr><td colspan='5'><div style='height:20px;'>&nbsp;</div></td></tr>
		</table>

<%

		If WOFilter Then
		    If rccount > 1 or shcount > 1 Then
			%>
				<div style='padding:2px; background-color:#c7d4e1; border-bottom:2px solid #005288; font-size:12pt;' class='Font1'>Dashboard Kriterien</div>
			<div style='padding:5px;'>
			<%
		    End If
		    If rccount > 1 Then %>
		    	<div class='Font1' style='font-size:14pt; cursor:pointer;' onclick="OpenWindow('Repair Center', 600, 400, 'default.asp', '&card=rclookup&amp;s=<% =SessionID %>');">
					<div class='RowElisp' style='max-width:75%; float:left;'><% =GetSession("RCNM") %></div><div style='float:right;' class='Font2' style='font-size:12pt;'>[Ändern]</div><div style='clear:both;'></div>
		    	</div>
		    	<div style='height:10px;'>&nbsp;</div>
		    <%End If %>
		    <% If shcount > 1 Then %>
		    	<div class='Font1' style='font-size:14pt; cursor:pointer;' onclick="OpenWindow('Repair Center', 600, 400, 'default.asp', '&card=shlookup&amp;s=<% =SessionID %>');">
					<div class='RowElisp' style='max-width:75%; float:left;'>Shop: <% =GetSession("SHID") %></div><div style='float:right;' class='Font2' style='font-size:12pt;'>[Ändern]</div><div style='clear:both;'></div>
		    	</div>

			<%End If%>
		    	</div>
		    <%
		End If
		If False Then %><div><a class='Font1' style='font-size:12pt;' href="default.asp?card=entertime&amp;s=<% =SessionID %>">Geben Sie eine Zeit an</a></div><%
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
		Call OutputWAPError("Der Arbeitsauftrag konnte nicht gefunden werden")
	Else
		woid = NullCheck(rs("WOID"))
		reason = WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),500))
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

        parentlocationall = Replace(WAPValidate(NullCheck(RS("parentlocationall"))),"!#!",",&nbsp;")
        parentequipmentall = Replace(WAPValidate(NullCheck(RS("parentequipmentall"))),"!#!",",&nbsp;")
        If Not parentlocationall = "" Then
            parentlocationall = parentlocationall & ",&nbsp;"
        End If
        If Not parentequipmentall = "" Then
            parentequipmentall = parentequipmentall & ",&nbsp;"
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
				isopen = false
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
	%>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
		<%If Not IsBOF Then %>
			<div style='float:left;' onclick="location.href = 'default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;pr=<% =WOPK %>';">
				<img src='images/icons/48/Navigation 1 Left Black.png' alt='' title='' style='width:38px; height:38px;' />
			</div>
		<%End If %>
		<%If Not IsEOF Then %>
			<div style='float:left; padding-left:10px;' onclick="location.href = 'default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;nr=<% =WOPK %>';">
				<img src='images/icons/48/Navigation 1 Right Black.png' alt='' title='' style='width:38px; height:38px;' />
			</div>
		<%End If%>

		<div style='float:left; padding-left:14px; padding-top:7px; font-size:16pt;' class='Font1' onclick="location.href = 'default.asp?card=<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>&amp;s=<% =SessionID %>&amp;back=1';">
			AA Nr.<% =WOID %>
		</div>
		<div style='clear:both;'></div>
		</div>
		<% If Not AssetPK = "" Then %>
			<div class='Font1 RowElisp' onclick="location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';" style=' max-width:75%; padding:3px;cursor:pointer; font-size:12pt;padding-top:5px;'><% =ParentLocationAll %> <% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)</div>
		<% End If %>

		<div class='Font2 RowElisp' style='font-size:14pt;padding-top:10px; padding:3px; max-width:75%;'>
			<% =Reason %>
		</div>
		<% If Not RequestedLine = "" Then %>
			<div class='Font2 RowElisp' style='font-size:12pt;padding-top:10px; padding:3px; max-width:75%;'>
				<% =RequestedLine %>
			</div>
		<% End If %>

		<% If Not DepartmentLine = "" Then %>
			<div class='Font1 RowElisp' style='font-size:12pt;padding-top:10px; padding:3px; max-width:75%;'>
				<% =DepartmentLine %>
			</div>
		<% End If %>
		<div class='Font1' style='font-size:12pt;padding-top:10px; padding:3px;'>
			<% If Not IsOpen Then %>
				<div class='Font2'>
					<b>Geschlossen: </b><% =wostatusdate %>&nbsp;<% =wostatustime %>
				</div>
			<% Else %>
				<div style='clear:both;'>
					<div style='float:left;'>
						<div style='float:left;width:120px;'>Zieldatum:</div><div style='float:left; padding-left:5px;'><% =TargetDate %></div><div style='clear:both;'></div>
					</div>
					<div style='clear:both;'></div>
				</div>
				<div>
					<div style='float:left;'>
						<div style='float:left;width:120px;'><% =wostatusdesc %>: </div>
						<div style='float:left; padding-left:5px;'><% =wostatusdate %>&nbsp;<% = wostatustime %></div>
						<div style='clear:both;'></div>
					</div>
					<div style='clear:both;'></div>
				</div>

			<% End If%>
		</div>
		<div style='height:20px;'>&nbsp;</div>

		<div class='Font1' style='padding:5px;background-color:#f0f7fe; font-size:12pt; cursor:pointer;'>
			Details des Arbeitsauftrags
		</div>
		<table cellpadding='2' cellspacing='0' style='width:100%;'>
			<tr valign='top' align='center'>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=wodetails&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Medical Invoice 3D Edit.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Details bearbeiten
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Clipboard.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Aufgaben (<% =wocount %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=wolabor&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/User Group Home.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Arbeiter (<% =wocount2 %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=wopart&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Toolbox.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Materialien (<% =wocount3 %>)
					</div>
				</td>
				</tr>
				<tr valign='top' align='center'>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=womisccost&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Coins.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Weitere Kosten (<% =wocount4 %>)
					</div>
				</td>
				<% If AccessToAssign Then %>
				<td style='width:25%; cursor:pointer;' <% If AccessToAssign Then %>onclick="location.href = 'default.asp?card=woassign&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';"<%End If%>>
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/User Group Labor.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Zuweisungen (<% =wocount6 %>)
					</div>
				</td>
				<%End If%>
			</tr>
		</table>
		<div style='height:20px;'>&nbsp;</div>
		<div class='Font1' style='padding:5px;background-color:#f0f7fe; font-size:12pt; cursor:pointer;'>
			Arbeitsauftrag Vorg&#228;nge
		</div>
		<table cellpadding='2' cellspacing='0' style='width:100%;'>
			<tr valign='top' align='center'>
					<% If IsOpen AND UCase(wostatus) = "REQUESTED" AND AccessToIssue Then %>
					<td style='width:25%; cursor:pointer;' <% If IsOpen AND UCase(wostatus) = "REQUESTED" AND AccessToIssue Then %>onclick="location.href = 'default.asp?card=woissue&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';"<%End If%>>
						<div style="padding-bottom: 5px;">
							<img src='images/icons/48/Thumbs Up.png' alt='' title='' />
						</div>
						<div class='Font1'>
							Ausgabe
						</div>
					</td>
					<% End If %>
					<% If AccessToRespond Then %>
					<td style='width:25%; cursor:pointer;' <% If AccessToRespond Then %>onclick="location.href = 'default.asp?card=worespond&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';"<%End If%>>
						<div style="padding-bottom: 5px;">
							<img src='images/icons/48/Notepad.png' alt='' title='' />
						</div>
						<div class='Font1'>
							Antworten
						</div>
					</td>
					<% End If %>
					<% If Not UCase(wostatus) = "ONHOLD" Then %>
				<td style='width:25%; cursor:pointer;' <% If Not UCase(wostatus) = "ONHOLD" Then %>onclick="location.href = 'default.asp?card=woonhold&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';"<%End If%>>
						<div style="padding-bottom: 5px;">
							<img src='images/icons/48/Stop.png' alt='' title='' />
						</div>
					<div class='Font1'>
						Warteschleife
					</div>
				</td>
					<% End If %>

					<% If IsOpen Then %>
							<% If AccessToComplete Then %>
							<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=wocomplete&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';">
									<div style="padding-bottom: 5px;">
										<img src='images/icons/48/Status Flag Green.png' alt='' title='' />
									</div>
								<div class='Font1'>
									Abgeschlossen
								</div>
							</td>
								<% End If %>
								<% If AccessToClose Then %>
				</tr>
				<tr valign='top' align='center'>
							<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=woclose&amp;s=<% =SessionID %>&amp;wopk=<% =WOPK %>';">
									<div style="padding-bottom: 5px;">
										<img src='images/icons/48/Status Flag Red.png' alt='' title='' />
									</div>
								<div class='Font1'>
									Schlie&#223;en
								</div>
							</td>
							<% End If %>
				<% End If %>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=ashistory&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Hospital 2 Rating.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Ausr&#252;stungs-Verlauf (<% =wocount5 %>)
					</div>
				</td>
				</tr>
		</table>



		<%
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
		Call OutputWAPError("Keine Ausr&#252;stung gefunden.")
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
	%>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)
		</div>
		<table cellpadding='2' cellspacing='0' style='width:100%;'>
			<tr valign='top' align='center'>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=asdetails&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Medical Invoice 3D Edit.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Details bearbeiten
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=asmeters&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Configuration.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Z&#228;hler
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=asspecs&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Chart Dot.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Spezifizierungen (<% =ascount %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=aslabor&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/User Group Home.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Arbeiter / Kontakte (<% =ascount2 %>)
					</div>
				</td>
			</tr>
			<tr valign='top' align='center'>
			<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=ashistory&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/History.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Verlauf (<% =ascount5 %>)
					</div>
				</td>
				<td style='width:25%;'></td>
				<td style='width:25%;'></td>
				<td style='width:25%;'></td>
			</tr>
		</table>
		<% If False Then %>
		<div class='Font1' style='padding:5px;background-color:#f0f7fe; font-size:12pt; cursor:pointer;'>
			Ausr&#252;stungs-Vorgang
		</div>
		<table cellpadding='2' cellspacing='0' style='width:100%;'>
			<tr valign='top' align='center'>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=asparts&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Hammer.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Materialien (<% =ascount6 %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=ascontracts&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Contact Card.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Vertr&#228;ge (<% =ascount7 %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=aspms&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Medical Chart Xray Configur.png' alt='' title='' />
					</div>
					<div class='Font1'>
						PW (<% =ascount4 %>)
					</div>
				</td>
				<td style='width:25%; cursor:pointer;' onclick="location.href = 'default.asp?card=astasks&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';">
					<div style="padding-bottom: 5px;">
						<img src='images/icons/48/Notepad.png' alt='' title='' />
					</div>
					<div class='Font1'>
						Gefundene Aufgaben (<% =ascount3 %>)
					</div>
				</td>
			</tr>
		</table>
		<%End If%>
		<%
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
		    Call OutputWAPError("Keine Ausr&#252;stung gefunden.")
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
	Response.Write "<div style='padding:5px;'>"
		If rseof and Not AssetPK = "-1" Then
			Call OutputWAPMsg("The Asset was not found")
		Else %>
		<%If NOT AssetPK = "-1" Then%>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
				<% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)
			</div>
		<%Else%>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
				Neue Ausr&#252;stung
			</div>
		<%End If%>
			<%
			OutputFields
		End If

		If IsOpen and (ASE or AssetPK = "-1") Then
		ASDetailsSubmit rs
		End If

		Call CloseObj(rs)
		Set db = Nothing
		SetContext
	Response.Write "</div>"
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
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
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
						    HeaderMSG = "Übergeordnete ID konnte nicht gefunden werden."
						    ASDetails
						Else
						    ParentPK = rs2("AssetPK")
					    End If
					Else
					    HeaderMSG = "Übergeordnete ID erforderlich."
					    ASDetails
					End If
					Call SaveField(db,"ClassificationID","Classification ID","Classification","","LM","ASDetails",True)
                    'Call SaveField(db,"AssetID","Asset ID","","","C","ASDetails",True)
				    If Not NullCheck(Request("AssetID")) = "" Then
					    sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetID = '" & Request("AssetID") & "' "
					    Set rs2 = db.RunSQLReturnRS(sql,"")
					    Call CheckDB(db)
					    If Not rs2.eof Then
						    HeaderMSG = "Ausr&#252;stungs ID ist bereits einer anderen Ausr&#252;stung zugeordnet (" & WAPValidate(NullCheck(rs2("AssetName"))) & ")."
						    ASDetails
						Else
						    rs("AssetID") = Trim(Request("AssetID"))
					    End If
					Else
					    HeaderMSG = "Ausr&#252;stungs ID ist erforderlich."
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

	CardTitle = "Ausr&#252;stungsz&#228;hler"
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
		    Call OutputWAPError("Keine Ausr&#252;stung gefunden.")
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If rseof and Not AssetPK = "-1" Then
			Call OutputWAPMsg("The Asset was not found")
		Else %>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div style='float:left;'>
				<% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)
			</div>
			<div style='float:right;'>
				Z&#228;hlerstand
			</div>
			<div style='clear:both;'></div>
		</div>
		<div style='padding:8px;'>
			<%
			OutputFields
			%>
			</div>
			<%
		End If
		If IsOpen and (ASE or AssetPK = "-1") Then
			ASMetersSubmit rs
		End If
		Call CloseObj(rs)
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub ASMetersSubmit(rs)
	%>
	<div style='padding:15px;'>
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
	<input type="hidden" name="card" value="ASOptions"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="Assetpk" value="<% =AssetPK %>"/>
	<input type="hidden" name="POSTEDASMETERS" value="Y"/>
	</div>
	<%
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		%>
		<p align="center">
		<% If searchby = "" Then %>
		<div class='Font2' style='text-align:center;font-size:16pt;'>Suchen mit:</div>
		<% Else %>
		<div style='text-align:center;'><a class='Font2' style='font-size:16pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>">Suchen mit:</a></div>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=1">AA Nr.</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=2">Grund</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=3">Ausr&#252;stungs-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=4">Ausr&#252;stungs-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=5">Verfahrens-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=6">Verfahrens-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=7">Typ </a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=8">Priorit&#228;t</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=wosearch&amp;s=<% =SessionID %>&amp;wosearchby=9">Sub-Zustand</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=myworkorders&amp;s=<% =SessionID %>&amp;back=1">Meine Aas (<% =wocount %>)</a><br/>
		<% If TWC_WOAllOpen Then %>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=allworkorders&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle AAs (<% =wocount2 %>)</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=allworkordersu&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle nicht-zugewiesenen Arbeitsauftr&#228;ge (<% =wocountu %>)</a><br/>
		<% End If %>
		<% Case "1" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>AA&nbsp;#: </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>Grund: </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "3" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>Ausr&#252;stungs-ID: </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "4" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>Ausr&#252;stungs-Name: </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "5" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>Verfahrens-ID: </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "6" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>Verfahrens-Name: </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "7" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>Typ : </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "8" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>Priorit&#228;t: </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "9" %>
		<b class='Font1' style='font-size:14pt; font-weight:normal;'>Sub-Zustand: </b><input class='Textbox_Normal' size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If Not SearchBy = "" Then
			SearchSubmit
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CMSearch()

	Dim db, sql, rs

	CardCurrent = "CMSearch"
	CardTitle = "Suche nach Firma"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=cmsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=cmsearch&amp;s=<% =SessionID %>&amp;cmsearchby=1">Firmen-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=cmsearch&amp;s=<% =SessionID %>&amp;cmsearchby=2">Firmen-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=cmlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Firmen</a><br/>
		<% Case "1" %>
		<b>Firmen-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Firmen-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
			CardTitle = "Suche nach Problem"
			desc = "Problems"
			desc2 = "Problem"
		Case "FAFASEARCH"
			CardTitle = "Suche nach Fehler"
			desc = "Failures"
			desc2 = "Failure"
		Case "FASOSEARCH"
			CardTitle = "Suche nach L&#246;sung"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=<% =card %>&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=<% =card %>&amp;s=<% =SessionID %>&amp;<% =LCase(card) %>by=1"><% =desc2 %> ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=<% =card %>&amp;s=<% =SessionID %>&amp;<% =LCase(card) %>by=2"><% =desc2 %> Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=<% =Replace(UCase(card),"SEARCH","LOOKUP") %>&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle <% =desc %></a><br/>
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
	CardTitle = "Suche nach Ausr&#252;stung"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=assearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=assearch&amp;s=<% =SessionID %>&amp;assearchby=1">Ausr&#252;stungs-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=assearch&amp;s=<% =SessionID %>&amp;assearchby=2">Ausr&#252;stungs-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=aslookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Gesamte Ausr&#252;stung</a><br/>
		<% Case "1" %>
		<b>Ausr&#252;stungs-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Ausr&#252;stungs-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Ausr&#252;stung"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=assearch2&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=assearch2&amp;s=<% =SessionID %>&amp;assearchby=1">Ausr&#252;stungs-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=assearch2&amp;s=<% =SessionID %>&amp;assearchby=2">Ausr&#252;stungs-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=assets&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Gesamte Ausr&#252;stung</a><br/>
		<% Case "1" %>
		<b>Ausr&#252;stungs-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Ausr&#252;stungs-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		If Not SearchBy = "" Then
			SearchSubmit
		End If

		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CLSearch()

	Dim db, sql, rs

	CardCurrent = "CLSearch"
	CardTitle = "Suche nach Klassifizierung"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=clsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=clsearch&amp;s=<% =SessionID %>&amp;clsearchby=1">Klassifizierungs-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=clsearch&amp;s=<% =SessionID %>&amp;clsearchby=2">Klassifizierungs-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Klassifizierungen</a><br/>
		<% Case "1" %>
		<b>Klassifizierungs-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Klassifizierungs-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Kostenstelle"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=acsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=acsearch&amp;s=<% =SessionID %>&amp;acsearchby=1">Kostenstellen-IID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=acsearch&amp;s=<% =SessionID %>&amp;acsearchby=2">Kostenstellen-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=aclookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Kostenstellen</a><br/>
		<% Case "1" %>
		<b>Kostenstelle-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Kostenstelle-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Reparaturzentrum"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=rcsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=rcsearch&amp;s=<% =SessionID %>&amp;rcsearchby=1">Reparaturzentrum-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=rcsearch&amp;s=<% =SessionID %>&amp;rcsearchby=2">Reparaturzentrum-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=rclookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Reparaturzentren</a><br/>
		<% Case "1" %>
		<b>ID des Reparaturzentrums: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Name des Reparaturzentrums: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% End Select %>
		</p>
		<%
		OutputBackButton
		%>
		</td><td align="right">
		<%
		If Not SearchBy = "" Then
			SearchSubmit
		End If
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub SHSearch()

	Dim db, sql, rs

	CardCurrent = "SHSearch"
	CardTitle = "Suche nach Werkstatt"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=shsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=shsearch&amp;s=<% =SessionID %>&amp;shsearchby=1">Werkstatt-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=shsearch&amp;s=<% =SessionID %>&amp;shsearchby=2">Werkstatt-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=shlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Shops</a><br/>
		<% Case "1" %>
		<b>Werkstatt-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Werkstatt Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Kategorie"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=casearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=casearch&amp;s=<% =SessionID %>&amp;casearchby=1">Kategorien-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=casearch&amp;s=<% =SessionID %>&amp;casearchby=2">Kategorien-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=calookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Kategorien</a><br/>
		<% Case "1" %>
		<b>Kategorien-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Kategorien-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Zone"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=znsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=znsearch&amp;s=<% =SessionID %>&amp;znsearchby=1">Zonen-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=znsearch&amp;s=<% =SessionID %>&amp;znsearchby=2">Zonen-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=znlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Werkst&#228;tte</a><br/>
		<% Case "1" %>
		<b>Zonen-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Zonenname: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Verfahren"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=prsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=prsearch&amp;s=<% =SessionID %>&amp;prsearchby=1">Verfahrens-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=prsearch&amp;s=<% =SessionID %>&amp;prsearchby=2">Verfahrens-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=prlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Verfahren</a><br/>
		<% Case "1" %>
		<b>Verfahrens-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Verfahrens-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Abteilung"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=dpsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=dpsearch&amp;s=<% =SessionID %>&amp;dpsearchby=1">Abteilungs-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=dpsearch&amp;s=<% =SessionID %>&amp;dpsearchby=2">Abteilungs-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=dplookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Abteilungen</a><br/>
		<% Case "1" %>
		<b>Abteilungs-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Abteilungs-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Kunden"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=tnsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=tnsearch&amp;s=<% =SessionID %>&amp;tnsearchby=1">Kunden-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=tnsearch&amp;s=<% =SessionID %>&amp;tnsearchby=2">Kunden-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=tnlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Kunden</a><br/>
		<% Case "1" %>
		<b>Kunden-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Kunden-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Projekt"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=pjsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=pjsearch&amp;s=<% =SessionID %>&amp;pjsearchby=1">Projekt-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=pjsearch&amp;s=<% =SessionID %>&amp;pjsearchby=2">Projekt-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=pjlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Projekte</a><br/>
		<% Case "1" %>
		<b>Projekt-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Projektname: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Arbeiter"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=lasearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=lasearch&amp;s=<% =SessionID %>&amp;lasearchby=1">Arbeiter-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=lasearch&amp;s=<% =SessionID %>&amp;lasearchby=2">Arbeiter-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=lalookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Arbeiter</a><br/>
		<% Case "1" %>
		<b>Arbeiter-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Arbeiter-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Artikel"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=insearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=insearch&amp;s=<% =SessionID %>&amp;insearchby=1">Artikel-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=insearch&amp;s=<% =SessionID %>&amp;insearchby=2">Artikel-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=insearch&amp;s=<% =SessionID %>&amp;insearchby=3">Artikelschreibung</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=insearch&amp;s=<% =SessionID %>&amp;insearchby=4">Lieferanten-Artikel Nr.</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=inlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Artikel</a><br/>
		<% Case "1" %>
		<b>Artikel-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Artikel-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "3" %>
		<b>Artikel Beschreibung: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "4" %>
		<b>Lieferanten-Artikel Nr: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
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
	CardTitle = "Suche nach Standort"
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		%>
		<p align="center">
		<% If searchby = "" Then %>
		<b>Suchen mit:</b>
		<% Else %>
		<b><a href="default.asp?card=srsearch&amp;s=<% =SessionID %>">Suchen mit:</a></b>
		<% End If %>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=srsearch&amp;s=<% =SessionID %>&amp;srsearchby=1">Standort-ID</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=srsearch&amp;s=<% =SessionID %>&amp;srsearchby=2">Standort-Name</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=srlookup&amp;s=<% =SessionID %>&amp;searchby=NONE&amp;back=1">Alle Standorte</a><br/>
		<a class='Font1' style='font-size:14pt;' href="default.asp?card=srlookup&amp;s=<% =SessionID %>&amp;searchby=3">Alle Standorte (Alle Reparaturzentren)</a><br/>
		<% Case "1" %>
		<b>Standort-ID: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "2" %>
		<b>Standort-Name: </b><input size="15" value="" tabindex="1" type="text" name="searchvalue<% =r %>" format="*M"/>
		<% Case "3" %>
		<b>Klicken Sie Zulassen an um alle Lagerr&#228;ume an allen RCs zu sehen</b>
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
	%>
	<div style='padding:15px;'>
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
	<input type="hidden" name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="searchby" value="<% =searchby %>"/>
	<input type="hidden" name="back" value="1"/>
	</div>
	<%
End Sub

'====================================================================================================================================

Sub ASPhoto()

	Dim db, sql, rs, woid, assetid, assetname, reason, photo, FromWO

	CardTitle = Ausr&#252;stungsfoto"
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
		Call OutputWAPError("Ausr&#252;stungsfoto nicht gefunden")
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
		<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">AA Nr.<% =WOID %></a></b><br/>
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

	CardTitle = "Aufgaben des AA"
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
			Call OutputWAPMsg("Keine Aufgaben gefunden")
		Else %>
			<% If IsPocketIE or IsBlackBerry Then %>
			<p align="center" mode="wrap">
			<b><a href="default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=1">AA Nr.<% =WOID %></a></b><br/>
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
				<% =WAPValidate(Shorten(Replace(NullCheck(rs("TaskAction")),"%0D%0A",CHR(13) & CHR(10)),100)) %><% =LStyleEnd %>
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
					HeaderMSG = "Der Wert f&#252;r Tats&#228;chliche Stunden ist ung&#252;ltig."
					WOTask
				End If
				If Not NullCheck(Request("Measurement")) = "" Then
					rs("Measurement") =	NullCheck(Request("Measurement"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Ma&#223;e ist ung&#252;ltig."
					WOTask
				End If
				If Not NullCheck(Request("Rate")) = "" Then
					rs("Rate") = NullCheck(Request("Rate"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Tarif ist ung&#252;ltig."
					WOTask
				End If
				If Not NullCheck(Request("Comments")) = "" Then
					rs("Comments") =	NullCheck(Request("Comments"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Kommentare ist ung&#252;ltig."
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

	CardTitle = "Aufgabe des AA"
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
	'headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "';")
	headerString = GetBackButton("self.location.href = 'default.asp?card=wotasks&amp;s=" & SessionID & "&amp;wopk=" & WOPK & "';")

	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If rs.eof Then
			Call OutputWAPMsg("Aufgabe nicht gefunden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'><% =NullCheck(rs("TaskNo")) & ": " %><% =Shorten(Replace(NullCheck(rs("TaskAction")),"%0D%0A",CHR(13) & CHR(10)),500) %></div>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<%End If%>
			<div style='padding:8px;'>
			<%
			OutputFields
			%>
			</div>
			<%
		End If

		WOTaskSubmit rs
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WOTaskSubmit(rs)
	%>
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
	<input type="hidden" name="card" value="WOTasks"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wopk" value="<% =WOPK %>"/>
	<input type="hidden" name="pk" value="<% =PK %>"/>
	<%
End Sub

'====================================================================================================================================

Sub WOTasks()

	'Call ASPDebug
	'Response.End

	Dim db, sql, rs, woid, reason, assetid, assetname, photo
	Dim CheckBoxType, FromAsset

	CardTitle = "Aufgaben des AA"
	CardCurrent = "WOTasks"
	CardCurrentLevel = GetCardLevel()

	If Not Trim(UCase(GetSession("ParentCard2"))) = "ASSETTASKS" Then
		GetWOPK
		FromAsset = False
	Else
		assetpk = GetSession("ASSETPK2")
		FromAsset = True
	End If

	pagesize = GlobalPageSize

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
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and False Then
			Call OutputWAPMsg("Keine Aufgaben gefunden")
		Else
		%>
			<% If Not FromAsset Then %>
				<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
					<div style='float:left; padding-left:4px; padding-top:2px; font-size:16pt;' class='Font1' onclick="location.href = 'default.asp?card=<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>&amp;s=<% =SessionID %>&amp;back=1';">
						AA Nr.<% =WOID %>
					</div>
					<div style='float:right; padding-top:5px;'>Aufgaben des Arbeitsauftrags</div>
					<div style='clear:both;'></div>
				</div>

				<% If Not AssetPK = "" Then %>
					<div><a class="Font1" style='padding:5px;font-size:12pt;' href="default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>"><% =AssetName %> (<% =AssetID %>)</a></div>
				<% End If %>
				<div class='Font2' style='font-size:14pt;padding:5px;'><% =Reason%></div>
		    <% End If%>
			<% If Not FromAsset Then %>
			<%
			End If %>
			<div style='padding:5px;'>
			<%

			Dim borderzero, wopkc
			wopkc="!"
			borderzero = ""

			Dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1

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
						AA Nr.<% =wopkc %><br/><%
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
				    <% =WAPValidate(Shorten(Replace(NullCheck(rs("TaskAction")),"%0D%0A",CHR(13) & CHR(10)),100)) %><%
				Else

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
								TaskText3 = " (Unter Minimalwert von " & rs("ValueLow") & ")"
								LineTemplate2 = 11
							Else
								SpecLowOK = True
								SpecText = SpecText & "Minimaler Wert: " & rs("ValueLow") & " "
							End If
						End If
					End If
					If rs("Spec") and Not NullCheck(rs("ValueHi")) = "" and Not NullCheck(rs("Measurement")) = "" Then
						If IsNumeric(rs("Measurement")) and IsNumeric(rs("ValueHi")) Then
							If CLng(rs("Measurement")) > CLng(rs("ValueHi")) Then
								LineTemplate2 = 11
								TaskText3 = " (Über Maximalwert von " & rs("ValueHi") & ")"
							Else
								SpecHiOK = True
								If SpecText = "" Then
									SpecText = SpecText & "Maximaler Wert: " & rs("ValueLow")
								Else
									SpecText = "Umfang: " & rs("ValueLow") & " - " & rs("ValueHi")
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
					'If RS("Complete") Then
					If CheckBoxType = "GRAPHIC" Then
					If BitNullCheck(rs("Complete")) Then %>
						<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>"></a>
						<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>"><% End If
					ElseIf CheckBoxType = "TEXT" Then
						If BitNullCheck(rs("Complete")) Then %>
							<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>"><% Else %>
							<a href="default.asp?card=wotasks&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>&amp;pagepos=<% =pagepos %>&amp;flipit=<% =RandomString(3) %>"><% End If
						Else%>
						<%If RS("Complete") Then%>
							<div class='Font1' style='cursor:pointer; background-color:<%=altStyle%>;'>
								<div onclick="location.href = 'default.asp?card=wotaskuncomplete&s=<% =SessionID %>&wopk=<%=wopk%>&action=wotaskuncomplete&taskid=<%=rs("PK")%>';" style='float:left;'>
									<img id="img<%=rs("PK")%>_1" src='images/icons/48/Status Flag Green.png' alt='' title='' />
								</div>
								<div class='Font1' style='float:left; font-size:14pt; padding-left:15px; padding-top:10px;' onclick="location.href = 'default.asp?card=wotask&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>';">
									<% =rs("TaskNo") %>: <% =WAPValidate(Shorten(Replace(NullCheck(TaskTextFinal),"%0D%0A","&nbsp;"),100)) %>
								</div>
								<div style='clear:both;'></div>
							</div>
						<%Else%>
							<div class='Font1' style='cursor:pointer; background-color:<%=altStyle%>;'>
								<div onclick="location.href = 'default.asp?card=wotaskcomplete&s=<% =SessionID %>&wopk=<%=wopk%>&action=wotaskcomplete&taskid=<%=rs("PK")%>';" style='float:left;'>
									<img id="img<%=rs("PK")%>" src='images/icons/48/Status Flag Black.png' alt='' title='' />
								</div>
								<div class='Font1' style='float:left; font-size:14pt; padding-left:15px; padding-top:10px;' onclick="location.href = 'default.asp?card=wotask&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;pk=<% =rs("PK") %>';">
									<% =rs("TaskNo") %>: <% =WAPValidate(Shorten(Replace(NullCheck(TaskTextFinal),"%0D%0A","&nbsp;"),100)) %>
								</div>
								<div style='clear:both;'></div>
							</div>
						<%End If%>
				<%
						End If
					End if
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
			</div><%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOLaborRec()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec

    newcontextpage = True

	CardTitle = "AA Arbeiter"
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
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;back=2';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("Arbeiter nicht gefunden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<% If PK = "-1" Then %>
				<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>Neuer Arbeiter</div>
			<%Else%>
				<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'><% =WAPValidate(NullCheck(rs("LaborName"))) %>: [<% =NullCheck(RS("EstimatedHours")) & " Est] [" & NullCheck(RS("TotalHours")) & "Tats&#228;chlich" & "] " & DateNullCheck(RS("WorkDate")) %></div>
			<%End If%>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>

			<% End If %>
			<div style='padding:8px;'>
			<%
			OutputFields
			%>
			</div>
			<%
		End If

		If IsOpen Then
			WOLaborRecSubmit rs
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
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
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
						HeaderMSG = "Arbeiter-ID konnte nicht gefunden werden."
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
						HeaderMSG = "Der Wert f&#252;r Regul&#228;re Stunden ist ung&#252;ltig."
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
						HeaderMSG = "Der Wert f&#252;r Überstunden ist ung&#252;ltig."
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
						HeaderMSG = "Der Wert f&#252;r Weitere Stunden ist ung&#252;ltig."
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
						HeaderMSG = "Der Wert f&#252;r Arbeitsdatum ist ung&#252;ltig."
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
						HeaderMSG = "Der Wert f&#252;r Zeit ist ung&#252;ltig."
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
						HeaderMSG = "Der Wert f&#252;r Zeitlimit ist ung&#252;ltig."
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
						HeaderMSG = "Keine Kostenstelle gefunden."
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
						HeaderMSG = "Die angegebene Kostenstelle-ID ist ung&#252;ltig."
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
						HeaderMSG = "Kategorien-ID konnte nicht gefunden werden."
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
						HeaderMSG = "Der Wert f&#252;r Kategorie-ID ist ung&#252;ltig."
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
				    HeaderMSG = "Der Wert f&#252;r Kommentare ist ung&#252;ltig."
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

	CardTitle = "AA Arbeiter"
	CardCurrent = "WOLabor"
	CardCurrentLevel = GetCardLevel()

	GetWOPK
	pagesize = GlobalPageSize

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

	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;back=2';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=WOLaborRec&s=" & SessionID & "&wopk=" & wopk & "&pk=-1';"" style='float:left; padding-right:10px; padding-top:6px;'><img src='images/icons/48/add.png' alt='New' title='New' style='border:none; cursor:pointer; width:42px; height:42px;' /></div>"
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If rs.eof and False Then
			Call OutputWAPMsg("Kein Arbeiter gefunden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>AA Arbeiter</div>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<%End If%>

			<p mode="nowrap">
			<%
			dim iCount, altStyle, returnURL
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1

			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnURL = "self.location.href = 'default.asp?card=wolaborrec&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;pk=" & rs("PK") & "';"

				Dim AssignmentsCompleted
				If RS("Completed") Then
					AssignmentsCompleted = "<img src=""images/taskchecked.gif"">"
				Else
					AssignmentsCompleted = "<img src=""images/taskline.gif"">"
				End If
				If (Not AccessToLARates And RS("ModuleID") = "LA" And Not Trim(RS("LaborPK")) = Trim(GetSession("UserPK"))) or (Not AccessToCRRates And RS("ModuleID") = "CR" And Not Trim(RS("LaborPK")) = Trim(GetSession("CraftPK"))) Then
				Else
				'DateNullCheck(RS("WorkDate"))
				%>

				<div class='Font1' style='clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
					<div style='float:left;cursor:pointer;' onclick="<%=returnurl%>">
						<% =WAPValidate(NullCheck(rs("LaborName"))) %>
					</div>
					<div style=' font-size:14pt;float:right;cursor:pointer;' onclick="<%=returnurl%>"><% =NullCheck(RS("EstimatedHours")) %>&nbsp;Gesch&#228;tzte Stunde(n)&nbsp;&nbsp;-&nbsp;&nbsp;<% =NullCheck(RS("TotalHours")) %>&nbsp;Gesamtstunde(n)</div>
					<div style='clear:both;'></div>
				</div>

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
		%><div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOPartRec()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec

    newcontextpage = True

	CardTitle = "AA Material"
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
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;back=2';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
	%>
		<script type='text/javascript'>
			$(document).ready(function(){
				try{
					$('#PartID').focus();
				}
				catch(e){}
			});
		</script>
	<%
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("Artikel konnte nicht gefunden werden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<% If PK = "-1" Then %>
				<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>Neuer Artikel</div>
			<%Else%>
				<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>[<% =WAPValidate(NullCheck(rs("PartID"))) %>] [<% =WAPValidate(NullCheck(rs("PartName"))) %>] [<% =WAPValidate(NullCheck(RS("LocationID"))) %>] [<% =WAPValidate(NullCheck(RS("QuantityEstimated"))) %>&nbsp;Gesch&#228;tzt] [<% =WAPValidate(NullCheck(RS("QuantityActual"))) %>&nbsp;Tats&#228;chlich]</div>
			<%End If%>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<%End If%>
			<div style='padding:8px;'>
			<%
			OutputFields
			%>
			</div>
			<%
		End If

		If IsOpen Then
			WOPartRecSubmit rs
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
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
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
						HeaderMSG = "Artikel-ID konnte nicht gefunden werden."
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
						HeaderMSG = "Standort-ID konnte nicht gefunden werden."
						WOPartRec
					Else
						rs("LocationPK") = rs2("LocationPK")
						rs("LocationID") = rs2("LocationID")
						rs("LocationName") = rs2("LocationName")
						LocationPK = rs2("LocationPK")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Standort-ID ist ung&#252;ltig."
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
						HeaderMSG = "Der Wert f&#252;r Tats&#228;chliche Menge ist ung&#252;ltig."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("OtherCost")) = "" Then
					rs("OtherCost") =	NullCheck(Request("OtherCost"))
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Weitere Kosten ist ung&#252;ltig."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("AccountID")) = "" Then
					sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Keine Kostenstelle gefunden."
						WOPartRec
					Else
						rs("AccountPK") = rs2("AccountPK")
						rs("AccountID") = rs2("AccountID")
						rs("AccountName") = rs2("AccountName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Die angegebene Kostenstelle-ID ist ung&#252;ltig."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("CategoryID")) = "" Then
					sql = "SELECT * FROM Category WITH (NOLOCK) WHERE CategoryID = '" & NullCheck(Request("CategoryID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Kategorien-ID konnte nicht gefunden werden."
						WOPartRec
					Else
						rs("CategoryPK") = rs2("CategoryPK")
						rs("CategoryID") = rs2("CategoryID")
						rs("CategoryName") = rs2("CategoryName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Kategorie-ID ist ung&#252;ltig."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("Serial")) = "" Then
					rs("Serial") =	NullCheck(Request("Serial"))
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Seriennummer ist ung&#252;ltig."
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
						HeaderMSG = "Der Wert f&#252;r vorhandene Seriennummer ist ung&#252;ltig."
						WOPartRec
					End If
				End If
				If Not NullCheck(Request("SerialReplaceToLocationID")) = "" Then
					sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetID = '" & NullCheck(Request("SerialReplaceToLocationID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Vorhandenes zum Standort bewegen' konnte nicht gefunden werden."
						WOPartRec
					Else
						rs("SerialReplaceToLocationID") = NullCheck(Request("SerialReplaceToLocationID"))
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Vorhandenes zum Standort Verschieben ist ung&#252;ltig."
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

	CardTitle = "AA Materialien"
	CardCurrent = "WOPart"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = LookupSize
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = LookupSize
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
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;back=2';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=WOPartRec&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;pk=-1';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Add.png' alt='Hinzuf&#252;gen' title='Hinzuf&#252;gen' style='border:none; cursor:pointer;' /></div>"
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and False Then
			Call OutputWAPMsg("No Materials were Found")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>Materialien</div>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<%End If%>

			<div>
			<%
			dim iCount, altStyle, returnURL
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1
			returnURL = "self.location.href = 'default.asp?card=wopartrec&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;pk=" & rs("PK") & "';"

			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else %>

			<div class='Font1' style='clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
				<div style='float:left;cursor:pointer;' onclick="<%=returnurl%>">
					<div><% =WAPValidate(NullCheck(rs("PartID"))) %>: <% =WAPValidate(NullCheck(rs("PartName"))) %></div>
					<div class='Font2' style='font-size:12pt;'><% =WAPValidate(NullCheck(RS("LocationID"))) %></div>
				</div>
				<div style=' font-size:14pt;float:right;cursor:pointer;' onclick="<%=returnurl%>"><% =WAPValidate(NullCheck(RS("QuantityEstimated"))) %>&nbsp;Gesch&#228;tzt&nbsp;&nbsp;|&nbsp;&nbsp;<% =WAPValidate(NullCheck(RS("QuantityActual"))) %>&nbsp;Tats&#228;chlich</div>
				<div style='clear:both;'></div>
			</div>

				<%
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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOMiscCostRec()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec

    newcontextpage = True

	CardTitle = "Weitere Kosten des AA"
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
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;back=2';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("Weitere Kosten konnte nicht gefunden werden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<% If PK = "-1" Then %>
				<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>Neue Weitere Kosten</div>
			<%Else%>
				<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>[<% =WAPValidate(NullCheck(rs("MiscCostName"))) %>] [<% =DateNullCheck(RS("MiscCostDate")) %>] [<% =FormatNumber(WAPValidate(NullCheck(RS("EstimatedCost"))),2,-2,0,0) %>&nbsp;Gesch&#228;tzt] [<% =FormatNumber(WAPValidate(NullCheck(RS("ActualCost"))),2,-2,0,0) %>&nbsp;Tats&#228;chlich]</div>
			<%End If%>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<%End If%>

			<div style='padding:8px;'>
			<%
			OutputFields
			%>
			</div>
			<%
		End If

		If IsOpen Then
			WOMiscCostRecSubmit rs
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
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
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
						HeaderMSG = "Der Wert f&#252;r Name ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("MiscCostDesc")) = "" Then
					rs("MiscCostDesc") = NullCheck(Request("MiscCostDesc"))
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Beschreibung ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("CompanyID")) = "" Then
					sql = "SELECT * FROM Company WITH (NOLOCK) WHERE CompanyID = '" & NullCheck(Request("CompanyID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Firmen-ID konnte nicht gefunden werden."
						WOMiscCostRec
					Else
						rs("CompanyPK") = rs2("CompanyPK")
						rs("CompanyID") = rs2("CompanyID")
						rs("CompanyName") = rs2("CompanyName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Firmen-ID ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("LaborID")) = "" Then
					sql = "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborID = '" & NullCheck(Request("LaborID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Arbeiter-ID konnte nicht gefunden werden."
						WOMiscCostRec
					Else
						rs("LaborPK") = rs2("LaborPK")
						rs("LaborID") = rs2("LaborID")
						rs("LaborName") = rs2("LaborName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Arbeiter-ID ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("InvoiceNumber")) = "" Then
					rs("InvoiceNumber") = NullCheck(Request("InvoiceNumber"))
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Rechnungsnummer ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("MiscCostDate")) = "" Then
					rs("MiscCostDate") = SQLdatetimeADO(Request("MiscCostDate"))
					If Err.Number <> 0 Then
						HeaderMSG = "Das angegebene Datum ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("AccountID")) = "" Then
					sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Keine Kostenstelle gefunden."
						WOMiscCostRec
					Else
						rs("AccountPK") = rs2("AccountPK")
						rs("AccountID") = rs2("AccountID")
						rs("AccountName") = rs2("AccountName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Die angegebene Kostenstelle-ID ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("CategoryID")) = "" Then
					sql = "SELECT * FROM Category WITH (NOLOCK) WHERE CategoryID = '" & NullCheck(Request("CategoryID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Kategorien-ID konnte nicht gefunden werden."
						WOMiscCostRec
					Else
						rs("CategoryPK") = rs2("CategoryPK")
						rs("CategoryID") = rs2("CategoryID")
						rs("CategoryName") = rs2("CategoryName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Kategorie-ID ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("ActualCost")) = "" Then
					rs("ActualCost") = NullCheck(Request("ActualCost"))
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Tats&#228;chliche Kosten ist ung&#252;ltig."
						WOMiscCostRec
					End If
				End If
				If Not NullCheck(Request("Comments")) = "" Then
					rs("Comments") =	NullCheck(Request("Comments"))
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Kommentare ist ung&#252;ltig."
					WOMiscCostRec
				End If
				On Error Goto 0
			End If
			db.dobatchupdate rs
			If db.dok Then
				'pl = rs("pk")
			Else
				If InStr(db.derror,"NULL") > 0 Then
					HeaderMSG = "Muss mit einem Wert f&#252;r Namen versehen werden."
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

	CardTitle = "Weitere Kosten des AA"
	CardCurrent = "WOMiscCost"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = LookupSize
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = LookupSize
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
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "';")
	headerString = headerString & GetAddButton("self.location.href = 'default.asp?card=WOMiscCostRec&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&pk=-1';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and False Then
			Call OutputWAPMsg("Keine weiteren Kosten gefunden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>Weitere Kosten</div>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<%End If%>

			<div>
			<%
			dim iCount, altStyle, returnURL
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1
			returnURL = "self.location.href = 'default.asp?card=womisccostrec&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;pk=" & rs("PK") & "';"

			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else %>

			<div class='Font1' style='clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
				<div style='float:left;cursor:pointer;' onclick="<%=returnurl%>">
					<div><% =WAPValidate(NullCheck(rs("MiscCostName"))) %></div>
					<div class='Font2' style='font-size:12pt;'><% =DateNullCheck(RS("MiscCostDate")) %></div>
				</div>
				<div style=' font-size:14pt;float:right;cursor:pointer;' onclick="<%=returnurl%>"><% =FormatNumber(WAPValidate(NullCheck(RS("EstimatedCost"))),2,-2,0,0) %>&nbsp;Gesch&#228;tzt&nbsp;&nbsp;-&nbsp;&nbsp;<% =FormatNumber(WAPValidate(NullCheck(RS("ActualCost"))),2,-2,0,0) %>&nbsp;Tats&#228;chlich</div>
				<div style='clear:both;'></div>
			</div>

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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>


		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASSpecsRec()

	Dim db, sql, rs, ASE, assetid, assetname, photo, parentlocationall, parentequipmentall, newrec, IsLocation, rseof

    newcontextpage = True

	CardTitle = "Ausr&#252;stungsspezifizierung"
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

	sql = "SELECT TOP 1 SpecificationName FROM AssetSpecification WITH ( NOLOCK ) WHERE PK = " & pk

	Set rs = db.RunSQLReturnRS(sql,"")
	Call CheckDB(db)
	Dim SpecificationName

	If Not rs.eof Then
		SpecificationName = NullCheck(rs("SpecificationName"))
	End If
	Set rs = Nothing

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
		    Call OutputWAPError("Keine Ausr&#252;stung gefunden.")
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If rseof and Not AssetPK = "-1" Then
			Call OutputWAPMsg("Ausr&#252;stungsspezifizierung nicht gefunden")
		Else %>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div style='float:left;'>
				<% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)
			</div>
			<div style='float:right;'>
				Spezifizierung: <%=SpecificationName%>
			</div>
			<div style='clear:both;'></div>
		</div>
			<div style='padding:8px;'>
			<%
			OutputFields
			%>
			</div>
			<%
		End If
		ASSpecsSubmit rs
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
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
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

	CardTitle = "Ausr&#252;stungsspezifizierungen"
	CardCurrent = "ASSpecs"
	CardCurrentLevel = GetCardLevel()

	GetAssetPK

	pagesize = GlobalPageSize

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
	    Call OutputWAPError("Keine Ausr&#252;stung gefunden.")
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and False Then
			Call OutputWAPMsg("Keine Spezifikationen gefunden")
		Else %>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div style='float:left;'>
				<% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)
			</div>
			<div style='float:right;'>
				Spezifizierung
			</div>
			<div style='clear:both;'></div>
		</div>

		<div>
			<%
			Dim iCount, altStyle, returnURL
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1

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
			        ValueCombined = "(" & ValueCombined & ")"
			    End If
				returnURL = "self.location.href = 'default.asp?card=asspecsrec&amp;s=" & SessionID & "&amp;assetpk=" & assetpk & "&amp;pk=" & rs("PK") & "';"
			    %>

				<div class='Font1' style='clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
					<div style='float:left;cursor:pointer;' onclick="<%=returnurl%>">
						<% =WAPValidate(NullCheck(rs("SpecificationName"))) %>
					</div>
					<div style=' font-size:14pt;float:right;cursor:pointer;' onclick="<%=returnurl%>"><% =ValueCombined %></div>
					<div style='clear:both;'></div>
				</div>

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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&assetpk=<%=assetpk%>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&assetpk=<%=assetpk%>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASLaborRec()

	Dim db, sql, rs, ASE, assetid, assetname, photo, parentlocationall, parentequipmentall, newrec, IsLocation, rseof

    newcontextpage = True

	CardTitle = "Ausr&#252;stung Arbeiter / Kontakte"
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
		    Call OutputWAPError("Keine Ausr&#252;stung gefunden.")
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rseof and Not AssetPK = "-1" Then
			Call OutputWAPMsg("Ausr&#252;stungs-Arbeiter / Kontakt konnte nicht gefunden werden")
		Else %>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div style='float:left;'>
				<% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)
			</div>
			<div style='float:right;'>Arbeiter-Daten</div>
			<div style='clear:both;'></div>
		</div>
		<div style='padding:8px;'>
			<%
			OutputFields
			%>
			</div>
			<%
		End If
		If False and IsOpen and (ASE or AssetPK = "-1") Then
			ASLaborSubmit rs
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
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
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

	CardTitle = "Ausr&#252;stung Arbeiter / Kontakte"
	CardCurrent = "ASLabor"
	CardCurrentLevel = GetCardLevel()

	GetAssetPK

    If IsPocketIE Then
        If lang = "HTML" Then
	        pagesize = LookupSize
        Else
	        pagesize = 7
	    End If
    ElseIf IsBlackBerry Then
        pagesize = LookupSize
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
	    Call OutputWAPError("Keine Ausr&#252;stung gefunden.")
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
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If rs.eof and False Then
			Call OutputWAPMsg("Keine Spezifikationen gefunden")
		Else %>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div style='float:left;'>
				<% =ParentLocationAll %><% =ParentEquipmentAll %><% =AssetName %> (<% =AssetID %>)
			</div>
			<div style='float:right;'>
				Arbeiter / Kontakte
			</div>
			<div style='clear:both;'></div>
		</div>

			<div>
			<%
			dim iCount, altStyle, returnURL
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1
			returnURL = "self.location.href = 'default.asp?card=aslaborrec&amp;s=" & SessionID & "&amp;assetpk=" & assetpk & "&amp;pk=" & rs("PK") & "';"

			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else %>


			<div class='Font1' style='clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
				<div style='float:left;cursor:pointer;' onclick="<%=returnurl%>">
					<% =WAPValidate(NullCheck(rs("LaborName"))) %>
				</div>
				<div style=' font-size:14pt;float:right;cursor:pointer;' onclick="<%=returnurl%>">Arbeit:&nbsp;<% =WAPValidate(NullCheck(rs("PhoneWork"))) %>&nbsp;&nbsp;-&nbsp;&nbsp;Mobil:&nbsp;<% =NullCheck(RS("PhoneMobile")) %></div>
				<div style='clear:both;'></div>
			</div>

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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&assetpk=<%=assetpk%>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&assetpk=<%=assetpk%>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOAssignRec()

	Dim db, sql, rs, woid, reason, assetid, assetname, photo, newrec, wostatus

    newcontextpage = True

	CardTitle = "AA Zuweisen"
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
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;back=2';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("Zuweisung nicht gefunden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<% If PK = "-1" Then %>
				<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>Neue Auftragszuweisung</div>
			<%Else%>
				<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>[<% =WAPValidate(NullCheck(rs("LaborName"))) %>] [<% =DateNullCheck(RS("AssignedDate")) %>] [<% =NullCheck(RS("AssignedHours")) %>&nbsp;Stunde(n)]</div>
			<%End If%>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<%End If%>

			<div style='padding:8px;'>
			<%
			OutputFields
			%>
			</div>
			<%
		End If

		If IsOpen Then
			WOAssignRecSubmit rs
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
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
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
						HeaderMSG = "Arbeiter-ID konnte nicht gefunden werden."
						WOAssignRec
					Else
						rs("LaborPK") = rs2("LaborPK")
						rs("LaborID") = rs2("LaborID")
						rs("LaborName") = rs2("LaborName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Arbeiter-ID ist ung&#252;ltig."
						WOAssignRec
					End If
				End If
				If Not NullCheck(Request("AssignedDate")) = "" Then
					rs("AssignedDate") = SQLdatetimeADO(Request("AssignedDate"))
					If Err.Number <> 0 Then
						HeaderMSG = "Das zugeordnete Datum ist ung&#252;ltig."
						WOAssignRec
					End If
				Else
					HeaderMSG = "Muss mit einem Wert f&#252;r die Vergabe von Datum versehen werden."
					WOAssignRec
				End If
				If Not NullCheck(Request("AssignedHours")) = "" Then
				    rs("AssignedHours") = Request("AssignedHours")
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r zugeordnete Stunden ist ung&#252;ltig."
						WOAssignRec
					End If
				Else
					HeaderMSG = "Muss mit einem Wert f&#252;r die Vergabe von Stunden versehen werden."
					WOAssignRec
				End If
				If Not NullCheck(Request("AssignedLead")) = "" Then
				    If Request("AssignedLead") = "Y" or Request("AssignedLead") = "2" Then
				        rs("AssignedLead") = True
				    Else
				        rs("AssignedLead") = False
				    End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r die zugeordnete F&#252;hrung ist ung&#252;ltig."
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
						HeaderMSG = "Der Wert f&#252;r die zugeordnete PDA ist ung&#252;ltig."
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
					HeaderMSG = "Muss mit einem Wert f&#252;r Arbeiter ID versehen werden."
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

	CardTitle = "AA Zuweisungen"
	CardCurrent = "WOAssign"
	CardCurrentLevel = GetCardLevel()

	GetWOPK

	pagesize = GlobalPageSize
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
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;back=2';")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=WOAssignRec&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;pk=-1';"" style='float:left; padding-right:10px; padding-top:6px;'><img src='images/icons/48/add.png' alt='Hinzuf&#252;gen' title='Hinzuf&#252;gen' style='border:none; cursor:pointer; width:42px; height:42px;' /></div>"
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and False Then
			Call OutputWAPMsg("Keine Zuordnungen gefunden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<div class='Font1' style='float:right; padding-top:3px; font-size:14pt;'>AA Zuweisungen</div>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<%End If%>

			<p mode="nowrap">
			<%

			dim altStyle, iCount,returnURL
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1

			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnURL = "self.location.href = 'default.asp?card=woassignrec&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;pk=" & rs("PK") & "';"
			%>

			<div class='Font1' style='clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
				<div style='float:left;cursor:pointer;' onclick="<%=returnurl%>">
					<% =WAPValidate(NullCheck(rs("LaborName"))) %>
				</div>
				<div style=' font-size:14pt;float:right;cursor:pointer;' onclick="<%=returnurl%>"><% =DateNullCheck(RS("AssignedDate")) %>&nbsp;&nbsp;-&nbsp;&nbsp;<% =NullCheck(RS("AssignedHours")) %>&nbsp;Stunde(n)]</div>
				<div style='clear:both;'></div>
			</div>



			<%
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
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&wopk=<%=wopk%>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub WOIssue()
	Dim HeaderTitle
	CardTitle = "AA ausstellen"
	HeaderTitle = "AA ausstellen"
	CardCurrent = "WOIssue"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WOOnHold()
	Dim HeaderTitle
	CardTitle = WAPValidate("AA Warteschleife")
	HeaderTitle = WAPValidate("AA in die Warteschleife setzen")
	CardCurrent = "WOOnHold"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WORespond()
	Dim HeaderTitle
	CardTitle = "AA beantworten"
	HeaderTitle = "AA beantworten"
	CardCurrent = "WORespond"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WOComplete()
	Dim HeaderTitle
	CardTitle = "AA fertiggestellt"
	HeaderTitle = "AA fertigstellen"
	CardCurrent = "WOComplete"
	Call WOStatusProcess(HeaderTitle)
End Sub

Sub WOClose()
	Dim HeaderTitle
	CardTitle = "WO Close"
	HeaderTitle = "AA schlie&#223;en"
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
		Call OutputWAPMsg("Der Arbeitsauftrag konnte nicht gefunden werden")
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
	If UCase(card) = "WODETAILS" Then

		'Call BuildFields("Reason","Reason","C",GlobalFieldLength,"*M","true",NullCheck(rs("Reason")),rs,"-1")
		Call BuildFields("AssetID","Asset ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("AssetID")),rs,"-1")
		'Call BuildFields("ProblemID","Problem ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("ProblemID")),rs,"-1")
		'Call BuildFields("ProcedureID","Procedure ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("ProcedureID")),rs,"-1")
		'Call BuildFields("TargetDate","Target Date","C",GlobalFieldLength,"*M","true",CStr(DateNullCheck(rs("TargetDate"))),rs,"-1")
		'Call BuildFields("TargetHours","Target Hours","C",GlobalFieldLength,"*M","true",NullCheck(rs("TargetHours")),rs,"-1")
		Call BuildFields("Priority","Priority","C",GlobalFieldLength,"*M","true",NullCheck(rs("Priority")),rs,"-1")
		Call BuildFields("Type","Type","C",GlobalFieldLength,"*M","true",NullCheck(rs("Type")),rs,"-1")
		Call BuildFields("ShopID","Shop ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("ShopID")),rs,"-1")
		'Call BuildFields("AccountID","Account ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("AccountID")),rs,"-1")
		'Call BuildFields("CategoryID","Category ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("CategoryID")),rs,"-1")
		Call BuildFields("DepartmentID","Department ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("DepartmentID")),rs,"-1")
		Call BuildFields("TenantID","Customer ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("TenantID")),rs,"-1")
		Call BuildFields("ProjectID","Project ID","C",GlobalFieldLength,"*M","true",NullCheck(rs("ProjectID")),rs,"-1")

	Else
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
	End If

	Call StartMobileDocument(CardTitle)
	headerString = GetBackButton("self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & wopk & "&amp;back=2';")
	'headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
		If rs.eof and Not PK = "-1" Then
			Call OutputWAPMsg("AA nicht gefunden")
		Else %>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div class='Font1' style='float:left;cursor:pointer;padding-left:4px; padding-top:3px; font-size:16pt;' onclick="self.location.href='default.asp?card=wooptions&amp;s=<% =SessionID %>&amp;wopk=<% =wopk %>&amp;back=2';">AA Nr.<% =WOID %></div>
			<div class='Font1' style='float:right;font-size:14pt; padding-top:3px;'><% =HeaderTitle %></div>
			<div style='clear:both;'></div>
			</div>
			<% If Not AssetPK = "" Then %>
			<div class='Font1' style='font-size:12pt; cursor:pointer; padding:2px; padding-left:6px;' onclick="self.location.href = 'default.asp?card=asoptions&amp;s=<% =SessionID %>&amp;assetpk=<% =AssetPK %>';"><% =AssetName %> (<% =AssetID %>)</div>

			<div class='Font2' style='padding:2px; padding-left:6px; font-size:14pt;'><% =Reason %></div>
			<div style='padding:8px;'>
			<%
			End If
			OutputFields
			%>
			</div>
			<%
		End If

		If IsOpen Then
			WOCloseSubmit rs
		End If
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WOCloseSubmit(rs)
	%>
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
	<% If UCase(card) = "WOONHOLD" or UCase(card) = "WORESPOND" or UCase(card) = "WOISSUE" or UCase(card) = "WODETAILS" Then %>
	<input type="hidden" name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
	<% Else %>
	<input type="hidden" name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-2)) %>"/>
	<% End If %>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wopk" value="<% =WOPK %>"/>
	<input type="hidden" name="woaction" value="<% =card %>"/>
	<%
End Sub

'====================================================================================================================================

Sub AssetMenu()

	Dim db, sql, rs

	CardTitle = "Ausr&#252;stungsmen&#252;"
	CardCurrent = "Asset Menu"
	CardCurrentLevel = GetCardLevel()

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call StartMobileDocument(CardTitle)
	headerString = GetBackButton("history.back();")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
	Call OutputWAPMsg("Derzeit sind keine Artikel im Ausr&#252;stungsmen&#252;.")
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub InventoryMenu()

	Dim db, sql, rs

	CardTitle = "Inventar"
	CardCurrent = "Inventory"
	CardCurrentLevel = GetCardLevel()

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

	Set db = New ADOHelper

	Call StartMobileDocument(CardTitle)
		If IsPocketIE or IsBlackBerry Then %>
		<p align="center" mode="wrap">
		<b>Inventar Menu</b>
		</p><%
		End If
        %>
        <p align="center">
        <a href="default.asp?card=viewinventory&amp;s=<% =SessionID %>">Inventar Ansehen</a><br/>
        <a href="default.asp?card=newitem&amp;s=<% =SessionID %>">Neues InventarArtikel</a><br/>
   		<a href="default.asp?card=adjustinventory&amp;s=<% =SessionID %>">Inventar anpassen</a><br/>
   		<a href="default.asp?card=countinventory&amp;s=<% =SessionID %>">Inventar z&#228;hlen</a><br/>
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

	CardTitle = "Neuer AA"
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
	%>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			Neuer Arbeitsauftrag
		</div>

		<div style='padding:5px;'>
	<%
		OutputFields
		WONewSubmit rs
	%>
	</div>
	<%
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub WONewSubmit(rs)
	%>
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
	<input type="hidden" name="card" value="<% =GetSession("ParentCard" & (CardCurrentLevel-1)) %>"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="wonew" value="y"/>
	<input type="hidden" name="preventduplicationsubmit" value="<% =RandomString(7) %>"/>
	<%
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
	CardTitle = "Inventar z&#228;hlen"
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
	            HeaderMsg = "Artikel-ID existiert nicht."
	        End If
	        If HeaderMsg = "" Then
                sql = "SELECT LocationPK FROM Location WITH (NOLOCK) WHERE LocationID = '" & Request("LocationID") & "'"
                Set rs = db.RunSQLReturnRS(sql,"")
	            Call CheckDB(db)
	            If rs.Eof Then
    	            HeaderMsg = "Standort-ID ist nicht vorhanden."
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
    	            HeaderMsg = "Spezifische Artikel-ID ist nicht bei spezifischer Standort-ID vorhanden."
	            End If
	        End If
            If Not HeaderMSG = "" Then
                CardCurrent = "CountInventory"
                CardCurrentLevel = GetCardLevel()
	        End If
	    Else
	        If CardCurrent = "CountInventory2" Then
	            HeaderMSG = "Bitte geben Sie die Artikel-ID und die Standort-ID genauer an."
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
	%>
		<script type='text/javascript'>
			$(document).ready(function(){
				try{
					$('#PartID').focus();
				}
				catch(e){}
			});
		</script>
	<%

        If Not CardCurrent = "CountInventory" Then
            If (Not HeaderMSG = "") or (Not Request("PartID") = "" and Not Request("LocationID") = "") Then
	%>
			<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
				<div style='float:left;'>
					<div>
						<%=rs("PartID")%> (<%=rs("PartName")%>)
					</div>
					<div class='Font2' style='font-size:11pt; padding-top:10px;'>
						<%=rs("LocationID")%> (<%=rs("LocationName")%>)
					</div>
				</div>
				<div style='float:right; text-align:right;'>
					<div>
						Inventar z&#228;hlen
					</div>
					<div class='Font1' style='font-size:11pt; padding-top:10px;'>
						Verf&#252;gbar: <%=rs("OnHand")%>
					</div>
				</div>
				<div style='clear:both;'></div>
			</div>
	<%
            End If
        End If
	%>
		<div style='padding:15px; clear:both;'>
	<%
		OutputFields
	%>
		</div>
	<%
		CountInventorySubmit rs

		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub CountInventorySubmit(rs)
	%>
	<div style='padding:15px;'>
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
	<input type="hidden" name="card" value="CountInventory"/>
	<input type="hidden" name="s" value="<% =SessionID %>"/>
	<input type="hidden" name="pk" value="<% =pk %>"/>
	<input type="hidden" name="posted" value="Y"/>
	</div>
	<%
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
				HeaderMSG = "Der Wert f&#252;r Neue Z&#228;hlung ist ung&#252;ltig."
				Exit Sub
			End If

            If Not Request("Bin") = "" Then
                rs("Bin") = Request("Bin")
		    Else
		        rs("Bin") = Null
            End If
			If Err.Number <> 0 Then
				HeaderMSG = "Der Wert f&#252;r M&#252;lleimer ist ung&#252;ltig."
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
	CardTitle = "Ausr&#252;stungs-Aufgaben"
	CardCurrentLevel = GetCardLevel()

	Set db = New ADOHelper

	If Request("Posted") = "Y" AND Not Request("AssetID") = "" Then
		' Validate AssetID
		sql = "SELECT AssetPK FROM Asset WITH (NOLOCK) WHERE AssetID = '" & Request("AssetID") & "'"
		Set rs = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs.Eof Then
			HeaderMsg = "Ausr&#252;stungs-ID ist nicht vorhanden."
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
			HeaderMSG = "Bitte geben Sie die Ausr&#252;stungs-ID genauer an."
		End If
	End If

	Call SetSession("ParentCard" & CardCurrentLevel,CardCurrent)
	Call SetSession("CardFrom",CardCurrent)
	Call SetSession("CardFromLevel",CardCurrentLevel)

    AssetID = ""
    Call BuildFields("AssetID","Asset ID","C",GlobalFieldLength,"*M","true","",rs,-1)

    pk = ""

	Call StartMobileDocument(CardTitle)
	headerString = GetBackButton("history.back();")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
	%>
		<script type='text/javascript'>
			$(document).ready(function(){
				try{
					$('#AssetID').focus();
				}
				catch(e){}
			});
		</script>
	<%
		Response.Write "<div style='padding:5px;'>"
		OutputFields

		AssetTasksSubmit rs
		Response.Write "</div>"

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
	<input style="width:100%;" type="submit" name="submit" value="Zulassen"/>
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
	Dim LaborReportExisting

	woaction = Request("woaction")
	If Not woaction = "" Then

		sql = "SELECT * FROM WO WITH (NOLOCK) WHERE WOPK = " & WOPK
		Set rs = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)

		If Not rs.Eof Then

			laborreport = NullCheck(Request("LaborReport"))
			LaborReportExisting = Trim(NullCheck(rs("LaborReport")))

			If Not NullCheck(Request("AccountID")) = "" Then
				sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
				Set rs2 = db.RunSQLReturnRS(sql,"")
				Call CheckDB(db)
				If rs2.eof Then
					HeaderMSG = "Keine Kostenstelle gefunden."
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
					HeaderMSG = "Kategorien-ID konnte nicht gefunden werden."
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
					HeaderMSG = "Problem-ID konnte nicht gefunden werden."
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
					HeaderMSG = "Fehler-ID konnte nicht gefunden werden."
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
					HeaderMSG = "L&#246;sungs-ID konnte nicht gefunden werden."
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
					HeaderMSG = "Der Wert f&#252;r Datum muss ausgef&#252;llt werden."
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
If UCase(woaction) = "WODETAILS" Then

				Dim ShopPK, ShopID, ShopName
				Dim DepartmentPK, DepartmentID, DepartmentName
				Dim TenantPK, TenantID, TenantName
				Dim AssetPK, AssetID, AssetName
				Dim ProjectPK, ProjectID, ProjectName
				Dim prefvalue, prefdesc, prefpk
				Dim txtType, txtTypeDesc, txtPriority, txtPriorityDesc

				AssetPK = ""
				AssetID = ""
				AssetName = ""

				ShopPK = ""
				ShopID = ""
				ShopName = ""

				DepartmentPK = ""
				DepartmentID = ""
				DepartmentName = ""

				TenantPK = ""
				TenantID = ""
				TenantName = ""

				ProjectPK = ""
				ProjectID = ""
				ProjectName = ""

				If Not NullCheck(Request("AssetID")) = "" Then
					sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetID = '" & NullCheck(Request("AssetID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Ausr&#252;stungs-ID nicht gefunden"
					Else
						AssetPK = rs2("AssetPK")
						AssetID = rs2("AssetID")
						AssetName = rs2("AssetName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Ausr&#252;stungs-ID ist ung&#252;ltig."
					End If
				End If

				If Not NullCheck(Request("Priority")) = "" Then
					sql = "SELECT * FROM LookupTableValues WITH (NOLOCK) WHERE LookupTable = 'WOPriority' AND CodeName = '" & NullCheck(Request("Priority")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Priorit&#228;t konnte nicht gefunden werden."
					Else
						txtPriority = rs2("CodeName")
						txtPriorityDesc = rs2("CodeDesc")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Priorit&#228;t ist ung&#252;ltig."
					End If
				End If

				If Not NullCheck(Request("Type")) = "" Then
					sql = "SELECT * FROM LookupTableValues WITH (NOLOCK) WHERE LookupTable = 'WOType' AND CodeName = '" & NullCheck(Request("Type")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Typ konnte nicht gefunden werden."
					Else
						txtType = rs2("CodeName")
						txtTypeDesc = rs2("CodeDesc")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Typ ist ung&#252;ltig."
					End If
				End If

				If Not NullCheck(Request("ProjectID")) = "" Then
					sql = "SELECT * FROM Project WITH (NOLOCK) WHERE ProjectID = '" & NullCheck(Request("ProjectID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Projekt-ID konnte nicht gefunden werden."
					Else
						ProjectPK = rs2("ProjectPK")
						ProjectID = rs2("ProjectID")
						ProjectName = rs2("ProjectName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Projekt-ID ist ung&#252;ltig."
					End If
				End If

				If Not NullCheck(Request("DepartmentID")) = "" Then
					sql = "SELECT * FROM Department WITH (NOLOCK) WHERE DepartmentID = '" & NullCheck(Request("DepartmentID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Abteilungs-ID konnte nicht gefunden werden."
					Else
						DepartmentPK = rs2("DepartmentPK")
						DepartmentID = rs2("DepartmentID")
						DepartmentName = rs2("DepartmentName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Abteilungs-ID ist ung&#252;ltig."
					End If
				End If

				If Not NullCheck(Request("TenantID")) = "" Then
					sql = "SELECT * FROM Tenant WITH (NOLOCK) WHERE TenantID = '" & NullCheck(Request("TenantID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Kunden-ID konnte nicht gefunden werden."
					Else
						TenantPK = rs2("TenantPK")
						TenantID = rs2("TenantID")
						TenantName = rs2("TenantName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Kunden-ID ist ung&#252;ltig."
					End If
				End If

				If Not NullCheck(Request("ShopID")) = "" Then
					sql = "SELECT * FROM Shop WITH (NOLOCK) WHERE ShopID = '" & NullCheck(Request("ShopID")) & "' "
					Set rs2 = db.RunSQLReturnRS(sql,"")
					Call CheckDB(db)
					If rs2.eof Then
						HeaderMSG = "Werkstatt-ID konnte nicht gefunden werden."
					Else
						ShopPK = rs2("ShopPK")
						ShopID = rs2("ShopID")
						ShopName = rs2("ShopName")
					End If
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Werkstatt-ID ist ung&#252;ltig."
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

				'On Error Resume Next

				'Set rs = db.RunSqlReturnRS("Select GETDATE() AS ServerDate","")
				'Call CheckDB(db)
				'ServerDate = rs("ServerDate")

				Set rs = db.RunSQLReturnRS_RW("SELECT * FROM WO WHERE WOPK = " & WOPK,"")
				Call CheckDB(db)

				On Error Resume Next

				If Not AssetPK = "" Then
					rs("AssetPK") = NullCheck(AssetPK)		' Nullable: YES Type: int
					rs("AssetID") = NullCheck(AssetID)		' Nullable: YES Type: nvarchar
					rs("AssetName") = NullCheck(AssetName)	' Nullable: YES Type: nvarchar
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Ausr&#252;stungs-ID ist ung&#252;ltig."
				End If

				If Not txtType = "" Then
					rs("Type") = Trim(Mid(txtType,1,25))	' Nullable: YES Type: nvarchar
					rs("TypeDesc") = NullCheck(txtTypeDesc)	' Nullable: YES Type: nvarchar
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Typ ist ung&#252;ltig."
					End If
				End If

				If Not txtPriority = "" Then
					rs("Priority") = Trim(Mid(txtPriority,1,25))	' Nullable: No Type: nvarchar
					rs("PriorityDesc") = NullCheck(txtPriorityDesc)	' Nullable: YES Type: nvarchar
					If Err.Number <> 0 Then
						HeaderMSG = "Der Wert f&#252;r Priorit&#228;t ist ung&#252;ltig."
					End If
				End If

				If Not ShopPK = "" Then
					rs("ShopPK") = ShopPK
					rs("ShopID") = ShopID
					rs("ShopName") = ShopName
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Werkstatt-ID ist ung&#252;ltig."
				End If

				If Not DepartmentPK = "" Then
					rs("DepartmentPK") = DepartmentPK
					rs("DepartmentID") = DepartmentID
					rs("DepartmentName") = DepartmentName
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Abteilungs-ID ist ung&#252;ltig."
				End If

				If Not TenantPK = "" Then
					rs("TenantPK") = TenantPK
					rs("TenantID") = TenantID
					rs("TenantName") = TenantName
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Kunden-ID ist ung&#252;ltig."
				End If

				If Not ProjectPK = "" Then
					rs("ProjectPK") = ProjectPK
					rs("ProjectID") = ProjectID
					rs("ProjectName") = ProjectName
				End If
				If Err.Number <> 0 Then
					HeaderMSG = "Der Wert f&#252;r Projekt-ID ist ung&#252;ltig."
				End If

				rs("RowVersionAction") = "EDIT"
				rs("RowVersionIPAddress") = GetSession("UserIPAddress")	' Nullable: YES Type: int
				rs("RowVersionUserPK") = GetSession("UserPK")	' Nullable: YES Type: int
				rs("RowVersionInitials") = Trim(Mid(GetSession("UserInitials"),1,5))									' Nullable: No Type: nvarchar
				rs("RowVersionDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))

				On Error Goto 0

				db.dobatchupdate rs

				Call CheckDB(db)

				'HeaderMSG = db.derror
				'HeaderMSG = "Error Saving"

				If Not HeaderMSG = "" Then

					If Not UCase(card) = "WOOPTIONS" Then
						CardSkipLevel = 1
					End If
					Card = WOAction
					WODetails

				End If

			Else
			' ************************************************************************************************************
			' CLOSE SINGLE WORK ORDER
			' ************************************************************************************************************
			If GetDefaultPreference(db,False,"WO_CLOSE_PREFIXLABORRPT",prefvalue, prefdesc, prefpk) Then
								If Trim(UCase(prefvalue)) = "YES" Then
									If LaborReportExisting = "" Then
										laborreport = GetSession("UserInitials") & " " & CStr(Date()) & " " & laborreport
									Else
										laborreport = Chr(13) & GetSession("UserInitials") & " " & CStr(Date()) & " " & laborreport
									End If
								End If
							End If

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
End If
End Sub

'====================================================================================================================================

Sub MyWorkOrders()

	Dim db, sql, rs, parentlocationall, parentequipmentall, parentall

	CardTitle = "Meine AAs"
	CardCurrent = "MyWorkOrders"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

	Dim iCount
	iCount = 0
	If Not sqlewhere = "" Then %>
	<div style='background-color:#fff0de; border: 2px solid #f8981d;'>
		<% =sqlewhere %>
	</div>
	<% End If
	If rs.eof Then
		Call OutputWAPMsg("Keine Zuweisung von Arbeitsauftr&#228;gen gefunden.")
	Else %>
		<div>
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
            	If iCount Mod 2 = 0 Then
			%>
				<div class='Font1' style='padding:5px;background-color:#f0f7fe; font-size:14pt; cursor:pointer;' onclick="location.href='default.asp?card=wooptions&s=<% =SessionID %>&wopk=<% =rs("WOPK") %>';">
					<div style='float:left;'><img src="images/icons/48/User Labor Male.png" alt="" title=""></div>
					<div style='float:left;'>
						<div>
						<% =rs("WOPK") %>:
						<% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) & " " %>
						</div>
						<div class='Font2' style='font-size:12pt;'>
							<% If Not NullCheck(rs("AssetID")) = "" Then %>
								<% =Shorten(ParentAll, 20) %>
								<% =Shorten(WapValidate(NullCheck(rs("AssetName"))),20) %>
									<% If Not BitNullCheck(rs("IsLocation")) Then Response.Write " (" & Shorten(NullCheck(rs("AssetID")),16) & ")" End If %>
							<% End If %>
						</div>
					</div>
					<div style='float:right'>
						<% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %>
					</div>
					<div style='clear:both;'></div>
				</div>
			<%
				Else
			%>
				<div class='Font1' style='padding:5px;background-color:#ffffff; font-size:14pt; cursor:pointer;' onclick="location.href='default.asp?card=wooptions&s=<% =SessionID %>&wopk=<% =rs("WOPK") %>';">
					<div style='float:left;'><img src="images/icons/48/User Labor Male.png" alt="" title=""></div>
					<div style='float:left;'>
						<div>
						<% =rs("WOPK") %>:
						<% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) & " " %>
						</div>
						<div class='Font2' style='font-size:12pt;'>
							<% If Not NullCheck(rs("AssetID")) = "" Then %>
								<% =Shorten(ParentAll, 20) %>
								<% =Shorten(WapValidate(NullCheck(rs("AssetName"))),20) %>
									<% If Not BitNullCheck(rs("IsLocation")) Then Response.Write " (" & Shorten(NullCheck(rs("AssetID")),16) & ")" End If %>
							<% End If %>
						</div>
					</div>
					<div style='float:right'>
						<% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %>
					</div>
					<div style='clear:both;'></div>
				</div>
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
		iCount = iCount + 1
		Loop
		%>
		</div>
		<%
	End If
	%>
	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
	<%
	rs.Close()
	Set db = Nothing
	SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub AllWorkOrders()

	Dim db, sql, rs, parentlocationall, parentequipmentall, parentall

	CardTitle = "Alle AAs"
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If Not sqlewhere = "" Then %>
			<div style='background-color:#fff0de; border: 2px solid #f8981d;">
				<% =sqlewhere %>
			</div>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Arbeitsauftr&#228;ge gefunden werden.")
		Else %>
			<div>
			<%
			dim iCount
			iCount = 0
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
                If iCount Mod 2 = 0 Then
				%>
				<div class='Font1' style='padding:5px;background-color:#f0f7fe; font-size:14pt; cursor:pointer;' onclick="location.href='default.asp?card=wooptions&s=<% =SessionID %>&wopk=<% =rs("WOPK") %>';">
					<div style='float:left;'><img src="images/icons/48/User Group Labor.png" alt="" title=""></div>
					<div style='float:left; padding-left:5px; width:50px;'>
						<% If BitNullCheck(rs("IsAssigned")) Then %>
							<img src="images/icons/48/Calendar Confirmed.png" alt="" title="">
						<% Else %>
							&nbsp;
						<% End If %>
					</div>
					<div style='float:left; padding-left:5px;'>
						<div>
						<% =rs("WOPK") %>:
						<% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) & " " %>
						</div>
						<div class='Font2' style='font-size:12pt;'>
							<% If Not NullCheck(rs("AssetID")) = "" Then %>
								<% =Shorten(ParentAll, 20) %>
								<% =Shorten(WapValidate(NullCheck(rs("AssetName"))),20) %>
									<% If Not BitNullCheck(rs("IsLocation")) Then Response.Write " (" & Shorten(NullCheck(rs("AssetID")),16) & ")" End If %>
							<% End If %>
						</div>
					</div>
					<div style='float:right'>
						<% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %>
					</div>
					<div style='clear:both;'></div>
				</div>
				<%Else%>
				<div class='Font1' style='padding:5px;background-color:#ffffff; font-size:14pt; cursor:pointer;' onclick="location.href='default.asp?card=wooptions&s=<% =SessionID %>&wopk=<% =rs("WOPK") %>';">
					<div style='float:left;'><img src="images/icons/48/User Group Labor.png" alt="" title=""></div>
					<div style='float:left; padding-left:5px; width:50px;'>
						<% If BitNullCheck(rs("IsAssigned")) Then %>
							<img src="images/icons/48/Calendar Confirmed.png" alt="" title="">
						<% Else %>
							&nbsp;
						<% End If %>
					</div>
					<div style='float:left; padding-left:5px;'>
						<div>
						<% =rs("WOPK") %>:
						<% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) & " " %>
						</div>
						<div class='Font2' style='font-size:12pt;'>
							<% If Not NullCheck(rs("AssetID")) = "" Then %>
								<% =Shorten(ParentAll, 20) %>
								<% =Shorten(WapValidate(NullCheck(rs("AssetName"))),20) %>
									<% If Not BitNullCheck(rs("IsLocation")) Then Response.Write " (" & Shorten(NullCheck(rs("AssetID")),16) & ")" End If %>
							<% End If %>
						</div>
					</div>
					<div style='float:right'>
						<% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %>
					</div>
					<div style='clear:both;'></div>
				</div>
				<%End If%>
				<% If False Then %>
				<div style='display:none;'>
				<select tabindex="2" name="S<% =rs("WOPK") %>" value="I">
				<option value="IS">Ausgestellt</option>
				<option value="OH">Warteschleife</option>
				<option value="CO">Abgeschlossen</option>
				<option value="CL">Geschlossen</option>
				</select>
				Hours: <input tabindex="1" size="2" name="H<% =rs("WOPK") %>" value=""/>
				<br/>
				Report: <input tabindex="4" size="10" name="R<% =rs("WOPK") %>" value=""/>
				<br/>
				</div>
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
			iCount = iCount + 1
			Loop
			%>
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub


'====================================================================================================================================

Sub AllWorkOrdersU()

	Dim db, sql, rs, parentlocationall, parentequipmentall, parentall

	CardTitle = "Alle nicht zugewiesenen AAs"
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
	headerString = GetBackButton("history.back();")
	headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If Not sqlewhere = "" Then %>
		<div style='background-color:#fff0de; border: 2px solid #f8981d;">
			<% =sqlewhere %>
		</div>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Arbeitsauftr&#228;ge gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%
			dim iCount, altStyle
			iCount = 0
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
                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if
				%>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="location.href='default.asp?card=wooptions&s=<% =SessionID %>&wopk=<% =rs("WOPK") %>';">
					<div style='float:left;'><img src="images/icons/48/User Labor Female.png" alt="" title=""></div>
					<div style='float:left;'>
						<div>
						<% =rs("WOPK") %>:
						<% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) & " " %>
						</div>
						<div class='Font2' style='font-size:12pt;'>
							<% If Not NullCheck(rs("AssetID")) = "" Then %>
								<% =Shorten(ParentAll, 20) %>
								<% =Shorten(WapValidate(NullCheck(rs("AssetName"))),20) %>
									<% If Not BitNullCheck(rs("IsLocation")) Then Response.Write " (" & Shorten(NullCheck(rs("AssetID")),16) & ")" End If %>
							<% End If %>
						</div>
					</div>
					<div style='float:right'>
						<% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %>
					</div>
					<div style='clear:both;'></div>
				</div>
				<% If False Then %>
				<div style='display:none;'>
				<select tabindex="2" name="S<% =rs("WOPK") %>" value="I">
				<option value="IS">Ausgestellt</option>
				<option value="OH">Warteschleife</option>
				<option value="CO">Abgeschlossen</option>
				<option value="CL">Geschlossen</option>
				</select>
				Hours: <input tabindex="1" size="2" name="H<% =rs("WOPK") %>" value=""/>
				<br/>
				Report: <input tabindex="4" size="10" name="R<% =rs("WOPK") %>" value=""/>
				<br/></div>
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
			iCount = iCount + 1
			Loop
			%>
			</p>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CMLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Firmen"
	CardCurrent = "CMLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	'headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)
		Dim iCount, altStyle
		iCount = 0
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnte keine Firma gefunden werden.")
		Else %>
			<div>
			<%
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			If Not ChoiceIndex > PagePos Then
				iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "javascript:parent.document.mcform." & ft & ".value='" & JSEncode(rs("CompanyID")) & "'; parent.CloseWindow();" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(NullCheck(rs("CompanyName"))) %> (<% =WAPValidate(NullCheck(rs("CompanyID"))) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

				<%
				iCount = iCount + 1
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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
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
			CardTitle = "Suchfeld f&#252;r Probleme"
			desc = "Problems"
			desc2 = "Problem"
		Case "FAFALOOKUP"
			CardTitle = "Suchfeld f&#252;r Fehler"
			desc = "Failures"
			desc2 = "Failure"
		Case "FASOLOOKUP"
			CardTitle = "Suchfeld f&#252;r L&#246;sungen"
			desc = "Solutions"
			desc2 = "Solution"
	End Select
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	'headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)
		Dim iCount, altStyle
		iCount = 0
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg( desc &"  konnte nicht gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if

            If Not ChoiceIndex > PagePos Then
            	iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "parent.document.mcform." & desc2 & "ID.value='" & JSEncode(rs("FailureID")) & "'; parent.CloseWindow();" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("FailureName")) %> (<% =WAPValidate(rs("FailureID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

<%
				iCount = iCount + 1
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
		%><div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
		<%
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

	CardTitle = "Ausr&#252;stung"
	CardCurrent = "Assets"
	CardCurrentLevel = GetCardLevel()

	pagesize = GlobalPageSize

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
	headerString = GetBackButton("history.back();")
	headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If Not sqlewhere = "" Then %>
	<div style='background-color:#fff0de; border: 2px solid #f8981d;">
		<% =sqlewhere %>
	</div>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnte keine Ausr&#252;stung gefunden werden.")
		Else %>
			<div>
			<%

            Dim NoIconFile_Location, NoIconFile_Asset, IconFile, ParentAll, parentlocationall, parentequipmentall

			NoIconFile_Location = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/facility_g.gif"
			NoIconFile_Asset = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/gearsxp_g.gif"

            If Not ParentPK = "" Then %>
				<div class='Font1' style='white-space: nowrap; clear:both;padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;' onclick="location.href = 'default.asp?card=assets&amp;s=<% =SessionID %>&amp;assetpk=<% =ParentPK %>&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>';">
					<div style="float:left;">
						Eine Ebene hoch
					</div>
					<div style="float:right;">
						<img src="images/icons/48/Navigation 2 Left.png" style='width:36px; height:36px; border:none;' />
					</div>
					<div style="clear:both;"></div>
				</div>

            <% If False Then %>
            	<% If lang="HTML" Then %>&nbsp;<% End If %>
            		<div class='Font1 RowElisp' style='clear:both;padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer; max-width:75%;' onclick="location.href = 'default.asp?card=assets&amp;s=<% =SessionID %>&amp;assetpk=&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>';>
            			Zur&#252;ck zum Anfang
            		</div>
            	<% End If %>
            	<% =LStyleEnd %>
            <%
            End If

			dim iCount, altStyle
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

				  If BitNullCheck(rs("islocation")) Then
					IconFile = NoIconFile_Location
				  Else
					IconFile = NoIconFile_Asset
				  End If

                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if
                'Symbol Add 2
                returnurl = "default.asp?card=asoptions&s=" & SessionID & "&AssetPK=" & WAPEncode(rs("AssetPK"))
				%>
				<div class='Font1' style='white-space: nowrap; clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
					<div style='float:left; width:80%;' onclick="location.href = '<% =returnurl %>';">
						<div class='Font1 RowElisp' style='font-size:16pt; max-width:100%;'>
							<% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %>
								&nbsp;
							<% Else %>
								<% =ParentAll %>
							<% End If %>
							<% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "<br /><span class='Font2' style='font-size:12pt;padding-left:10px;'>Ausr&#252;stungs-ID " & WAPValidate(rs("AssetID")) & "</span>" End If %>
						</div>
					</div>
					<% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %>
						<% If BitNullCheck(rs("HasChildren")) Then %>
							<div class='Font1' style='font-size:12pt; float:right; position:relative; padding-right:5px;' onclick="location.href = 'default.asp?card=assets&amp;s=<% =SessionID %>&amp;assetpk=<% =rs("AssetPK") %>&amp;pagepos=1&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>';">
								<img src="images/icons/48/Navigation 2 Right.png" style='width:36px; height:36px; border:none;' />
							</div>
						<%else%>
							<div class='Font1' style='font-size:12pt; float:right; position:relative; padding-right:5px;'>
								<img src="images/icons/48/blank.png" style='width:36px; height:36px; border:none;' />
							</div>
						<% End If %>
					<% End If %>
					<div style='clear:both;'></div>
				</div>
				<%

				RegNode = True
				iCount = iCount + 1
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
			</div>
			<%
		End If
		%>
	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&assetpk=<% =AssetPK %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&assetpk=<% =AssetPK %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASLookup()

	Dim db, sql, rs, rsp, assetpk, regnode, parentpk

	CardTitle = "Suchfeld f&#252;r Ausr&#252;stung"
	CardCurrent = "ASLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	'headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)

		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnte keine Ausr&#252;stung gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%

            Dim NoIconFile_Location, NoIconFile_Asset, IconFile, ParentAll, parentlocationall, parentequipmentall

			NoIconFile_Location = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/facility_g.gif"
			NoIconFile_Asset = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/gearsxp_g.gif"

            If Not ParentPK = "" Then %>
				<div class='Font1' style='white-space: nowrap; clear:both;padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;' onclick="location.href = 'default.asp?card=aslookup&amp;s=<% =SessionID %>&amp;assetpk=<% =ParentPK %>&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>';">
					<div style="float:left;">
						Eine Ebene hoch
					</div>
					<div style="float:right;">
						<img src="images/icons/48/Navigation 2 Left.png" style='width:36px; height:36px; border:none;' />
					</div>
					<div style="clear:both;"></div>
				</div>

            <%
            End If

			dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if
			If Not ChoiceIndex > PagePos Then
				iCount = iCount + 1
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
					returnurl = "parent.document.mcform." & ft & ".value='" & JSEncode(rs("AssetID")) & "'; parent.CloseWindow();" %>

				<div class='Font1' style='white-space: nowrap; clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
					<div style='float:left; width:80%;' onclick="<% =returnurl %>;">
						<div class='Font1' style='font-size:16pt;'>
							<% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %>
								&nbsp;
							<% Else %>
								<% =ParentAll %>
							<% End If %>
							<% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "<br /><span class='Font2' style='font-size:12pt;padding-left:10px;'>Ausr&#252;stungs-ID " & WAPValidate(rs("AssetID")) & "</span>" End If %>
						</div>
					</div>
					<% If NullCheck(rs("AssetLevel")) > 1 and RegNode Then %>
						<% If BitNullCheck(rs("HasChildren")) Then %>
							<div class='Font1' style='font-size:12pt; float:right; position:relative; padding-right:5px;' onclick="location.href = 'default.asp?card=aslookup&amp;s=<% =SessionID %>&amp;assetpk=<% =rs("AssetPK") %>&amp;pagepos=1&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>';">
								<img src="images/icons/48/Navigation 2 Right.png" style='width:36px; height:36px; border:none;' />
							</div>
						<%else%>
							<div class='Font1' style='font-size:12pt; float:right; position:relative; padding-right:5px;'>
								<img src="images/icons/48/blank.png" style='width:36px; height:36px; border:none;' />
							</div>
						<% End If %>
					<% End If %>
					<div style='clear:both;'></div>
				</div>


					<%

				RegNode = True
				iCount = iCount + 1
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
		%><div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&assetpk=<% =AssetPK %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&assetpk=<% =AssetPK %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CLLookup()

	Dim db, sql, rs, rsp, classificationpk, regnode, parentpk

	CardTitle = "Suchfeld f&#252;r Klassifizierung"
	CardCurrent = "CLLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnte keine Klassifizierung gefunden werden.")
		Else %>
			<div>
			<%

            Dim NoIconFile_Location, NoIconFile_Asset, IconFile, ParentAll, parentlocationall, parentequipmentall

			NoIconFile_Location = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/facility_g.gif"
			NoIconFile_Asset = Application("MCVirtualDirectory") & Application("mapp_path") & "images/icons/gearsxp_g.gif"

            If Not ParentPK = "" Then %>
				<div class='Font1' style='white-space: nowrap; clear:both;padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;' onclick="location.href = 'default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;classificationpk=<% =ParentPK %>&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>';">
					<div style="float:left;">
						Eine Ebene hoch
					</div>
					<div style="float:right;">
						<img src="images/icons/48/Navigation 2 Left.png" style='width:36px; height:36px; border:none;' />
					</div>
					<div style="clear:both;"></div>
				</div>

            <% If False Then %>
            		<div class='Font1' style='white-space: nowrap; clear:both;padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;' onclick="location.href = 'default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;classificationpk=&amp;pagepos=0&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>';>
            			Zur&#252;ck zum Anfang
            		</div>
            	<%
            	End if
            End If

			Dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
				iCount = iCount + 1
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
					returnurl = "parent.document.mcform.ClassificationID.value='" & JSEncode(rs("ClassificationID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<div class='Font1' style='white-space: nowrap; clear:both;padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;'>
						<div style='float:left; width:80%;' onclick="<% =returnurl %>">
							<div class='Font1' style='font-size:16pt;'>
								<% If NullCheck(rs("ClassificationLevel")) > 1 and RegNode Then %>
								<% Else %>
									<% =ParentAll %>
								<% End If %>
								<% =WAPValidate(rs("ClassificationName")) & " " %><% If Not IsLocation Then Response.Write "<br /><span class='Font2' style='font-size:12pt;padding-left:10px;'>Klassifizierungs-ID: " & WAPValidate(rs("ClassificationID")) & "</span>" End If %>
							</div>
						</div>
						<% If NullCheck(rs("ClassificationLevel")) > 1 and RegNode Then %>
							<% If BitNullCheck(rs("HasChildren")) Then %>
								<div class='Font1' style='font-size:12pt; float:right; position:relative; padding-right:5px;' onclick="location.href = 'default.asp?card=cllookup&amp;s=<% =SessionID %>&amp;classificationpk=<% =rs("ClassificationPK") %>&amp;pagepos=1&amp;searchby=NONE&amp;flipit=<% =RandomString(3) %>';">
									<img src="images/icons/48/Navigation 2 Right.png" style='width:36px; height:36px; border:none;' />
								</div>
							<%else%>
								<div class='Font1' style='font-size:12pt; float:right; position:relative; padding-right:5px;'>
									<img src="images/icons/48/blank.png" style='width:36px; height:36px; border:none;' />
								</div>
							<% End If %>
						<% End If %>
						<div style='clear:both;'></div>
					</div>

				<%
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
			</div>
			<%
		End If
		%>
	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>&amp;Classificationpk=<% =ClassificationPK %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>&amp;Classificationpk=<% =ClassificationPK %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASForSerialPartLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Ausr&#252;stung"
	CardCurrent = "ASForSerialPartLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, True)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnte keine Ausr&#252;stung gefunden werden.")
		Else %>
			<div>
			<%
			Dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
					returnurl = "parent.document.mcform.SerialReplaceToLocationID.value='" & JSEncode(rs("AssetID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<%=returnurl%>">
						<div style='float:left;'>
							<div>
									<% =WAPValidate(rs("AssetName")) & " " %><% If Not BitNullCheck(rs("IsLocation")) Then Response.Write "(" & WAPValidate(rs("AssetID")) & ")" End If %>
							</div>
						</div>
						<div style='clear:both;'></div>
					</div>


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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub LTLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld"
	CardCurrent = "LTLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)

		Dim iCount, altStyle
		iCount = 0
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Werte gefunden werden.")
		Else %>
			<div>
			<%
			Do Until rs.Eof
                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if
			If Not ChoiceIndex > PagePos Then
				iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "parent.document.mcform." & Request("lf") & ".value='" & JSEncode(rs(Trim(Request("lr")))) & "'; parent.CloseWindow();" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<div style='float:left;'>
						<div>
							<% If rs("CodeDesc") = "" Then %>(Nicht festgelegt)<% Else %><% =WAPValidate(rs("CodeDesc")) & " " %>(<% =WAPValidate(rs("CodeName")) & ")" %><% End If %>
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

<%
				iCount = iCount + 1
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
			</div>
			<%
		End If
		%>		<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?lt=<%=request("lt")%>&amp;lf=<%=request("lf")%>&amp;lr=<%=request("lr")%>&amp;card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?lt=<%=request("lt")%>&amp;lf=<%=request("lf")%>&amp;lr=<%=request("lr")%>&amp;card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ACLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Kostenstelle"
	CardCurrent = "ACLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, True)
		Dim iCount, altStyle
		iCount = 0
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnte keine Kostenstelle gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if

            If Not ChoiceIndex > PagePos Then
            	iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else

				returnurl = "parent.document.mcform.AccountID.value='" & JSEncode(rs("AccountID")) & "'; parent.CloseWindow();" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("AccountID")) %> <% =WAPValidate(rs("AccountName")) %>
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>


				<%
            	iCount = iCount + 1
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
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub RCLookup()

	Dim db, sql, rs, rs2, anchoroutput

	CardTitle = "Suchfeld f&#252;r Reparaturzentrum"
	CardCurrent = "RCLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	'headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)
	If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Reparaturzentren gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%
			Dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
              if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if

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

				If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
					returnurl = "closeiframe('" & JSEncode(returnurl) & "'); parent.location.replace('" & JSEncode(returnurl) & "');"
				Else
					returnurl = "closeiframe('" & JSEncode(returnurl) & "'); parent.document.mcform.RepairCenterID.value='" & JSEncode(rs("RepairCenterID")) & "';"
				End If %>

				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<% =anchoroutput %>
				</div>

				<%

				iCount = iCount + 1
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
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub SHLookup()

	Dim db, sql, rs, rs2, anchoroutput, addall

	CardTitle = "Suchfeld f&#252;r Werkstatt"
	CardCurrent = "SHLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	'headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)

		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		dim iCount, altStyle
		iCount = 0
		If rs.eof and Not AddAll Then
			Call OutputWAPMsg("Es konnte keine Werkstatt gefunden werden.")
		Else %>
			<div>
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

					If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
						returnurl = "closeiframe('" & JSEncode(returnurl) & "');parent.location.replace('" & JSEncode(returnurl) & "');"
					Else
						returnurl = "closeiframe('" & JSEncode(returnurl) & "');parent.document.mcform.ShopID.value='ALL';"
					End If %>
				<div class='Font1' style='padding:5px;background-color:#fdf2c5; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<% =anchoroutput %>
				</div>

			<%
				'End If
			End If
			Do Until rs.Eof
                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if

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

				If UCase(GetSession("ParentCard" & (CardCurrentLevel-1))) = "MAINMENU" Then
					returnurl = "closeiframe('" & JSEncode(returnurl) & "');parent.location.replace('" & JSEncode(returnurl) & "');"
				Else
					returnurl = "closeiframe('" & JSEncode(returnurl) & "');parent.document.mcform.ShopID.value='" & JSEncode(rs("ShopID")) & "';"
				End If %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<% =anchoroutput %>
				</div>


				<%
				iCount = iCount + 1
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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub CALookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Kategorie"
	CardCurrent = "CALookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = Lookupsize
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
	headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, True)

		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Kategorien gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%
			Dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if

			If Not ChoiceIndex > PagePos Then
				iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "parent.document.mcform.CategoryID.value='" & JSEncode(rs("CategoryID")) & "'; parent.CloseWindow();" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("CategoryName")) %> (<% =WAPValidate(rs("CategoryID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>


				<%

				iCount = iCount + 1
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
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ZNLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Zone"
	CardCurrent = "ZNLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, True)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Zonen gefunden werden.")
		Else %>
			<div>
			<%
			Dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
					returnurl = "parent.document.mcform.ZoneID.value='" & JSEncode(rs("ZoneID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<%=returnurl%>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("ZoneName")) %> (<% =WAPValidate(rs("ZoneID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>


		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub PRLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Verfahren"
	CardCurrent = "PRLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, True)

		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		Dim iCount, altStyle
		iCount = 0
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Verfahren gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if

			If Not ChoiceIndex > PagePos Then
				iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "parent.document.mcform.ProcedureID.value='" & JSEncode(rs("ProcedureID")) & "'; parent.CloseWindow();" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("ProcedureName")) %> (<% =WAPValidate(rs("ProcedureID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

				<%
				iCount = iCount + 1
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
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub DPLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Abteilung"
	CardCurrent = "DPLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	'headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)
		Dim iCount, altStyle
		iCount = 0
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Abteilungen gefunden werden.")
		Else %>
			<div>
			<%

			Do Until rs.Eof
                if iCount Mod 2 = 0 Then
                	altStyle = "#f0f7fe"
                else
                	altStyle = "#ffffff"
                end if

			If Not ChoiceIndex > PagePos Then
				iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "parent.document.mcform.DepartmentID.value='" & JSEncode(rs("DepartmentID")) & "'; parent.CloseWindow();" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<% =returnurl %>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("DepartmentName")) %> (<% =WAPValidate(rs("DepartmentID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

				<%
				iCount = iCount + 1
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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub TNLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Kunden"
	CardCurrent = "TNLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	'headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)

		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		Dim iCount, altStyle
		iCount = 0
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Kunden gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			If Not ChoiceIndex > PagePos Then
				iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "javascript:parent.document.mcform.TenantID.value='" & JSEncode(rs("TenantID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<%=returnurl%>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("TenantName")) %> (<% =WAPValidate(rs("TenantID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

				<%
				iCount = iCount + 1
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
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub PJLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Projekt"
	CardCurrent = "PJLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize
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
	'headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)
		Dim iCount, altStyle
		iCount = 0
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Projekte gefunden werden.")
		Else %>
			<p mode="nowrap">
			<%
			Do Until rs.Eof
			iCount = iCount + 1
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else

					returnurl = "parent.document.mcform.ProjectID.value='" & JSEncode(rs("ProjectID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<%=returnurl%>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("ProjectName")) %> (<% =WAPValidate(rs("ProjectID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>


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
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>


		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub LALookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Arbeiter"
	CardCurrent = "LALookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize
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
	headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, True)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		Dim iCount, altStyle
		iCount = 0

		If rs.eof Then
			Call OutputWAPMsg("Es konnte kein Arbeiter gefunden werden.")
		Else %>
			<div>
			<%
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			If Not ChoiceIndex > PagePos Then
				iCount = iCount + 1
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "parent.document.mcform." & ft & ".value='" & JSEncode(rs("LaborID")) & "'; parent.CloseWindow();" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<%=returnurl%>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("LaborName")) %> (<% =WAPValidate(rs("LaborID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>



				<%
				iCount = iCount + 1
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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub INLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Artikel"
	CardCurrent = "INLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, True)
		If Not sqlewhere = "" Then %>
		<p>
			<b><% =sqlewhere %></b>
		</p>
		<% End If%>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<div style='float:left;'>
			</div>
			<div style='float:right; text-align:right;'>
				<div>
					Inventar
				</div>
			</div>
			<div style='clear:both;'></div>
		</div>

		<%
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Artikel gefunden werden.")
		Else %>
			<div>
			<%
			dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1
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
				returnurl = "parent.document.mcform." & ft & ".value='" & JSEncode(rs("PartID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<%=returnurl%>">
					<div style='float:left;'>
						<div>
							[<% =rs("PartID") %>] [<% =WAPValidate(rs("PartName")) %>]
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub SRLookup()

	Dim db, sql, rs

	CardTitle = "Suchfeld f&#252;r Standort"
	CardCurrent = "SRLookup"
	CardCurrentLevel = GetCardLevel()

	pagesize = LookupSize

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
	headerString = GetSearchButton()
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, True)
		If Not sqlewhere = "" Then %>
		<p>
		<b><% =sqlewhere %></b>
		</p>
		<% End If
		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Standorte gefunden werden.")
		Else %>
			<div>
			<%
			Dim iCount, altStyle
			iCount = 0
			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1

			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				returnurl = "parent.document.mcform.LocationID.value='" & JSEncode(rs("LocationID")) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
				<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer;' onclick="<%=returnurl%>">
					<div style='float:left;'>
						<div>
							<% =WAPValidate(rs("LocationName")) %> (<% =WAPValidate(rs("LocationID")) %>)
						</div>
					</div>
					<div style='clear:both;'></div>
				</div>

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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>

		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub ASHistory()

	Dim db, sql, rs

	CardTitle = "Ausr&#252;stungs-Verlauf"
	CardCurrent = "ASHistory"
	CardCurrentLevel = GetCardLevel()

	GetWOPK
	GetAssetPK
	If assetpk = "" Then
		assetpk = "-1"
	End If

	pagesize = GlobalPageSize
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
	"                      AssetHierarchy.ParentLocation, AssetHierarchy.ParentEquipment, ltp.CodeIcon AS PriorityIcon, lts.CodeIcon AS StatusIcon,ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!'), " + nl + _
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
	"                      AssetHierarchy.ParentLocation, AssetHierarchy.ParentEquipment, ltp.CodeIcon AS PriorityIcon, lts.CodeIcon AS StatusIcon, ParentLocationAll=REPLACE(AssetHierarchy.ParentLocationAll,'<br>','!#!'), ParentEquipmentAll=REPLACE(AssetHierarchy.ParentEquipmentAll,'<br>','!#!'), " + nl + _
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


	If UCase(CardFrom) = "WOSEARCH"	Then
		Pagepos = 0
	End If

	Call StartMobileDocument(CardTitle)
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)

		If rs.eof Then
			Call OutputWAPMsg("Es konnten keine Arbeitsauftr&#228;ge f&#252;r die spezifische Ausr&#252;stung gefunden werden.")
		Else %>
		<div class='Font1' style='padding:5px;background-color:#fbebc5; font-size:14pt; cursor:pointer;'>
			<% =rs("ParentLocationAll") %><% =rs("ParentEquipmentAll") %><% =rs("AssetName") %> (<% =rs("AssetID") %>)
		</div>
		<div>
			<%
			Dim iCount, altStyle,rURL
			iCount = 0

			Do Until rs.Eof
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if
			iCount = iCount + 1

			If Not ChoiceIndex > PagePos Then
				rs.MoveNext()
				ChoiceIndex = ChoiceIndex + 1
			Else
				rURL = "self.location.href = 'default.asp?card=wooptions&amp;s=" & SessionID & "&amp;wopk=" & rs("wopk") & "';"
				%>
				<div class='Font1' style='cursor:pointer;padding:5px;background-color:<%=altStyle%>; font-size:12pt;' onclick="<%=rURL%>">
					<div style='float:left;cursor:pointer;' onclick="<%=rURL%>">
						<div><% =rs("WOPK") %>: <% =WAPValidate(DateNullCheck(rs("TargetDate"))) & " " %></div>
						<div><% =WAPValidate(DateNullCheck(rs("CLOSED"))) & " " %></div>
					</div>
					<div class='Font2' style='float:right;cursor:pointer;' onclick="<%=rURL%>">
						<% =WAPValidate(Shorten(Replace(NullCheck(rs("Reason")),"%0D%0A",CHR(13) & CHR(10)),35)) %>
					</div>
					<div class='Font2' style='clear:both;'></div>
				</div>
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
			</div>
			<%
		End If
		%>	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Zur&#252;ck' title='Zur&#252;ck' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize Then
	%><a href="default.asp?card=<% =CardCurrent %>&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Weiter' title='Weiter' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
		<%
		rs.Close()
		Set db = Nothing
		SetContext
	EndWMLDocument

End Sub

Sub CalendarLookup()

	Dim db, sql, rs
	Dim datToday, intThisMonth, intThisYear, strMonthName, datFirstDay, intFirstWeekDay, intLastDay, intPrevMonth, intPrintDay, LastMonthDate, NextMonthDate, dFirstDay, dLastDay, EndRows, intLoopDay, intPrevYear, intNextMonth, intNextYear, intLastMonth, dToday, bEvents

	CardTitle = "Kalender"
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
	headerString = headerString & "<div onclick='parent.CloseWindow();' style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Symbol Restricted.png' alt='Fenster schlie&#223;en' title='Fenster schlie&#223;en' style='border:none; cursor:pointer; width:46px; height:46px;' /></div>"
	Call IPadHeader(headerString, False)
    %>
	<div>
	<div class='Font2' style='text-align:center;'>
    <a class='Font1' href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;month=<% =IntPrevMonth %>&amp;year=<% =IntPrevYear %>">[vorheriger]</a>
    &nbsp;&nbsp;<b><% = strMonthName & " " & intThisYear %></b>&nbsp;&nbsp;
    <a class='Font1' href="default.asp?card=<% =CardCurrent %>&amp;s=<% =SessionID %>&amp;month=<% =IntNextMonth %>&amp;year=<% =IntNextYear %>">[n&#228;chster]</a>

    </div>
    <div style='padding-left:25px; white-space:nowrap; padding-top:10px;'>
    <div>
    	<div class='CalHeader' style='float:left;'>So</div>
    	<div class='CalHeader' style='float:left;'>Mo</div>
    	<div class='CalHeader' style='float:left;'>Di</div>
    	<div class='CalHeader' style='float:left;'>Mi</div>
    	<div class='CalHeader' style='float:left;'>Do</div>
    	<div class='CalHeader' style='float:left;'>Fr</div>
    	<div class='CalHeader' style='float:left;'>Sa</div>
    	<div style='clear:both;'></div>
    </div>
	<%
			EndRows = False

			Dim iCount
			iCount = 0
		 	Do While EndRows = False
		 		iCount = iCount + 1
				' This is the loop for the days in the week
				For intLoopDay = cSUN To cSAT
					' If the first day is not sunday then print the last days of previous month in grayed font
					If intFirstWeekDay > cSUN Then
						Write_CalDay "<div class='CalOff' style='float:left;'>" & CheckDigits(LastMonthDate) & "</div>", "NON"
						LastMonthDate = LastMonthDate + 1
						intFirstWeekDay = intFirstWeekDay - 1
					' The month starts on a sunday
					Else
						' If the dates for the month are exhausted, start printing next month's dates
						' in grayed font
						If intPrintDay > intLastDay Then
							Write_CalDay "<div class='CalOff' style='float:left;'>" & CheckDigits(NextMonthDate) & "</div>", "NON"
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
					                returnurl = "parent.document.mcform." & ft & ".value='" & JSEncode(dToday) & "';closeiframe('" & JSEncode(returnurl) & "');" %>
					                <%
					                	If iCount >= 8 Then
					                		iCount = 0
											Response.Write "<div style='clear:both;'></div>"
					                	End If
					                %>
					                <div class='CalOn' style='float:left; cursor:pointer;' onclick="<% =returnurl %>"><% If dToday = Date Then %><b><% End If %><% =CheckDigits(intPrintDay) %><% If dToday = Date Then %></b><% End If %></div><%
				                Else %>
					                <% =LStyleBegin %><% If dToday = Date Then %><b><% End If %><anchor><% =CheckDigits(intPrintDay) %><go href="<% =returnurl %>"><setvar name="<% =ft %>" value="<% =WAPEncode(dToday) %>"/></go></anchor> <% If dToday = Date Then %></b><% End If %><% =LStyleEnd %> <%
				                End If

							End If
						End If

						' Increment the date. Done once in the loop.
						intPrintDay = intPrintDay + 1
					End If
				' Move to the next day in the week
				iCount = iCount + 1
				Next
			Loop
	%>
	<div style='clear:both;'></div>
	</div></div>
    <%
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
		HeaderMSG = "Der Grund muss ausgef&#252;llt werden."
		Call WONew()
	End If

	If Not NullCheck(Request("AssetID")) = "" Then
		sql = "SELECT * FROM Asset WITH (NOLOCK) WHERE AssetID = '" & NullCheck(Request("AssetID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Ausr&#252;stungs-ID nicht gefunden"
			Call WONew()
		Else
			AssetPK = rs2("AssetPK")
			AssetID = rs2("AssetID")
			AssetName = rs2("AssetName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Ausr&#252;stungs-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("ProblemID")) = "" Then
		sql = "SELECT * FROM Failure WITH (NOLOCK) WHERE FailureID = '" & NullCheck(Request("ProblemID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Problem-ID konnte nicht gefunden werden."
			Call WONew()
		Else
			ProblemPK = rs2("FailurePK")
			ProblemID = rs2("FailureID")
			ProblemName = rs2("FailureName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Problem-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("ProcedureID")) = "" Then
		sql = "SELECT * FROM ProcedureLibrary WITH (NOLOCK) WHERE ProcedureID = '" & NullCheck(Request("ProcedureID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Vorgangs-ID konnte nicht gefunden werden."
			Call WONew()
		Else
			ProcedurePK = rs2("ProcedurePK")
			ProcedureID = rs2("ProcedureID")
			ProcedureName = rs2("ProcedureName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Verfahrens-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("CategoryID")) = "" Then
		sql = "SELECT * FROM Category WITH (NOLOCK) WHERE CategoryID = '" & NullCheck(Request("CategoryID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Kategorien-ID konnte nicht gefunden werden."
			Call WONew()
		Else
			CategoryPK = rs2("CategoryPK")
			CategoryID = rs2("CategoryID")
			CategoryName = rs2("CategoryName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Kategorie-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("AccountID")) = "" Then
		sql = "SELECT * FROM Account WITH (NOLOCK) WHERE AccountID = '" & NullCheck(Request("AccountID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Keine Kostenstelle gefunden."
			Call WONew()
		Else
			AccountPK = rs2("AccountPK")
			AccountID = rs2("AccountID")
			AccountName = rs2("AccountName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Die angegebene Kostenstelle-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("Priority")) = "" Then
		sql = "SELECT * FROM LookupTableValues WITH (NOLOCK) WHERE LookupTable = 'WOPriority' AND CodeName = '" & NullCheck(Request("Priority")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Priorit&#228;t konnte nicht gefunden werden."
			Call WONew()
		Else
			txtPriority = rs2("CodeName")
			txtPriorityDesc = rs2("CodeDesc")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Priorit&#228;t ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("Type")) = "" Then
		sql = "SELECT * FROM LookupTableValues WITH (NOLOCK) WHERE LookupTable = 'WOType' AND CodeName = '" & NullCheck(Request("Type")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Typ konnte nicht gefunden werden."
			Call WONew()
		Else
			txtType = rs2("CodeName")
			txtTypeDesc = rs2("CodeDesc")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Typ ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("ProjectID")) = "" Then
		sql = "SELECT * FROM Project WITH (NOLOCK) WHERE ProjectID = '" & NullCheck(Request("ProjectID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Projekt-ID konnte nicht gefunden werden."
			Call WONew()
		Else
			ProjectPK = rs2("ProjectPK")
			ProjectID = rs2("ProjectID")
			ProjectName = rs2("ProjectName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Projekt-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("DepartmentID")) = "" Then
		sql = "SELECT * FROM Department WITH (NOLOCK) WHERE DepartmentID = '" & NullCheck(Request("DepartmentID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Abteilungs-ID konnte nicht gefunden werden."
			Call WONew()
		Else
			DepartmentPK = rs2("DepartmentPK")
			DepartmentID = rs2("DepartmentID")
			DepartmentName = rs2("DepartmentName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Abteilungs-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("TenantID")) = "" Then
		sql = "SELECT * FROM Tenant WITH (NOLOCK) WHERE TenantID = '" & NullCheck(Request("TenantID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Kunden-ID konnte nicht gefunden werden."
			Call WONew()
		Else
			TenantPK = rs2("TenantPK")
			TenantID = rs2("TenantID")
			TenantName = rs2("TenantName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Kunden-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("ShopID")) = "" Then
		sql = "SELECT * FROM Shop WITH (NOLOCK) WHERE ShopID = '" & NullCheck(Request("ShopID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Werkstatt-ID konnte nicht gefunden werden."
			Call WONew()
		Else
			ShopPK = rs2("ShopPK")
			ShopID = rs2("ShopID")
			ShopName = rs2("ShopName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r Werkstatt-ID ist ung&#252;ltig."
			Call WONew()
		End If
	End If

	If Not NullCheck(Request("LaborID")) = "" Then
		sql = "SELECT * FROM Labor WITH (NOLOCK) WHERE LaborID = '" & NullCheck(Request("LaborID")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If rs2.eof Then
			HeaderMSG = "Zugeordnete ID konnte nicht gefunden werden."
			Call WONew()
		Else
			LaborPK = rs2("LaborPK")
			LaborID = rs2("LaborID")
			LaborName = rs2("LaborName")
		End If
		If Err.Number <> 0 Then
			HeaderMSG = "Der Wert f&#252;r zugeordnete ID ist ung&#252;ltig."
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
	    HeaderMSG = "Der angegebene Grund ist ung&#252;ltig."
	    Call WONew()
    End If

	rs("Status") = Trim(Mid("REQUESTED",1,15))	' Nullable: No Type: nvarchar
	rs("StatusDesc") = Trim(Mid("Requested",1,50))	' Nullable: YES Type: nvarchar
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Zustand ist ung&#252;ltig."
	    Call WONew()
    End If

	If GetSession("WOAuthReq") = "0" Then
		rs("AuthStatus") = Trim(Mid("NOTREQUIRED",1,15))	' Nullable: No Type: nvarchar
		rs("AuthStatusDesc") = Trim(Mid("(nicht erforderlich)",1,50))	' Nullable: YES Type: nvarchar
	Else
		rs("AuthStatus") = Trim(Mid("REQUIRED" & GetSession("WOAuthReq"),1,15))	' Nullable: No Type: nvarchar
		rs("AuthStatusDesc") = Trim(Mid("Erforderlich - Ebene" & GetSession("WOAuthReq"),1,50))	' Nullable: YES Type: nvarchar
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Auth Zustand ist ung&#252;ltig."
	    Call WONew()
    End If

	rs("StatusDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))	' Nullable: YES Type: datetime
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Zustands-Datum ist ung&#252;ltig."
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
	    HeaderMSG = "Der Wert f&#252;r Anforderer-ID ist ung&#252;ltig."
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
	    HeaderMSG = "Der Wert f&#252;r Ausr&#252;stungs-ID ist ung&#252;ltig."
	    Call WONew()
    End If

    If Request("TargetDate") = "" Then
	    rs("TargetDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))
	Else
	    rs("TargetDate") = SQLdatetimeADO(Request("TargetDate"))
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Zieldatum ist ung&#252;ltig."
	    Call WONew()
    End If

	rs("Type") = Trim(Mid(txtType,1,25))	' Nullable: YES Type: nvarchar
	rs("TypeDesc") = NullCheck(txtTypeDesc)	' Nullable: YES Type: nvarchar
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Typ ist ung&#252;ltig."
	    Call WONew()
    End If

	rs("Priority") = Trim(Mid(txtPriority,1,25))	' Nullable: No Type: nvarchar
	rs("PriorityDesc") = NullCheck(txtPriorityDesc)	' Nullable: YES Type: nvarchar
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Priorit&#228;t ist ung&#252;ltig."
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
	    HeaderMSG = "Der Wert f&#252;r Reparaturzentrums-ID ist ung&#252;ltig."
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
	    HeaderMSG = "Der Wert f&#252;r Werkstatt-ID ist ung&#252;ltig."
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
		HeaderMSG = "Der Wert f&#252;r Leiter-ID ist ung&#252;ltig."
		Call WONew()
	End If

	If Not DepartmentPK = "" Then
		rs("DepartmentPK") = DepartmentPK
		rs("DepartmentID") = DepartmentID
		rs("DepartmentName") = DepartmentName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "Der Wert f&#252;r Abteilungs-ID ist ung&#252;ltig."
		Call WONew()
	End If

	If Not TenantPK = "" Then
		rs("TenantPK") = TenantPK
		rs("TenantID") = TenantID
		rs("TenantName") = TenantName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "Der Wert f&#252;r Kunden-ID ist ung&#252;ltig."
		Call WONew()
	End If

	If Not AccountPK = "" Then
		rs("AccountPK") = AccountPK
		rs("AccountID") = AccountID
		rs("AccountName") = AccountName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "Die angegebene Kostenstelle-ID ist ung&#252;ltig."
		Call WONew()
	End If

	If Not CategoryPK = "" Then
		rs("CategoryPK") = CategoryPK
		rs("CategoryID") = CategoryID
		rs("CategoryName") = CategoryName
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "Der Wert f&#252;r Kategorie-ID ist ung&#252;ltig."
		Call WONew()
	End If

	If CLng(TargetHours) > 0 Then
		rs("TargetHours") = TargetHours
	End If
	If Err.Number <> 0 Then
		HeaderMSG = "Der Wert f&#252;r Ziel-Stunden ist ung&#252;ltig."
		Call WONew()
	End If

	If GetPreference(db,True,RCPreference,"WO_DefaultSurvey",prefvalue, prefdesc, prefpk) Then
		rs("SurveyBox") = 1
		rs("Survey_ID") = prefpk
	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Gutachten ist ung&#252;ltig."
	    Call WONew()
    End If

	rs("Requested") = SQLdatetimeADO(DateTimeNullCheck(Now()))
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Angefordert ist ung&#252;ltig."
	    Call WONew()
    End If

	If Not ProblemPK = "" Then

		rs("ProblemPK") = ProblemPK
		rs("ProblemID") = ProblemID
		rs("ProblemName") = ProblemName

	End If
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Problem-ID ist ung&#252;ltig."
	    Call WONew()
    End If

	If Not ProcedurePK = "" Then

		rs("ProcedurePK") = ProcedurePK
		rs("ProcedureID") = ProcedureID
		rs("ProcedureName") = ProcedureName

    End If
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Verfahrens-ID ist ung&#252;ltig."
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
	    HeaderMSG = "Der Wert f&#252;r Projekt-ID ist ung&#252;ltig."
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
	    HeaderMSG = "Der Wert f&#252;r Abschalten ist ung&#252;ltig."
	    Call WONew()
    End If
    rs("WarrantyBox") = WarrantyBox
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Garantie ist ung&#252;ltig."
	    Call WONew()
    End If
    rs("FollowupWork") = FollowupWork
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Aufgabe weiter bearbeiten ist ung&#252;ltig."
	    Call WONew()
    End If
    rs("LockoutTagoutBox") = LockoutTagoutBox
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Blockierung ist ung&#252;ltig."
	    Call WONew()
    End If
    rs("Chargeable") = ChargeableBox
    If Err.Number <> 0 Then
	    HeaderMSG = "Der Wert f&#252;r Kostenpflichtig ist ung&#252;ltig."
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

Sub WODetails()
	Dim HeaderTitle
	CardTitle = "AA Details"
	HeaderTitle = "AA Details"
	CardCurrent = "WODetails"
	Call WOStatusProcess(HeaderTitle)
	headerString = GetBackButton("history.back();")
	headerString = headerString & GetSearchButton()
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Abmelden' title='Abmelden' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, True)
End Sub

%>