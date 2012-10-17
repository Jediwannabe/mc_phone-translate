<%
'******************************************************************************************************

Dim jsstart,jsend,itemicon,mcCopyright,formjs,aok

Sub SetAppVarsFromDB()

	If ErrorHandler Then
		On Error Resume Next
	End If

	' *************************************************
	' THE PART IS CUSTOM FOR MAPP APPLICATION
	' *************************************************

	Dim pdb

	If Not Application("web_path") = "" and Not UCase(Request.QueryString("reset")) = "Y" Then
		Exit Sub
	Else

		Dim SQLServer, Database, UserID, Password, AppProfileName, INIFile, INIKeyValue
		Dim NamedInstance, DatabasePort, IntegratedSecurity, NamedPipes

		If Application("web_path") = "" or UCase(Request.QueryString("reset")) = "Y" Then
			Application.Lock

				If Request.ServerVariables("SERVER_SOFTWARE") = "WEBSvr/ASP" Then
					Application("IsCompiled") = True
					Application("Web_Server_Port") = Trim(Request.ServerVariables("Server_Port"))
				Else
					Application("IsCompiled") = False
					Application("Web_Server_Port") = ""
				End If

				Application("web_path") = Left(Request.ServerVariables("PATH_INFO"), InStr(2,Request.ServerVariables("PATH_INFO"), "/"))
				If Application("web_path") = "" or Application("IsCompiled") or Application("web_path") = "/ipod/" Then
					Application("web_path") = "/"
				End If

				INIFile = FixMapPath(Server.MapPath(Application("web_path") & "mc.ini"))

				Application("Onsite") = ReadINI(INIFile, "Profile", "Onsite", "0")
				If Application("Onsite") = 1 Then
					Application("AppProfileName") = ReadINI(INIFile, "Profile", "Profile Name", "Onsite-Customer")
				Else
					Application("AppProfileName") = ReadINI(INIFile, "Profile", "Profile Name", "Production")
				End If

				' Get the Registration Database Info
				' =================================================================================================================
				SQLServer = ReadINI(INIFile, "Registration Database", "SQL Server", "(local)")
				NamedInstance = ReadINI(INIFile, "Registration Database", "Named Instance", "")
				DatabasePort = ReadINI(INIFile, "Registration Database", "Port", "1433")
				If Application("IsCompiled") or Application("Onsite") = "1" Then
					Database = ReadINI(INIFile, "Registration Database", "Database Name", "mcRegistrationSA")
				Else
					Database = ReadINI(INIFile, "Registration Database", "Database Name", "mcRegistration")
				End If
				NamedPipes = ReadINI(INIFile, "Registration Database", "Named Pipes", "1")
				IntegratedSecurity = ReadINI(INIFile, "Registration Database", "Integrated Security", "0")
				UserID = ReadINI(INIFile, "Registration Database", "User ID", "")
				Password = ReadINI(INIFile, "Registration Database", "Password", "")

				Application("RegDB") = Database

				If Not NamedInstance = "" Then
					SQLServer = SQLServer & "\" & NamedInstance
				End If

				If IntegratedSecurity = "1" Then
					If NamedPipes = "1" Then
						Application("app_dsn") = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & Database & ";Data Source=" & SQLServer & ";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
					Else
						Application("app_dsn") = "Provider=SQLOLEDB.1;Network Library=DBMSSOCN;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & Database & ";Data Source=" & SQLServer & "," & DatabasePort & ";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
					End If
				Else
					If NamedPipes = "1" Then
						Application("app_dsn") = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & UserID & ";Password=" & Password & ";Initial Catalog=" & Database & ";Data Source=" & SQLServer & ";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
					Else
						Application("app_dsn") = "Provider=SQLOLEDB.1;Network Library=DBMSSOCN;Persist Security Info=True;User ID=" & UserID & ";Password=" & Password & ";Initial Catalog=" & Database & ";Data Source=" & SQLServer & "," & DatabasePort & ";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
					End If
				End If

				' Get the Entity Database Info
				' =================================================================================================================
				If ReadINI(INIFile, "Entity Database", "Entity Database Settings From INI File", "") = "1" Then
					SQLServer = ReadINI(INIFile, "Entity Database", "SQL Server", "(local)")
					NamedInstance = ReadINI(INIFile, "Entity Database", "Named Instance", "")
					DatabasePort = ReadINI(INIFile, "Entity Database", "Port", "1433")
					NamedPipes = ReadINI(INIFile, "Entity Database", "Named Pipes", "1")
					Database = ReadINI(INIFile, "Entity Database", "Database Name", "demoAMI")
					IntegratedSecurity = ReadINI(INIFile, "Entity Database", "Integrated Security", "0")
					UserID = ReadINI(INIFile, "Entity Database", "User ID", "")
					Password = ReadINI(INIFile, "Entity Database", "Password", "")

					Application("ds") = SQLServer
					Application("db") = Database

					If Not NamedInstance = "" Then
						SQLServer = SQLServer & "\" & NamedInstance
					End If

					If IntegratedSecurity = "1" Then
						If NamedPipes = "1" Then
							Application("entity_dsn") = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & Database & ";Data Source=" & SQLServer & ";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
						Else
							Application("entity_dsn") = "Provider=SQLOLEDB.1;Network Library=DBMSSOCN;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & Database & ";Data Source=" & SQLServer & "," & DatabasePort & ";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
						End If
					Else
						If NamedPipes = "1" Then
							Application("entity_dsn") = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & UserID & ";Password=" & Password & ";Initial Catalog=" & Database & ";Data Source=" & SQLServer & ";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
						Else
							Application("entity_dsn") = "Provider=SQLOLEDB.1;Network Library=DBMSSOCN;Persist Security Info=True;User ID=" & UserID & ";Password=" & Password & ";Initial Catalog=" & Database & ";Data Source=" & SQLServer & "," & DatabasePort & ";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
						End If
					End If
				Else
					Application("entity_dsn") = ""
					Application("ds") = ""
					Application("db") = ""
				End If

				' Get Timeouts
				' =================================================================================================================
				Application("app_dsn_ConnectionTimeout") = ReadINI(INIFile, "Timeouts", "Connection Timeout", "60")
				Application("app_dsn_CommandTimeout")= ReadINI(INIFile, "Timeouts", "Command Timeout", "120")
				Application("SessionTimeout")= ReadINI(INIFile, "Timeouts", "Session Timeout", "15")

				' Get File System Object Domain, User ID, and Password
				' =================================================================================================================
				Application("objectdomain") = ReadINI(INIFile, "Object Security", "Domain", "")
				Application("objectusername") = ReadINI(INIFile, "Object Security", "User ID", "")
				Application("objectpassword") = ReadINI(INIFile, "Object Security", "Password", "")

				' Get Security
				' =================================================================================================================
				If ReadINI(INIFile, "Security", "Use SSL", "0") = "1" Then
					Application("Use_SSL") = True
					Application("webHTTP") = "https://"
				Else
					Application("Use_SSL") = False
					Application("webHTTP") = "http://"
				End If

				' Get SMTP Server
				' =================================================================================================================
				Application("SMTP_Server") = ReadINI(INIFile, "SMTP", "SMTP Server", "localhost")
				Application("ReturnMail") = ReadINI(INIFile, "SMTP", "ReturnMail", "returnmail@maintenanceconnection.com")
				Application("FromMail") = ReadINI(INIFile, "SMTP", "FromMail", "agent@maintenanceconnection.com")

				' Get Location
				' =================================================================================================================
				Application("Location") = UCase(ReadINI(INIFile, "Profile", "Location", "Production"))

				' Get MC Virtual Directory (NEW!) --> to get access to images and iconlib for Asset Tree
				' =================================================================================================================
                Application("MCVirtualDirectory") = UCase(ReadINI(INIFile, "Directory Structure", "MCVirtualDirectory", ""))

			Application.UnLock

		End If
		'======================================================================================================================

		Set pdb = New ADOHelper
		pdb.oledbstr = Application("app_dsn")

	End If

	' *************************************************

	Dim RS1,AppProfileDate,UpdateAppVars

	Set RS1 = pdb.RunSPReturnMultiRS("CFG_GetActiveProfileValues",Array(Array("@ProfileNM", adVarChar, adParamInput, 30, Application("AppProfileName"))),"")
	Call dok_check(pdb,"Startup Message","There was a problem retrieving the Configuration Values. The details of the problem are described below. You can try to Log-in again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

	If RS1.Eof Then
		Call CloseObj(RS1)
		Exit Sub
	End If

	AppProfileDate = Trim(RS1.Fields("UpdateDTM"))
	If AppProfileDate = Application("AppProfileDate") and Not UCase(Request.QueryString("reset")) = "Y" Then
		Call CloseObj(RS1)
		Exit Sub
	End If

	Set RS1 = RS1.NextRecordset

	UpdateAppVars = True

	If RS1.Eof Then
		UpdateAppVars = False
		Call CloseObj(RS1)
	End If

	Application.Lock

		Application("AppProfileDate") = AppProfileDate

		If UpdateAppVars Then

			Do Until RS1.Eof

				If Not IsNull(RS1.Fields("Value")) Then

					Select Case UCase(Trim(RS1.Fields("VarType")))

						Case "CHAR"
							'If Not VarType(Application(Trim(RS1.Fields("VarNM")))) = vbEmpty Then
								Application(Trim(RS1.Fields("VarNM"))) = Trim(RS1.Fields("Value"))
							'End If
						Case "INT"
							If UCase(Trim(RS1.Fields("VarNM"))) = "SERVER.SCRIPTTIMEOUT" Then
								Server.ScriptTimeout = CInt(Trim(RS1.Fields("Value")))
								Application("DefaultScriptTimeout") = CInt(Trim(RS1.Fields("Value")))
							Else
								'If Not VarType(Application(Trim(RS1.Fields("VarNM")))) = vbEmpty  Then
									Application(Trim(RS1.Fields("VarNM"))) = CInt(Trim(RS1.Fields("Value")))
								'End If
							End If
						Case "BOOLEAN"
							'If Not VarType(Application(Trim(RS1.Fields("VarNM")))) = vbEmpty  Then
								If UCase(Trim(RS1.Fields("Value"))) = "TRUE" Then
									Application(Trim(RS1.Fields("VarNM"))) = TRUE
								Else
									Application(Trim(RS1.Fields("VarNM"))) = FALSE
								End If
							'End If
						Case "DATE","DATETIME"
							'If Not VarType(Application(Trim(RS1.Fields("VarNM")))) = vbEmpty  Then
								Application(Trim(RS1.Fields("VarNM"))) = CDate(Trim(RS1.Fields("Value")))
							'End If
						Case Else
							'If Not VarType(Application(Trim(RS1.Fields("VarNM")))) = vbEmpty  Then
								Application(Trim(RS1.Fields("VarNM"))) = Trim(RS1.Fields("Value"))
							'End If

					End Select

				End If
				RS1.MoveNext
			Loop
		End If

		If Application("web_server") = "" then
			Application("web_Server") = Trim(LCase(Request.ServerVariables("HTTP_HOST")))
		End If

		Dim pc
		Set pc = Server.CreateObject("Wscript.Network")
		Application("ComputerName") = pc.ComputerName
		Set pc = nothing

		' If the domain is empty then set the domain to the computername
		'If Trim(Application("objectdomain")) = "" Then
		'	Application("objectdomain") = Application("ComputerName")
		'End If

		' This is if running under localhost we can still do multi-user
		' We use replace because the web_server may have a :port on it
		If InStr(UCase(Application("web_Server")),"LOCALHOST") > 0 Then
			Application("web_Server") = Replace(LCase(Application("web_Server")),"localhost",LCase(Application("ComputerName")))
		End If

	Application.UnLock
	Call CloseObj(RS1)

	' *************************************************
	' THE PART IS CUSTOM FOR MAPP APPLICATION
	' *************************************************

	pdb.CloseClientConnection
	Set pdb = Nothing

	' *************************************************

End Sub

'******************************************************************************************************

Function SEZ(s)
	SEZ = Sanitize(s,2,False,True)
End Function

Function SEZQ(s)
	SEZQ = Sanitize(s,2,True,True)
End Function

Function Sanitize(bodystring,level,addquote,wordfilter)

	If Not bodystring = "" then

		Dim i
		Dim instring
		Dim nonchar
		Dim LastRC
		Dim bodybuild
		Dim Length
		Dim nonword
		Dim getword
		Dim r
		Dim sanitizing

		instring = bodystring

		If level > 0 Then
			If level = 1 then
				' Very tight filter
				nonchar = Array("~","`","#","$","%","^","&","*","(",")","_","<","=","+","{","}","[","]","\","|",",",">",".","?","/",":",";")
			ElseIf level = 2 then
				' Very loose filter
				nonchar = Array("[","]",";","--")
			ElseIf level = 3 then
				' Custom
				'**********************************************************************************
				nonchar = Array("#","%","<","=","+",">","@")
			End if

			LastRC = Ubound(nonchar)

			For i = 0 to LastRC
				instring = replace(instring,nonchar(i),"")
			Next

			If addquote Then
				Length = Len(instring)
   				For i = 1 to length
      				bodybuild = bodybuild & Mid(instring, i, 1)
      				If Mid(instring, i, 1) = Chr(39) Then
				 		bodybuild = bodybuild & Mid(instring, i, 1)
      				End If
   				Next
   				instring = bodybuild
			End If
		End If

		If wordfilter Then

			If InStr(instring," ") > 0 Then

   				'****** begin sql word filter
 				nonword = Array("revoke","grant","alter","create","drop","delete","insert","update","exec","execute","truncate")

				getword=Split(instring)
				for i=0 to ubound(nonword)
					For r=0 to ubound(getword)
						if lcase(getword(r)) = lcase(nonword(i)) then
							getword(r)= null
							' to replace the bad word with a set of characters comment the above line and remove the comment from the line below.
							'getword(r)= String(len(getword(r)),"*")
						end if
					next
				next
				instring = ""
				for r=0 to ubound(getword)
					if not isnull(getword(r)) then
						instring = instring & " " & getword(r)
					end if
				next

			End If

		End If

		'****** end word filter
		Sanitize = Trim(instring)

	End if
End Function

'******************************************************************************************************

Sub scattertoarrays(instructions)

	Do Until RS Is Nothing
		If Not RS.State = adStateClosed Then
			If Not RS.EOF then
				execute("rsdata" & CStr(rscounter) & " = RS.GetRows")
				execute("rsdata" & CStr(rscounter) & "recs = True")
			Else
				execute("rsdata" & CStr(rscounter) & "recs = False")
			End If
		Else
			execute("rsdata" & CStr(rscounter) & "recs = False")
			Exit Do
		End If
		rscounter = rscounter + 1
		Set RS = RS.NextRecordset
	Loop

	If retvalcounter = 1 Then
		If InStr(UCase(instructions),"ADVISORY") > 0 Then
			advisory = Trim(cmd.Parameters.Item("@p_advisory").Value)
		End If
		Retval = CInt(cmd.Parameters.Item("@Return").Value)
	Else
		execute ("Retval" & CStr(retvalcounter) & " = CInt(cmd.Parameters.Item(""@Return"").Value)")
	End If
	retvalcounter = retvalcounter + 1

	If instructions = "KEEPALIVE" or instructions = "KEEPALIVE_ADVISORY" Then
		Set con = cmd.ActiveConnection
	End If

	Set cmd.ActiveConnection = Nothing
	Set cmd = Nothing
	Set RS = Nothing

End Sub

'******************************************************************************************************

Sub sp_start(spname,instructions)

	Dim newconnection
	newconnection = False

	Set cmd = Server.CreateObject("ADODB.Command")
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.CursorLocation = adUseClient
	cmd.CommandType = adCmdStoredProc
	cmd.CommandTimeout = 180
	If instructions = "NEWCONNECTION" Then
		Set cmd.ActiveConnection = OpenConnection(Application("app_dsn"))
	Else
		If IsNull( con ) Then
			newconnection = True
		Else
			If IsObject( con ) Then
				If con Is Nothing Then
					newconnection = True
				Else
					If con.State = 0 Then
						newconnection = True
					End If
				End If
			Else
				newconnection = True
			End If
		End If
		If newconnection Then
			Set cmd.ActiveConnection = OpenConnection(Application("app_dsn"))
		Else
			Set cmd.ActiveConnection = con
		End If
	End If
	cmd.CommandText = "dbo.""" & spname & """"
	cmd.Parameters.Append cmd.CreateParameter("@Return", adInteger, adParamReturnValue, 4)

End Sub

'******************************************************************************************************

Function OpenApplicationConnection( ByRef AppConn )

	If ErrorHandler Then
		On Error Resume Next
	End If

	Dim newconnection
	newconnection = False

	If IsNull( AppConn ) Then
		newconnection = True
	Else
		If IsObject( AppConn ) Then
			If AppConn Is Nothing Then
				newconnection = True
			Else
				If AppConn.State = 0 Then
					newconnection = True
				End If
			End If
		Else
			newconnection = True
		End If
	End If

	If newconnection Then
		Set AppConn = OpenConnection(Application("app_dsn"))
	End If

End Function

'******************************************************************************************************

Function OpenConnection(constr)

	On Error Resume Next
	Err.Clear

	Dim MaxCount,Count,MyErrMessage

	MaxCount = 2
	Count = 0

	Do While (Count < MaxCount)
		Err.Clear

		Set OpenConnection = Server.CreateObject("ADODB.Connection")
		OpenConnection.ConnectionTimeout = Application("app_dsn_ConnectionTimeout")
		OpenConnection.CommandTimeout = Application("app_dsn_CommandTimeout")
		OpenConnection.Open constr

		If Err.Number = 0 Then
			Exit Do
		Else
			MyErrMessage = Trim(Err.Description)
		End If

		OpenConnection.Close
		Set OpenConnection = Nothing
		Count = Count + 1
	Loop

	If Count = MaxCount Then
		If InStr(UCase(Request.ServerVariables("path_info")),"/MC_LOGIN.ASP") > 0 Then
		errortext = "Maintenance Connection is not available for logons at this time due to issues with the registration database. We apologize for the inconvenience."
		%>
		<!-- Copyright: © 1998-2002 Maintenance Connection, Inc. All rights reserved. -->

		<html>
		<head>
		<title></title>
		<script language="JavaScript">
		try
		{
			parent.document.getElementById('welcomediv').style.display = 'none';
			parent.document.getElementById('errorspan').innerHTML = '<% =errortext %>';
			parent.document.getElementById('errordiv').style.display = '';
			parent.document.welcome.fld_password.value = '';
			parent.playsound('sounds/error.wav');
		}
		catch(e)
		{
			document.write('<% =errortext %>')
		}
		</script>
		</head>
		<body>
		</body>
		</html>
		<%
		Response.End
		Else
			Call DisplayError("An error has occurred on page " & Trim(Request.ServerVariables("PATH_INFO")) & ".<br><br>" & MyErrMessage,"L","")
		End If
		Set OpenConnection = Nothing
	End If

End Function

'******************************************************************************************************

Function mcval(fieldval)

	Dim BlankSpace,shownull

	BlankSpace="&nbsp;"
	shownull=BlankSpace

	mcval = fieldval

	if isnull(fieldval) then
	   mcval=shownull
	end if
	if trim(fieldval)="" then
	   mcval=BlankSpace
	end if

End Function

'******************************************************************************************************

Function GetWebServer()

	If Application("IsCompiled") Then
		GetWebServer = Trim(LCase(Request.ServerVariables("HTTP_HOST")))
	Else
		GetWebServer = Application("web_Server")
	End If

End Function

'******************************************************************************************************

Sub DoResponse(theaction,newjs,playsound,setfocus,functiononly,keyvalue)

	If Application("OnErrorResumeNext") Then
		On Error Resume Next
	End If

	If Not functiononly Then
		donocacheheader
		CommitSession

		Response.Write mcCopyright

		Response.Write("<html>")
		Response.Write("<head>")
		Response.Write("<meta HTTP-EQUIV=""Expires"" content=""-1"">")
		Response.Write("<title></title>")
		Response.Write("<script language=""JavaScript"">")
		Response.Write("function window_load()")
		Response.Write("{	")
		'Response.Write("if (top.doError) {window.onerror = top.doError;}")
	End If
	    Dim formjs,aok,errorfield
		Response.Write formjs
		%>
			top.endprocess();

			<% If aok Then %>
				<% If playsound Then %>
					top.playsound('sounds/done.wav');
				<% End If %>
			<% Else %>
				<% If playsound Then %>
					top.playsound('sounds/error.wav');
				<% End If %>
			<% End If %>

			top.showactions('<% =Trim(theaction) %>');
			<% =newjs %>

			<% If aok or errorfield = "" Then %>
				<% If setfocus Then %>
					//top.frames['fraTopic'].focus();
					top.dofocus();
				<% End If %>
			<% Else %>
				if (top.document.frames['fraTopic'].<% =errortabinfo %>) {
					top.showtabinfo('<% =errortabinfo %>');
				}
				if (<% =errorfield %>) {
					<% =errorfield %>.focus();
				}
			<% End If %>

			<% Dim returnmessage, returnclass %>
			top.showmessage('<% =returnmessage %>','<% =returnclass %>');
			<% If Not returnmessage = "" Then %>
				if (top.timeoutid != null)
				{
					top.clearTimeout(top.timeoutid);
				}
				<% If returnclass = "errormessage" Then %>
				top.timeoutid = top.setTimeout('top.removemessage()',4000);
				<% Else %>
				top.timeoutid = top.setTimeout('top.removemessage()',1000);
				<% End If %>
			<% End If %>
			<% If aok Then %>
			if (top.addontheflymode == true)
			{
				var myparent = top.dialogArguments.caller;
				//alert(myparent.name);

				// If addontheflypk == '' then we are in add on the fly (not view on the fly)

				//if (top.addontheflypk == '')
				//{
				//	if (myparent.timeoutid != null)
				//	{
				//		myparent.clearTimeout(myparent.timeoutid);
				//	}
				//
				//	myparent.timeoutid = myparent.setTimeout('doeventsafterchildcloses(\'<% =keyvalue %>\',true,\'<% =addontheflymodule %>\')',10);
				//	top.close();
				//	return;
				//}
				//else
				//{
					// addontheflypk is not '' so we are in view on the fly
					//myparent.timeoutid = myparent.setTimeout('doeventsafterchildcloses(\'<% =keyvalue %>\',false,\'<% =addontheflymodule %>\')',10);

					// we are in add on the fly mode since we don't have a PK when we went into on the fly mode
					if (top.addontheflypk == '')
					{
						top.addonthefly_recaddedpk = '<% =keyvalue %>';
					}
					top.addonthefly_recupdated = true;
				//}
			}
			<% End If %>
	<% If Not functiononly Then %>
		}
		window.onload = window_load;
	</script>
	</head>
	<body>
	<%

	FlushItNoStore

	If Application("ASPDEBUG") then
		aspdebug
	End If

	%>
	</body>
	</html>
	<%
	Response.End
	End If

End Sub

'******************************************************************************************************

Sub DoScript(newjs,playsound)

	If Application("OnErrorResumeNext") Then
		On Error Resume Next
	End If

	donocacheheader

	CommitSession
	Response.Write mcCopyright
	%>

	<html>
	<head>
	<meta HTTP-EQUIV="Expires" content="-1">
	<title></title>
	<script language="JavaScript">
	<!--
			window.onerror = top.doError;
			top.endprocess();
			<% If aok Then %>
				<% If playsound Then %>
					top.playsound('sounds/done.wav');
				<% End If %>
			<% Else %>
				<% If playsound Then %>
					top.playsound('sounds/error.wav');
				<% End If %>
			<% End If %>
			<% =newjs %>
	//-->
	</script>
	</head>
	<body>
	<%

	FlushItNoStore

	If Application("ASPDEBUG") then
		aspdebug
	End If

	%>
	</body>
	</html>
	<%
	Response.End
End Sub

'******************************************************************************************************

Sub GlobalInit()

	addontheflymode = "N"
	addontheflymodule = ""

	duprecord = False
	norecord = False
	firstload = False
	treefiltervalue = ""

	If Request.Form("asubmit") = "" Then
		keyvalue = Trim(Request.QueryString("kv"))

		If keyvalue = "" Then
			' When first loading a module - there will not be a key value
			firstload = True
			norecord = True
		Else
			mcmode = UCase(Request.QueryString("mcmode"))
			If keyvalue = "NEW" Then
				keyvalue = ""
				newrecord = True
			ElseIf Mid(keyvalue,1,2)="!D" Then
				'When coming on the QueryString as Duplicate set these values accordingly
				newrecord = True
				duprecord = True
				keyvalue = Mid(keyvalue,3)
			ElseIf keyvalue = "TOP" Then
				'Will need to change this to point to the top record!
				keyvalue = ""
				newrecord = False
				'norecord = True
			Else
				newrecord = False
			End If

		End If
	Else
		addontheflymode = Trim(UCase(Request.Form("addontheflymode")))
		addontheflymodule = Trim(UCase(Request.Form("addontheflymodule")))
		keyvalue = Trim(Request.Form("kv"))
		'If Not keyvalue = "" Then
			If keyvalue = "NONE" Then
				norecord = True
			End If
			mcmode = UCase(Request.Form("mcmode"))
			If mcmode = "ADD" Then
				' Must be ADD mode
				newrecord = True
			ElseIf mcmode = "DUPLICATE" Then
				'When coming from a Submit as Duplicate set these values accordingly
				newrecord = False
				duprecord = True
			Else
				' Must be EDIT or DUPLICATE mode
				newrecord = False
			End If
		' This does not work below because keyvalue = "" for new records!
		'Else
		'	' List Submit
		'	keyvalues = Trim(Request.Form("mckeyvalues"))
		'End If
	End If

	curtab = Trim(Request("curtab"))
	currentmodule = Trim(Request("currentmodule"))

End Sub

'******************************************************************************************************

Function setuptabs(numoftabs,tabno,tabtext,tabdiv,isduprecord)
	setuptabs = setuptabs + "top.settabtext("+CStr(numoftabs)+",'"+CStr(tabno)+"','" + tabtext + "','" + tabdiv + "');" + nl
End Function

'******************************************************************************************************

Function setuptoctabs(numoftabs,tabno,tabtext,tabdiv)
	setuptoctabs = "top.settabtexttoc("+CStr(numoftabs)+",'"+CStr(tabno)+"','" + tabtext + "','" + tabdiv + "');" + nl
End Function


'******************************************************************************************************

Sub dofilldataheader
Response.Write mcCopyright
%>
<html>
<head>
<meta HTTP-EQUIV="Expires" content="-1">
<meta http-equiv="Cache-Control" content="no-cache">
<%
End Sub

'******************************************************************************************************

Sub dofilldatafooter
%>
</head>
<body>

<%
If Application("ASPDEBUG") then
	aspdebug
End If
%>

</body>
</html>
<%
End Sub

'******************************************************************************************************

Sub DoMCTop(bodyparams,scroll,headercmds)

	Response.Write mcCopyright
	Response.Write nl
	Response.Write "<html>" + nl
	Response.Write "<head>" + nl
	Response.Write "<title></title>" + nl
	If Not headercmds = "" Then
		Response.Write headercmds
	End If
	Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../../css/mc_css.css"">" + nl
	%>
	<script language="javascript">
	function doeventsafterchildcloses(recordaddedpk,dofilter,frommodule)
	{
		try
		{
			// this get's called after a child window closes
			if (dofilter == false && (frommodule = 'AS' || frommodule == 'CL'))
			{
				if (frommodule == 'AS')
				{
					top.refreshcurrentexplorer(false,true);
				}
				else
				{
					if (moduleforexplorer == 'CL')
					{
						top.refreshcurrentexplorer(false);
					}
				}
			}
			top.playsound('sounds/done.wav');
			if (top.dirtydata == false)
			{
				top.refreshcurrentrecord();
			}
		}
		catch(e)
		{}
	}
	</script>
	<%
	Response.Write "</head>" + nl
	If Application("SCRIPTENCODE") Then
		Response.Write "<body " + bodyparams + " style=""scrollbar-base-color: #EAEAEA; border-left:#808080 5px solid; background-position: center;background-repeat: no-repeat;background-attachment: fixed;  background-image: url('../../images/mclogo_watermark3.jpg');"" bgColor = ""#FBFBFB"" topmargin=""0"" leftmargin=""0"" oncontextmenu=""return false;"" onkeypress=""return top.checkKey();"">" + nl
	Else
		Response.Write "<body " + bodyparams + " style=""scrollbar-base-color: #EAEAEA; border-left:#808080 5px solid; background-position: center;background-repeat: no-repeat;background-attachment: fixed;  background-image: url('../../images/mclogo_watermark3.jpg');"" bgColor = ""#FBFBFB"" topmargin=""0"" leftmargin=""0"" onkeypress=""return top.checkKey();"">" + nl
	End If
	If Not scroll Then
	%>
	<script language="javascript">
	document.body.scroll = 'no';
	</script>
	<%
	End If

End Sub

Sub domcstartform()

	Response.Write "<form name=""mcform"" method=""post"" target=""fraLoad"" onReset=""return top.CheckReset(self);"" AutoComplete=""OFF"">" + nl

	Response.Write "<input type=""hidden"" name=""asubmit"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""lastaction"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""mcmode"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""kv"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""curtab"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""txtRowVersionUserPK"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""txtRowVersionInitials"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""addontheflymode"" class=""mckeep"" value=""n"">" + nl
	Response.Write "<input type=""hidden"" name=""addontheflymodule"" class=""mckeep"" value="""">" + nl
	Response.Write "<input type=""hidden"" name=""currentmodule"" class=""mckeep"" value="""">" + nl
	Response.Write "<input type=""hidden"" name=""noexplorerrefresh"" class=""mckeep"" value=""N"">" + nl

End Sub

Sub domcstartlistform()

	Response.Write "<form name=""mcform"" method=""post"" onReset=""return top.CheckReset(self);"" AutoComplete=""OFF"">" + nl

	Response.Write "<input type=""hidden"" name=""asubmit"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""lastaction"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""doaction"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""mcmode"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""kv"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""currentmodule"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""sallpages"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""curtab"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""txtRowVersionUserPK"" class=""mckeep"" value>" + nl
	Response.Write "<input type=""hidden"" name=""txtRowVersionInitials"" class=""mckeep"" value>" + nl

End Sub
'******************************************************************************************************

Sub DoMCBottom(windowonloadjs)

	Response.Write jsstart
	Response.Write "function window_load()" + nl
	Response.Write "{" + nl
	Response.Write windowonloadjs + nl
	Response.Write "document.onclick = top.doc_click;" + nl
	Response.Write "document.ondblclick = top.doc_dblclick;" + nl
	Response.Write "document.onkeydown = top.doc_keydown;" + nl
	Response.Write "window.onerror = top.doError;" + nl
	Response.Write "top.dofocus();" + nl
	Response.Write "}" + nl + nl
	Response.Write "window.onload = window_load;" + nl
	Response.Write jsend

	DoMCDivs
	DoIFrame

	If Application("ASPDEBUG") then
		aspdebug
	End If

End Sub

'******************************************************************************************************

Sub DoIFrame()
	If Application("ASPDEBUG") then %>
	<iframe HEIGHT="200" WIDTH="200" STYLE="position:absolute;z-index=100" ID="frapopupouter" NAME="frapopupouter" MARGINHEIGHT="0" MARGINWIDTH="0" NORESIZE FRAMEBORDER="0" SCROLLING="yes" allowTransparency="true" SRC="<% =Application("web_path") & Application("mapp_path") & "mc_sslfix.htm" %>"></iframe>
	<% Else %>
	<iframe HEIGHT="0" WIDTH="0" STYLE="display:none;position:absolute;z-index=100" ID="frapopupouter" NAME="frapopupouter" MARGINHEIGHT="0" MARGINWIDTH="0" NORESIZE FRAMEBORDER="0" SCROLLING="yes" allowTransparency="true" SRC="<% =Application("web_path") & Application("mapp_path") & "mc_sslfix.htm" %>"></iframe>
	<% End If
End Sub

'******************************************************************************************************

Sub DoMCDivs()
%>
<div id="FKeditmenu" onclick="top.clickMenu('FKeditmenu',top.fraTopic)" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(top.fraTopic)" oncontextmenu="return false;" style="z-index:100;position:absolute;display:none; background-Color:#FEFEFE; height:45px; width:175px; border: 2 outset #FFFFFF; padding-top:5; padding-bottom:5;overflow-y:none;background-image: url('../../images/menuleftbg5.gif');">
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/openmc2_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenFKNormal">Open</td></tr></table>
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/openinnewwindow_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenFK">Open in New Window</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/findxp_g.gif" style="margin-left:3;margin-right:9;" WIDTH="15" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenLookupForFK">Lookup</td></tr></table>
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/exploreit_g.gif" style="margin-left:3;margin-right:9;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenExplorerForFK">Explore (Double-Click Field)</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>
</div>

<div id="FKFLYeditmenu" onclick="top.clickMenu('FKFLYeditmenu',top.fraTopic)" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(top.fraTopic)" oncontextmenu="return false;" style="z-index:100;position:absolute;display:none; background-Color:#FEFEFE; height:45px; width:175px; border: 2 outset #FFFFFF; padding-top:5; padding-bottom:5;overflow-y:none;background-image: url('../../images/menuleftbg5.gif');">
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/openinnewwindow_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenFK">Open in New Window</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/findxp_g.gif" style="margin-left:3;margin-right:9;" WIDTH="15" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenLookupForFK">Lookup</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>
</div>

<div id="FKNVeditmenu" onclick="top.clickMenu('FKNVeditmenu',top.fraTopic)" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(top.fraTopic)" oncontextmenu="return false;" style="z-index:100;position:absolute;display:none; background-Color:#FEFEFE; height:45px; width:175px; border: 2 outset #FFFFFF; padding-top:5; padding-bottom:5;overflow-y:none;background-image: url('../../images/menuleftbg5.gif');">
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/findxp_g.gif" style="margin-left:3;margin-right:9;" WIDTH="15" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenLookupForFK">Lookup</td></tr></table>
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/exploreit_g.gif" style="margin-left:3;margin-right:9;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenExplorerForFK">Explore (Double-Click Field)</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>
</div>

<div id="FKNVFLYeditmenu" onclick="top.clickMenu('FKNVeditmenu',top.fraTopic)" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(top.fraTopic)" oncontextmenu="return false;" style="z-index:100;position:absolute;display:none; background-Color:#FEFEFE; height:45px; width:175px; border: 2 outset #FFFFFF; padding-top:5; padding-bottom:5;overflow-y:none;background-image: url('../../images/menuleftbg5.gif');">
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/findxp_g.gif" style="margin-left:3;margin-right:9;" WIDTH="15" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenLookupForFK">Lookup</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>
</div>

<div id="loadingdiv" style="position:absolute;display:none; background-Color:white; border:1px solid grey; border-top:2px solid #cccccc; border-left:2px solid #cccccc; height:15px; width:20px;z-index:100;">
	<table bgcolor="#FFFFFF" border="0" cellpadding="0" cellspacing="0" width="100%" height="100%"><tr><td width="100%" align="center"><font size="1" face="Arial"><center>Loading...Please Wait</center></font></td></tr></table>
</div>

<div id="lookupsavecancel" STYLE="display: none; position:absolute; z-index:100;">
	<table bgcolor="#FFFFCC" border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
	<td width="100%" align="center" height="25" valign="bottom" background="../../images/lookuptopbg.jpg">
	<img border="0" style="cursor:hand;" src="../../images/button_save.gif" width="80" height="15" onclick="document.frapopupouter.submitform('SAVE','Saving');event.cancelBubble = true;">&nbsp;<img style="cursor:hand;" border="0" src="../../images/button_cancel.gif" width="80" height="15" onclick="document.frapopupouter.cancelform();event.cancelBubble = true;">
	</td>
	</tr>
	</table>
</div>

<div id="lookuppopupupper" STYLE="display: none; position:absolute; z-index:100;">
	<table border="0" width="100%" height="20" bgcolor="#FFFFCC" CELLSPACING="0" CELLPADDING="0">
	<tr>
	  <th width="100%" NOWRAP align="left">
	    <table border="0" cellpadding="0" cellspacing="0" width="100%">
	      <tr>
	        <td>
	        <table CELLSPACING="0" CELLPADDING="0" border="0"><tr><td><img border="0" src="../../images/red-arrow.gif" width="8" height="12" ondblclick="document.frapopupouter.dolookupedit(self);event.cancelBubble = true;"></td><td><font style="font-family: Arial; font-size: 8pt; font-weight: bold" color="#000000">&nbsp;<span id="lookuppopupuppertext">Lookup Table</span></font></td></tr></table></td>
	        <td align="right"><font style="font-family: Arial; font-size: 7pt; font-weight: bold; color=#FFFFFF" color="#FFFFFF"><img id="lookuptableeditbutton" border="0" style="cursor:hand;" src="../../images/popupedit.gif" width="26" height="14" onclick="document.frapopupouter.dolookupedit(self);event.cancelBubble = true;"><img border="0" style="cursor:hand;" src="../../images/closepopup.gif" hspace="1" width="15" height="14" onclick="top.clearallpopups(self);event.cancelBubble = true;" title="Close"></font></td>
	      </tr>
	    </table>
	  </th>
	</tr>
	</table>
</div>

<% If False Then %>
<div id="caption_l" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 267; height: 100; padding-left: 28; padding-right: 12; padding-top: 10; background-image: url('../../images/caption_l.gif'); z-index:100;"></div>
<div id="caption_r" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 267; height: 100; padding-left: 12; padding-right: 28; padding-top: 10; background-image: url('../../images/caption_r.gif'); z-index:100;"></div>
<div id="caption_tl" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 28; background-image: url('../../images/caption_tl.gif'); z-index:100;"></div>
<div id="caption_tr" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 28; background-image: url('../../images/caption_tr.gif'); z-index:100;"></div>
<div id="caption_tr_2" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 28; background-image: url('../../images/caption_tr2.gif'); z-index:100;"></div>
<div id="caption_bl" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 10; background-image: url('../../images/caption_bl.gif'); z-index:100;"></div>
<div id="caption_br" onclick="this.style.display='none'" style="cursor:hand; position:absolute; display:none; font-family: arial; font-size: 12; width: 251; height: 115; padding-left: 12; padding-right: 12; padding-top: 10; background-image: url('../../images/caption_br.gif'); z-index:100;"></div>
<% End If %>

<div id="roweditmenu" onclick="top.clickMenu('roweditmenu',top.fraTopic)" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(self)" oncontextmenu="return false;" style="z-index:100;position:absolute;display:none; background-Color:#FEFEFE; height:45px; width:150px; border: 2 outset #FFFFFF; padding-top:5; padding-bottom:5;overflow-y:none;background-image: url('../../images/menuleftbg5.gif');">
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/openmc2_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenItemFromRow">Open</td></tr></table>
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/openinnewwindow_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenItemFromRowNewWindow">Open in New Window</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/mnu_delete_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuRemoveRow">Remove</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>
</div>

<div id="roweditmenu_nodel" onclick="top.clickMenu('roweditmenu_nodel',top.fraTopic)" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(self)" oncontextmenu="return false;" style="z-index:100;position:absolute;display:none; background-Color:#FEFEFE; height:45px; width:150px; border: 2 outset #FFFFFF; padding-top:5; padding-bottom:5;overflow-y:none;background-image: url('../../images/menuleftbg5.gif');">
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/openmc2_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenItemFromRow">Open</td></tr></table>
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/openinnewwindow_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuOpenItemFromRowNewWindow">Open in New Window</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>
</div>

<div id="roweditmenu_nojump" onclick="top.clickMenu('roweditmenu_nojump',top.fraTopic)" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(self)" oncontextmenu="return false;" style="z-index:100;position:absolute;display:none; background-Color:#FEFEFE; height:45px; width:150px; border: 2 outset #FFFFFF; padding-top:5; padding-bottom:5;overflow-y:none;background-image: url('../../images/menuleftbg5.gif');">
  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/mnu_delete_g.gif" style="margin-left:3;margin-right:9;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuRemoveRow">Remove</td></tr></table>
  <div class="menuItemHR" id="mnuLine"><hr size="1" class="menuhr"></div>
  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>
</div>

<%
End Sub

Sub DoHiddenFieldsDiv
%>
<div id="hiddenfields">
</div>
<%
End Sub

Function domodelogic(justdata)

		domodelogic = True

		Response.Write("")+nl
		If newrecord Then

			If duprecord Then
				Response.Write("top.mcmode = 'DUPLICATE';")+nl
				If Not writerecorddata(null) Then
					Response.Write("top.endprocess();") + nl
					'Response.Write("top.wocostviewsetting = -1;") + nl
					Response.Write("myform.noexplorerrefresh.value == 'N';") + nl
					domodelogic = False
					Exit Function
				End If
				Call SetupFields("DUPLICATE",justdata)
			Else
				Response.Write("top.mcmode = 'ADD';")+nl
				Call SetupFields("NEW",justdata)
			End If

			Call EnableDisableFields("NEW")

		Else

			Response.Write("top.mcmode = 'EDIT';")+nl
			If (justdata) Then
				If Not writerecorddata(null) Then
					Response.Write("top.endprocess();") + nl
					'Response.Write("top.wocostviewsetting = -1;") + nl
					Response.Write("myform.noexplorerrefresh.value == 'N';") + nl
					domodelogic = False
					Exit Function
				End If
				Call SetupFields("EDIT",justdata)
				Call EnableDisableFields("EDIT")
			Else
				Call SetupFields("EDIT",justdata)
			End If

		End If

		Response.Write("// Ending Process")+nl
		Response.Write("// -------------------------------------------------------------------------")+nl
		Response.Write("top.checkpendingmsg();") + nl
		Response.Write("top.endprocess();") + nl
		'Response.Write("top.wocostviewsetting = -1;") + nl
		Response.Write("myform.noexplorerrefresh.value == 'N';") + nl
		'Response.Write("if (top.addontheflymode == true) {top.isdirty();}") + nl

		' Do not put gosplitview here because it is set on new record by calling isdirty
		'Response.Write("top.gosplitview();") + nl

End Function

'******************************************************************************************************

Sub donocacheheader

	Response.ExpiresAbsolute=#July 7,1998 12:00:00#
	If Application("IEPragma") Then
		Response.AddHeader "Pragma", "no-cache"
	End If

	Response.AddHeader "cache-control","no-store"

	' PipeBoost Directives
	' =============================================================================
	' Turn off Caching
	' =============================================================================
	 Response.AddHeader "X-Pb-CacheOn", "0"

End Sub

'******************************************************************************************************

Sub checkforjustdata()

	If Application("OnErrorResumeNext") Then
		On Error Resume Next
	End If

	If Request.QueryString("justdata") = "" Then
		Exit Sub
	Else
		'we only need the data so do NOT cache this page

		donocacheheader
		CommitSession

		Call filldata(true)
		Response.End
	End If

End Sub

'******************************************************************************************************

Sub filldata(justdata)

	If justdata Then
		dofilldataheader
	End If

	Response.Write jsstart
	Response.Write "window.onerror = top.doError;" + nl
	Response.Write formjs

	If domodelogic(justdata) Then
		' Only do this on initial load for now
		Call Setup(justdata)
	End If

	Response.Write jsend

	If justdata Then
		dofilldatafooter
	End If

End Sub

'******************************************************************************************************

Function ErrorHandler()
	ErrorHandler = Application("OnErrorResumeNext")
End Function

'******************************************************************************************************

Sub FlushIt()
	If Application("IsCompiled") Then
		Exit Sub
	End If
	Response.Flush
End Sub

'******************************************************************************************************

Sub FlushItNoStore()
	' There are problems with the browser cacheing pages
	' with the no-store header + response.flushes.

	' It will cache a .ASP file or a .HTM file if URL
	' parameters are used
	If Application("IsCompiled") or Application("UsePipeBoost") Then
		Exit Sub
	Else
		Call FlushIt()
	End If
End Sub

'******************************************************************************************************

Function BFDecrypt(str)
	BFDecrypt = ""
	If Len(str) > 0 and Application("UseSessionEncryption") Then
		Dim mccrypt
		Set mccrypt = Server.CreateObject("mccrypt.cryptor")
		'Response.Write str
		'Response.End
		On Error Resume Next
		BFDecrypt = Trim(mccrypt.strDecrypt(str))
		Set mccrypt = Nothing
	Else
		BFDecrypt = str
	End If
End Function

'******************************************************************************************************

Function BFEncrypt(str)
	BFEncrypt = ""
	If Application("UseSessionEncryption") Then
		Dim mccrypt
		Set mccrypt = Server.CreateObject("mccrypt.cryptor")
		On Error Resume Next
		BFEncrypt = mccrypt.strEncrypt(str)
		Set mccrypt = Nothing
	Else
		BFEncrypt = str
	End If
End Function

'******************************************************************************************************

Sub ReadSession()
	If Not Request.QueryString("s") = "" Then
		SessionVars = Trim(Request.QueryString("s"))
	Else
		SessionVars = Request.Cookies("mcmain")
	End If
	'SessionVars = "34B3B746866800134E561993B07F299307129E4E39621814911A32B87E4B077CF008E7FB633651DAFBE1B645CF7722CBCB2B663F303026AB2571908C9567676AEFA40E2EF6244BA651B0FC47B21D8A483E5173280DBDACEEBEC08AE96920699D1717521F2AB8B02C26EDD9FEDB9BB820037851F1EE145C4A72778352A349671F9A679F43B741C1B2FEAE67635ABBBD13222D4C9C89D3CE819445C41852367876AD183BF9F2403A2AA0F1CF67D7D9DE51C36BCF2DE648BA34709122242D9467B27C52696EFF7C7E2169A2F48CC77DCF6A6487414E9CAC747AD3C7186562AA9B50A7E78101BD22369588F631D4D50C125FA6622AAA7D0A3DB686ED5E09E6403D294C4E777B1F8093B8510593E63387CE7FC495FC417897EFF6B351D7027ECB50271A0C71F951C75DCB266CC48B52C5E3F71B14B1D09D7790266094A1D2EC0E4C0D7C9856BCAACBC8F5DB0FE3E93716DF232B32BF5116A524D82D1079681B47CC388CA6F8342AE093543948D758CC636DE2D035E74989C5FB69BEAC69F37CA0A55779F2C9A5F2B7B17E1D7C5A815463B0EEAB65808DB6B8E5758A1DE6B6EA7A82072778BEF3C6D2EC07E6B3D5FEE8AAE78E39817921E0CD491140B8964B299F3EC22DE5B613F59CEB7D66E90D0ECACEE29A7B34AA096319FBACD2F379D11E416E4A1987B980B4ECBC9999FBA4B7469E3671"

	If Not SessionVars = "" Then
		SessionVars = BFDecrypt(SessionVars)
	End If

	If Request.QueryString("FromAgent") = "Y" and GetSession("FromAgent") = "Y" Then
		Exit Sub
	End If

	SessionID = GetSession("SessionID")

	If Instr(UCase(Request.ServerVariables("PATH_INFO")),"/SURVEY/") > 0 Then
		Exit Sub
	End If

	If Not ValidateAndUpdateSessionTimeStampSQL(SessionID) Then
		' Put Code Here for Each Application
		If GetSession("webHTTP") = "" Then
			Response.Redirect("http://" & GetWebServer() & Application("web_path") & Application("mapp_path") & "session_timeout.htm")
		Else
			Response.Redirect(GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "session_timeout.htm")
		End If
	End If
End Sub

'******************************************************************************************************

Sub ReadSessionNoCheck()
	If Not Request.QueryString("s") = "" Then
		SessionVars = Trim(Request.QueryString("s"))
	Else
		SessionVars = Request.Cookies("mcmain")
	End If
	If Not SessionVars = "" Then
		SessionVars = BFDecrypt(SessionVars)
	End If
	SessionID = GetSession("SessionID")
End Sub

'******************************************************************************************************

Sub CommitSession()
	If (UpdateVars) and (Not SessionVars = "") Then
		UpdateVars = False
		' Attach current time so that encryption will look different upon each request
		Call SetSession("tm",Now)
		Response.Cookies("mcmain") = BFEncrypt(SessionVars)
		Response.Cookies("mcmain").Path = "/"
	End If
End Sub

'******************************************************************************************************

Sub CheckIfSessionIsValid
	Dim ok
	ok = True

	' Do not use GetSession("IsAdmin") here because IsAdmin is not initialize until after this call on default.asp
	If GetSession("UT") = "MC" Then
		Exit Sub
	End If

	'Response.Write CDate(GetSession("TM"))
	'Response.Write "<br>"
	'Response.Write Now
	'Response.Write "<br>"
	'Response.Write DateDiff("s", CDate(GetSession("TM")), Now)
	'Response.End
	If Not GetSession("TM") = "" Then
		If Not Application("location") = "DEVELOPMENT" and Not Application("location") = "QA" Then
			If Not DateDiff("s", CDate(GetSession("TM")), Now) <= (GetSession("tf") * 60) Then
			'	ok = False
			End If
		End If
	Else
		ok = False
	End If
	If Not ok Then
		If GetSession("webHTTP") = "" Then
			Response.Redirect("http://" & GetWebServer() & Application("web_path") & Application("mapp_path") & "session_timeout.htm")
		Else
			Response.Redirect(GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "session_timeout.htm")
		End If
		If False Then
			Response.Clear
			Response.AddHeader "cache-control", "no-store"
			%>
			<script language="JavaScript">
				top.mcwarn('Session Expired','Your login session has expired. In order to keep a high level of security when using the Maintenance Connection, you will need to login again.');
				top.close();
			</script>
			<%
			Response.End
		End If
	End If
End Sub

'******************************************************************************************************

Function ValidateAndUpdateSessionTimeStampSQL(ByVal p_sessionKey)

	Dim pdb, OutArray

    ValidateAndUpdateSessionTimeStampSQL = False

	Set pdb = New ADOHelper
	pdb.oledbstr = Application("app_dsn")

    Call pdb.RunSP("SES_ValidateSession", _
    Array(_
    Array("RETURN_VALUE", adInteger, adParamReturnValue, 4, 0), _
    Array("@p_sessionKey", adVarChar, adParamInput, 100, p_sessionKey) _
    ),OutArray)

	'Response.Write Replace(Replace(GetSession(""),"^^","<br>"),"^","")
	'Response.End

	If pdb.dok Then

		'Response.Write p_sessionKey
		'Response.End

		If OutArray(0) = 0 Then
			ValidateAndUpdateSessionTimeStampSQL = True
        Else
			'errortext = "There was a problem validating the Licensed Session. Please try again."
        End If

		pdb.CloseClientConnection
		Set pdb = Nothing

    Else
		'errortext = pdb.derror

		'Response.Write pdb.derror
		'Response.End

    End If

	' Do not use GetSession("IsAdmin") here because IsAdmin is not initialize until after this call on default.asp
	If GetSession("UT") = "MC" Then
		ValidateAndUpdateSessionTimeStampSQL = True
	End If

End Function

'******************************************************************************************************

Function GetSession(ByVal varname)
	If ErrorHandler Then
		On Error Resume Next
	End If
	If varname = "" Then
		GetSession = SessionVars
	Else
		GetSession = Extract(SessionVars, "^" + UCase(varname) + "=", "^", "", True)
	End If
End Function

'******************************************************************************************************

Function SetSession(v1a,v1b)
	If ErrorHandler Then
		On Error Resume Next
	End If
	SessionVars = ExtractAdd(SessionVars, v1a, v1b, "^")
	UpdateVars = True
End Function

'******************************************************************************************************

Function ExtractRemove(ByVal lcString, ByVal lcVar, ByVal lcDelim)
	If ErrorHandler Then
		On Error Resume Next
	End If
	ExtractRemove = Replace(lcString, lcDelim & UCase(lcVar) & "=" & Extract(lcString, lcDelim + UCase(lcVar) + "=", lcDelim, "", True) & lcDelim, "")
End Function

'******************************************************************************************************

Function ExtractAppend(ByVal lcString, ByVal lcVar, ByVal lcVarValue, ByVal lcDelim)
	If ErrorHandler Then
		On Error Resume Next
	End If
	ExtractAppend = lcString & lcDelim & UCase(lcVar) & "=" & lcVarValue & lcDelim
End Function

'******************************************************************************************************

Function ExtractAdd(ByVal lcString, ByVal lcVar, ByVal lcVarValue, ByVal lcDelim)
	If ErrorHandler Then
		On Error Resume Next
	End If
	Dim lcExtract
    lcExtract = Extract(lcString, lcDelim + UCase(lcVar) + "=", lcDelim, "", True)

	If IsNull(lcVarValue) Then
		lcVarValue = ""
	End If

    If Not lcExtract = "" Then
       If lcVarValue = "" Then
         ExtractAdd = ExtractRemove(lcString, lcVar, lcDelim)
       Else
         ExtractAdd = Replace(lcString, lcDelim & UCase(lcVar) & "=" & lcExtract & lcDelim, lcDelim & UCase(lcVar) & "=" & lcVarValue & lcDelim)
       End If
    Else
       If Not lcVarValue = "" Then
          ExtractAdd = ExtractAppend(lcString,lcVar,lcVarValue,lcDelim)
       Else
		  ExtractAdd = lcString
       End If
    End If

End Function

'******************************************************************************************************

Function Extract(ByVal lcString,ByVal lcDelim1,ByVal lcDelim2,ByVal lcDelim3,ByVal llEndOk)

    Dim x, lnLocation, lcRetVal, lcChar, lcNewString, lnEnd

    lnLocation = InStr(lcString, lcDelim1)
    If lnLocation = 0 Then
       Extract = ""
       Exit Function
    End If

    lnLocation = lnLocation + Len(lcDelim1)

    lcNewString = Mid(lcString, lnLocation)

    lnEnd = InStr(lcNewString, lcDelim2) - 1
    If lnEnd > 0 Then
       Extract = Mid(lcNewString, 1, lnEnd)
       Exit Function
    End If
    If lnEnd = 0 Then
       Extract = ""
       Exit Function
    End If

    lnEnd = InStr(lcNewString, lcDelim3) - 1
    If lnEnd > 0 Then
       Extract = Mid(lcNewString, 1, lnEnd)
       Exit Function
    End If

    If llEndOk Then
      Extract = Mid(lcNewString, 1)
      Exit Function
    End If

    Extract = ""

End Function

'******************************************************************************************************

Function GetPageName()
	If ErrorHandler Then
		On Error Resume Next
	End If
	GetPageName=UCase(Trim(Mid(Request.ServerVariables("path_info"),InStrRev(Request.ServerVariables("path_info"),"/")+1)))
End Function

'******************************************************************************************************

Function RandomString(length)
	If ErrorHandler Then
		On Error Resume Next
	End If
	Dim i
	RandomString = ""
	Randomize
	For i = 0 to length - 1
		RandomString = RandomString & Chr(Int(26 * Rnd + 65))
	Next
End Function

Function GetRandomNumber(high, low)
	Randomize
    GetRandomNumber = Int((high - low + 1) * Rnd + low)
End Function

'******************************************************************************************************

Function PadString(sString, iLen, sPad)
	If ErrorHandler Then
		On Error Resume Next
	End If
	Dim sTemp, i
	If sPad = "" then sPad = "0"
	sTemp = sString
	For i = 1 to iLen - Len(sString)
		sTemp = sPad & sTemp
	Next
	PadString = sTemp
End Function

'******************************************************************************************************

Function SQLEncode(ByVal str)

	If ErrorHandler Then
		On Error Resume Next
	End If

    Dim vntPosition

    If IsNull(str) Or str = "" Then

    Else
        vntPosition = InStr(str, "'")
        Do While vntPosition <> 0
            str = Mid(str, 1, vntPosition - 1) & "''" & Mid(str, vntPosition + 1)
            vntPosition = InStr(vntPosition + 2, str, "'")
        Loop
    End If
    SQLEncode = Trim(str)

End Function

'******************************************************************************************************

Function JSEncode(ByVal str)

	'If ErrorHandler Then
		On Error Resume Next
	'End If

    Dim vntPosition

    If IsNull(str) Or str = "" Then
		str = ""
    Else
		str = Trim(str)

        'vntPosition = InStr(str, "\")
        'Do While vntPosition <> 0
        '    str = Mid(str, 1, vntPosition - 1) & "\\" & Mid(str, vntPosition + 1)
        '    vntPosition = InStr(vntPosition + 2, str, "\")
        'Loop
        'vntPosition = InStr(str, "'")
        'Do While vntPosition <> 0
        '    str = Mid(str, 1, vntPosition - 1) & "\'" & Mid(str, vntPosition + 1)
        '    vntPosition = InStr(vntPosition + 2, str, "'")
        'Loop

		str = Replace(str, "\", "\\")
		str = Replace(str, Chr(39), "\'")
		str = Replace(str, Chr(34), "\" & Chr(34))
		str = Replace(str, Chr(13), "\r")
		str = Replace(str, Chr(10), "\n")

    End If

	If Err.Number <> 0 Then
		str = ""
	End If

	JSEncode = str

End Function

'******************************************************************************************************

Function DateNullCheck(ByVal str)

	'If ErrorHandler Then
		On Error Resume Next
	'End If

    If IsNull(str) Or str = "" Then
		str=""
    Else
		str = DateValue(CStr(str))
    End If

	If Err.Number <> 0 Then
		str = ""
	End If

    DateNullCheck = str

End Function

'******************************************************************************************************

Function DateTimeNullCheck(ByVal str)

	'If ErrorHandler Then
		On Error Resume Next
	'End If

    If IsNull(str) Or str = "" Then
		str=""
    Else
		str = FormatDateTime(CStr(str))
    End If

	If Err.Number <> 0 Then
		str = ""
	End If

    DateTimeNullCheck = str

End Function

'******************************************************************************************************

Function TimeNullCheck(ByVal str)

	'If ErrorHandler Then
		On Error Resume Next
	'End If

    If IsNull(str) Or str = "" Then
		str=""
    Else
		str = TimeValue(CStr(str))
    End If

	If Err.Number <> 0 Then
		str = ""
	Else
		str = FixTime(str)
	End If

    TimeNullCheck = str

End Function

'******************************************************************************************************

Function FixTime(ByVal str2)
	Dim pos,pos2,l
	l = Len(str2)
	If Not str2 = "" Then
		pos = Instr(4,str2,":")
		If pos > 0 Then
			pos2 = Len(str2) - pos
			str2 = Left(str2,pos-1) + Mid(str2,pos+3)
		End If
	End If
	FixTime = str2
End Function

'******************************************************************************************************

Function SQLdatetime(dtDateTime)
    SQLdatetime = DatePart("yyyy", dtDateTime) & _
        Right("0" & DatePart("m", dtDateTime), 2) & _
        Right("0" & DatePart("d", dtDateTime), 2) & " " & _
        Right("00" & DatePart("h", dtDateTime), 2) & ":" & _
        Right("00" & DatePart("n", dtDateTime), 2) & ":" & _
        Right("00" & DatePart("s", dtDateTime), 2)
End Function

'******************************************************************************************************

Function SQLdatetimeADO(dtDateTime)
	If dtDateTime = "" Then
		SQLdatetimeADO = Null
	Else
		'dtDateTime = FormatDateTime(dtDateTime)
		SQLdatetimeADO = DatePart("yyyy", dtDateTime) & _
		    "/" & Right("0" & DatePart("m", dtDateTime), 2) & _
		    "/" & Right("0" & DatePart("d", dtDateTime), 2) & " " & _
		    Right("00" & DatePart("h", dtDateTime), 2) & ":" & _
		    Right("00" & DatePart("n", dtDateTime), 2) & ":" & _
		    Right("00" & DatePart("s", dtDateTime), 2)
	End If
End Function

'******************************************************************************************************

Function NullCheck(ByVal str)

	'If ErrorHandler Then
		On Error Resume Next
	'End If

    If IsNull(str) Or str = "" Then
		str=""
    Else
		str = CStr(str)
    End If

	If Err.Number <> 0 Then
		str = ""
	End If

    NullCheck = Trim(str)

End Function

'******************************************************************************************************

Function BitNullCheck(ByVal str)

	'If ErrorHandler Then
		On Error Resume Next
	'End If

    If IsNull(str) Then
		str=False
    End If

	If Err.Number <> 0 Then
		str = False
	End If

    BitNullCheck = str

End Function

'******************************************************************************************************

Function NumericNullCheck(ByVal str)

	'If ErrorHandler Then
		On Error Resume Next
	'End If

    If IsNull(str) Then
		str=0
    End If

	If Err.Number <> 0 Then
		str = 0
	End If

    NumericNullCheck = str

End Function

'******************************************************************************************************

Function NullCheckNBSP(ByVal str)

	'If ErrorHandler Then
		On Error Resume Next
	'End If

    If IsNull(str) Or str = "" Then
		str="&nbsp;"
    Else
		str = CStr(str)
    End If

	If Err.Number <> 0 Then
		str = "&nbsp;"
	End If

    NullCheckNBSP = Trim(str)

End Function

'******************************************************************************************************

Function ShowImage(ByVal str)

	If ErrorHandler Then
		On Error Resume Next
	End If

    ShowImage = ""
    If IsNull(str) Or str = "" Then

    Else
		ShowImage = "<img src=""" & Application("web_path") & Application("mapp_path") & str & """>"
    End If

End Function

'******************************************************************************************************

Function SendMail(toname,Byval toaddress,fromname,fromaddress,subject,body,bodyhtml)

	If ErrorHandler Then
		On Error Resume Next
	End If

	If Not Application("location") = "DEVELOPMENT" and Not Application("location") = "QA" Then
		' We are in PRODUCTION!
		If Application("DEBUGMAIL") Then
			If Not Application("DEBUGMAILADDR") = "" Then
				toaddress = Application("DEBUGMAILADDR")
			End If
		End If
	Else
		' We are in DEVELOPMENT / QA!
		If Application("DEBUGMAIL") Then
			If Not Application("DEBUGMAILADDR") = "" Then
				toaddress = Application("DEBUGMAILADDR")
			Else
				Exit Function
			End If
		Else
			Exit Function
		End If
	End If

	Dim Mailer, TimeoutLimit
	TimeoutLimit = 600
	Set Mailer = Server.CreateObject("Persits.MailSender")

	Mailer.Host  = Application("SMTP_Server")
	Mailer.FromName = fromname
	Mailer.From = fromaddress

	' If mail get's bounced - it goes here
	Mailer.MailFrom = Application("ReturnMail")

	Mailer.AddAddress toaddress, toname
	Mailer.Subject     = subject
	Mailer.Timeout = TimeoutLimit

	'Mailer.ContentTransferEncoding = "quoted-printable"
	'Mailer.Charset = "iso-2022-jp"

	Mailer.IsHTML = True
	Mailer.Body = bodyhtml
	Mailer.AltBody = body

	 'Mailer.AddAttachment attachment

	 'Mailer.Priority = CInt(priority)

	If Application("ASPQMail") Then

		On Error Resume Next
		If Trim(Application("objectdomain")) = "" or Len(Trim(Application("objectdomain"))) > 0 or IsNull(Application("objectdomain")) Then
			Mailer.LogonUser "", Trim(Application("objectusername")), Trim(Application("objectpassword"))
		Else
			Mailer.LogonUser Trim(Application("objectdomain")), Trim(Application("objectusername")), Trim(Application("objectpassword"))
		End If
		If Err.Number <> 0 Then
			Err.Clear
		End If
		If ErrorHandler Then
			On Error Resume Next
		Else
			On Error Goto 0
		End If

		Mailer.SendToQueue

		' if EmailAgent is running on a remote machine you will have to
		' explicitly specify the message queue directory on that machine.
		' Impersonation may be needed to create a file on the remote machine

		' Mail.LogonUser "domain", "username", "password"
		' Mail.SendToQueue "\\remoteserver\MessageQueuePath\"

		SendMail = "OK"
	Else
		On Error Resume Next
		if not Mailer.Send then
			SendMail = Err.Description
		else
			SendMail = "OK"
		end if
		On Error Goto 0
	End If

	Set Mailer = Nothing

End Function

'******************************************************************************************************

Function SendMailWithAttachment(toaddress,ccaddress,bccaddress,fromname,fromaddress,subject,body,bodyhtml,attachment,priority)

	If ErrorHandler Then
		On Error Resume Next
	End If

	If Not Application("location") = "DEVELOPMENT" and Not Application("location") = "QA" Then
		' We are in PRODUCTION!
		If Application("DEBUGMAIL") Then
			If Not Application("DEBUGMAILADDR") = "" Then
				toaddress = Application("DEBUGMAILADDR")
			End If
		End If
	Else
		' We are in DEVELOPMENT / QA!
		If Application("DEBUGMAIL") Then
			If Not Application("DEBUGMAILADDR") = "" Then
				'toaddress = Application("DEBUGMAILADDR")
			Else
				Exit Function
			End If
		Else
			Exit Function
		End If
	End If

	'Response.Write bccaddress
	'Response.End

	Dim Mailer, TimeoutLimit
	Dim CheckEmails,i

	TimeoutLimit = 600
	Set Mailer = Server.CreateObject("Persits.MailSender")

	Mailer.Host  = Application("SMTP_Server")
	Mailer.FromName = fromname
	Mailer.From = fromaddress

	' If mail get's bounced - it goes here
	Mailer.MailFrom = Application("ReturnMail")

	CheckEmails = split(toaddress & ",",",")
	For i = 0 to ubound(CheckEmails)
		If Not Trim(CheckEmails(i)) = "" Then
			Mailer.AddAddress CheckEmails(i)
		End If
	Next

	CheckEmails = split(ccaddress & ",",",")
	For i = 0 to ubound(CheckEmails)
		If Not Trim(CheckEmails(i)) = "" Then
			Mailer.AddCc CheckEmails(i)
		End If
	Next

	CheckEmails = split(bccaddress & ",",",")
	For i = 0 to ubound(CheckEmails)
		If Not Trim(CheckEmails(i)) = "" Then
			Mailer.AddBcc CheckEmails(i)
		End If
	Next

	Mailer.Subject = subject
	Mailer.Timeout = TimeoutLimit

	'Mailer.ContentTransferEncoding = "quoted-printable"
	'Mailer.Charset = "iso-2022-jp"

	Mailer.IsHTML = True
	If Not bodyhtml = "" Then
		Mailer.Body = bodyhtml
	Else
		If Not attachment = "" Then
			Mailer.AppendBodyFromFile attachment
		End If
	End If

	Mailer.AltBody = body

	If Not attachment = "" Then

		 'Mailer.ClearAttachments
		Mailer.AddAttachment attachment

	End If

	 If priority = "" Then
		Mailer.Priority = 3
	 Else
		Mailer.Priority = CInt(priority)
	 End If

	If Application("ASPQMail") Then

		On Error Resume Next
		If Trim(Application("objectdomain")) = "" or Len(Trim(Application("objectdomain"))) > 0 or IsNull(Application("objectdomain")) Then
			Mailer.LogonUser "", Trim(Application("objectusername")), Trim(Application("objectpassword"))
		Else
			Mailer.LogonUser Trim(Application("objectdomain")), Trim(Application("objectusername")), Trim(Application("objectpassword"))
		End If
		If Err.Number <> 0 Then
			Err.Clear
		End If
		If ErrorHandler Then
			On Error Resume Next
		Else
			On Error Goto 0
		End If
		Mailer.SendToQueue

		' if EmailAgent is running on a remote machine you will have to
		' explicitly specify the message queue directory on that machine.
		' Impersonation may be needed to create a file on the remote machine

		' Mail.LogonUser "domain", "username", "password"
		' Mail.SendToQueue "\\remoteserver\MessageQueuePath\"

		SendMailWithAttachment = "OK"
	Else
		On Error Resume Next
		if not Mailer.Send then
			SendMailWithAttachment = Err.Description
		else
			SendMailWithAttachment = "OK"
		end if
		On Error Goto 0
	End If

	Set Mailer = Nothing

End Function

'******************************************************************************************************

FUNCTION Proper(ByVal strInput)

	If ErrorHandler Then
		On Error Resume Next
	End If

  Dim S, L, SLen, UChars, PrevChar, CurChar

  S = ""
  PrevChar = " "
  SLen = Len(strInput)
  UChars = " `1234567890-=" & _
     "~!@#$%^&*()_+[]\{}|;':"",./<>?" & _
     Chr(9) & Chr(10) & Chr(13)

  ' See if we have a string
  IF (SLen < 1) THEN
    Proper = ""
    EXIT FUNCTION
  END IF

  ' Loop through and properize
  For L = 1 To SLen
    IF (InStr(UChars, PrevChar) > 0) THEN
      CurChar = UCase(Mid(strInput, L, 1))
    ELSE
      CurChar = LCase(Mid(strInput, L, 1))
    END IF
    S = S & CurChar
    PrevChar = CurChar
  NEXT

  ' Return value
  Proper = S

END FUNCTION

'******************************************************************************************************

FUNCTION CloseObj(byref the_object )

	On Error Resume Next

	If IsNull ( the_object ) Then
		Err.Clear
		Exit Function
	End If

	If IsObject( the_object ) then
		If not the_object Is Nothing Then
			If Not the_object.state = 0 Then
				the_object.Close
			End If
			Set the_object = Nothing
		End If
	End If

	Err.Clear

END FUNCTION

'******************************************************************************************************

Function WeAreDown()
	If ErrorHandler Then
		On Error Resume Next
	End If
	WeAreDown = Application("LOGINSDISABLED")
	If Trim(UCase(Request.QueryString("ipod"))) = "Y" Then
		WeAreDown = False
	End If
End Function

'******************************************************************************************************

FUNCTION DisplayError(tmessage,taction,treferer)

	If ErrorHandler Then
		On Error Resume Next
	End If

    If isNull(tmessage) or isEmpty(tmessage) or Trim(tmessage) = "" Then
	  tmessage = "None"
	Else
	  tmessage = Trim(tmessage)
	End If

    If isNull(taction) or isEmpty(taction) or Trim(taction) = "" Then
	  taction = "L"
	Else
	  taction = Trim(taction)
	End If

    If isNull(treferer) or isEmpty(treferer) or Trim(treferer) = "" Then
	  taction = "L"
	  treferer = ""
	Else
	  treferer = Trim(treferer)
	End If

	If Application("SendEmailOnError") Then
		SendMail "IPOD",Application("DEBUGMAILADDR"),Trim(Application("ProductShortName")) & " (" & Proper(Application("Location")) & " Environment)",Application("DEBUGMAILADDR"),Trim(Application("ProductShortName"))& " (" & Proper(Application("Location")) & " Environment)" & " Error", "Dear IPOD," & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "FYI: The error below has occurred in the " & Trim(Application("ProductShortName")) & " " & Proper(Application("Location")) & " Environment." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Replace(tmessage,"<br>",Chr(13) & Chr(10))
	End If

	Call GotoPage(Application("web_path") & Application("mapp_path") & "DBError.asp?ASPError="_
					 +Server.URLEncode(tmessage)+"&ErrorAction="+taction _
					 +"&Http_Referer="+treferer, True)

	' This is cool - If the redirect to DBERROR got an error because content has
	' been written to the browser via a Response.Flush - it will fail and simply write
	' out the info here. :)

	Response.Write ("<br><br><font color=""" & graphics("TEXTWARN") & _
							""" size=""2"" face=""Arial""><strong>" & _
							tmessage & "</strong></font></center><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>")

	Response.End


END FUNCTION

'******************************************************************************************************

FUNCTION CheckForError(ErrorType,Conn,RS)

	If ErrorHandler Then
	'	On Error Resume Next
	End If

	Dim MyErrMessage
	Dim	dbError
	Dim	RetryQuery
	Dim accessViolation

	accessViolation = false

	MyErrMessage = ""

	If Trim(ErrorType) = "" Then
		ErrorType = "V"
	End If

	Select Case Trim(ErrorType)

		Case "V"

				If Err.Number <> 0 Then
					MyErrMessage = Trim(Err.Description)
					Trim(Err.Source)
				Else
					Exit Function
				End If

		Case "A"

			If Not IsObject(RS) Then

				If Err.Number <> 0 Then
					MyErrMessage = Trim(Err.Description)
				Else
					For Each dbError In Conn.Errors
						MyErrMessage = MyErrMessage + dbError.Description + "<br>"
					Next
					Conn.Errors.Clear
				End If

			Else

				If Err.Number <> 0 Then
					MyErrMessage = Trim(Err.Description)
				Else
					do	WHILE (Conn.Errors.Count > 0)

						RetryQuery = False
						For Each dbError In Conn.Errors
							If (dberror.NativeError = 1205) Then
									MyErrMessage = ""
									RetryQuery = True
									EXIT FOR
							end if

							If (dberror.NativeError = 91111) Then
								accessViolation = true
							end if

							MyErrMessage = MyErrMessage + dbError.Description + "<br>"
							'MyErrMessage = MyErrMessage + "ADO Native Error #" + CStr(dbError.NativeError) + "<br>"
							'MyErrMessage = MyErrMessage + "ADO SQLState #" + dbError.SQLState + "<br>"
						Next

						Conn.Errors.Clear
						If RetryQuery Then
							RS.Open cmdTemp, , 0, 1
						End If

					Loop
				End If

			End If

			If Trim(MyErrMessage) = "" Then
				Exit Function
			End If

		Case Else

			Exit Function

	End Select

	Err.Clear

	IF (accessViolation) Then
		Call OutputWAPError("Access Violation")
	Else
		Call OutputWAPError("An error has occurred on page " & _
							Trim(Request.ServerVariables("PATH_INFO")) & "." & _
							"<br/><br/>" & MyErrMessage)
	End If

END FUNCTION

'******************************************************************************************************

Function ShowError(ByVal ErrorMessage)

	If ErrorHandler Then
		On Error Resume Next
	End If

	ErrorMessage = Trim(ErrorMessage + " " + RedMessage)
	If (not ErrorMessage = "") then
		Response.Write ("<tr><td><font color=""" & graphics("TEXTWARN") & _
							""" size=""2"" face=""Arial""><strong>" & _
							ErrorMessage & "</strong></font></td></tr><tr><td>&nbsp;</td></tr>")
	End If
End Function

'******************************************************************************************************

Function GotoPage(newurl,top)

	' Always Resume Next incase headers were written before the redirect
	On Error Resume Next

	Dim wheretogo

	CommitSession

	If InStr(newurl,"http://") > 0 or InStr(newurl,"https://") > 0 Then
		wheretogo = newurl
	Else
		If InStr(newurl,"../") > 0 Then
			wheretogo = newurl
		Else
			If not InStr(newurl,Application("web_path") & Application("mapp_path")) > 0 Then
				wheretogo = Application("web_path") & Application("mapp_path") & newurl
			Else
				wheretogo = newurl
			End If
		End If
	End If

	%>
	<script language="JavaScript">
	<% If top Then %>
		top.location.href = "<% = wheretogo %>"
	<% Else %>
		self.location.href = "<% = wheretogo %>"
	<% End If %>
	</script>
	<%
	Response.End

End Function

'******************************************************************************************************

Sub dogenericvalid(isok,thedesc,thepk,theid)

	If Not isok Then

		newjs = newjs + "var changedesc = eval('myframe." & fieldname & "Desc');" & nl
		newjs = newjs + "if (changedesc)" & nl
		newjs = newjs + "{" & nl
		newjs = newjs + "	changedesc.innerHTML = '<br>' + '<img name=""" + fieldname + "Desc_msgicon"" border=""0"" src=""" & Application("web_path") & Application("mapp_path") & "images/blank.gif"">&nbsp;<font class=""mc_lookupdescerror"">" + fieldlabel + " Not Found</font>';" & nl
		newjs = newjs + "	changedesc.style.display = '';" & nl
		newjs = newjs + "	if (top.fraPaneBar) {mydoc.images." & fieldname & "Desc_msgicon.src = top.fraPaneBar.warnicon.src;} else {mydoc.images." & fieldname & "Desc_msgicon.src = '../../images/warn.gif';}" & nl
		newjs = newjs + "}" & nl

		Response.Clear
		returnmessage = ""
		returnclass = "errormessage"
		Call DoResponse(theaction,newjs,False,False,False,"")

	Else

		newjs = newjs + "mydoc.getElementsByName('" & fieldname & "').item(0).value = '" & theid & "';" & nl
		newjs = newjs + "var changedesc = eval('myframe." & fieldname & "Desc');" & nl
		newjs = newjs + "if (changedesc)" & nl
		newjs = newjs + "{" & nl
		If UCase(thedoc) = "TOP.FRAEXTERNAL" Then
			newjs = newjs + "	myframe.setdesc(changedesc,'" & JSEncode(thedesc) & "');" & nl
			newjs = newjs + "	myframe.setpk('"+fieldname+"PK','" & thepk & "');" & nl
			'newjs = newjs + "	alert(myframe.name+': " + fieldname+"PK');"+nl
		Else
			newjs = newjs + "	top.setdesc(changedesc,'" & JSEncode(thedesc) & "');" & nl
			newjs = newjs + "	top.setpk('"+fieldname+"PK','" & thepk & "');" & nl
			'newjs = newjs + "	alert(top.name+': " + fieldname+"PK');"+nl
		End If
		newjs = newjs + "	if (changedesc.mchidden != null && changedesc.mchidden.toUpperCase() == 'Y') {changedesc.style.display = 'none';}" & nl
		newjs = newjs + "}" & nl

		'newjs = newjs + "alert('here');"

		'Response.Write "<code>" & newjs & "</code>"
		'Response.End

	End If

End Sub

Sub dogenericendprocess(newjs)
	' Used for validation procedures

	Response.Write("	top.showmessage('" + returnmessage + "','" + returnclass + "');") + nl
	If Not returnmessage = "" Then

		Response.Write("	if (top.timeoutid != null)") + nl
		Response.Write("	{") + nl
		Response.Write("		top.clearTimeout(top.timeoutid);") + nl
		Response.Write("	}") + nl

		Response.Write("	top.timeoutid = top.setTimeout(""top.removemessage()"",1000);") + nl
	End If
	Response.Write("	top.endprocess();") + nl
	Response.Write("	top.showactions('" + Trim(theaction) + "');") + nl

	Response.Write(newjs)

End Sub

Sub dogenericendaction(nosound)
	' Used for Action procedures (Save)

	Response.Write("// Ending Process")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl

	Response.Write("	top.checkpendingmsg();") + nl
	Response.Write("	top.endprocess();") + nl

	Response.Write("	top.showmessage('" + returnmessage + "','" + returnclass + "');") + nl
	If Not returnmessage = "" Then

		Response.Write("	if (top.timeoutid != null)") + nl
		Response.Write("	{") + nl
		Response.Write("		top.clearTimeout(top.timeoutid);") + nl
		Response.Write("	}") + nl

		Response.Write("	top.timeoutid = top.setTimeout(""top.removemessage()"",1000);") + nl
	End If

	If Not nosound Then
		Response.Write("	top.playsound('sounds/done.wav');") + nl
	End If

End Sub

Sub dogenericheader()

	If Application("OnErrorResumeNext") Then
		On Error Resume Next
	End If

	donocacheheader
	CommitSession

	Response.Write mcCopyright
	Response.Write("<html>") + nl
	Response.Write("<head>") + nl
	Response.Write("<title></title>") + nl
	Response.Write("<script language=""JavaScript"">") + nl
	Response.Write("<!--") + nl
	Response.Write("window.onerror = top.doError;") + nl
	Response.Write(formjs)

End Sub

Sub dogenericfooter()

	If Application("OnErrorResumeNext") Then
		On Error Resume Next
	End If

	Response.Write("//-->") + nl
	Response.Write("</script>") + nl
	Response.Write("<body>") + nl

	FlushItNoStore

	If Application("ASPDEBUG") then
		aspdebug
	End If

	Response.Write("</body>") + nl
	Response.Write("</html>") + nl

End Sub

Function db_child(db,d,htmltable,sqltable,cwhere,nrecs,dwhere)

	Dim rs,OutArray,isinsert
	Dim theindex,theindexanddate,rowversiondate,n,newrecs,eof,suffix,fp

	If ErrorHandler Then
		On Error Resume Next
	End If

	If Not db.dok Then
		Exit Function
	End If

	' Perform Adds and Updates First
	If Not cwhere = "" or Not nrecs = "" Then

		If Not nrecs = "" Then
			newrecs = split(nrecs,"#!#")
		Else
			newrecs = Array("None")
		End If

		If Not cwhere = "" Then
			cwhere = Left(cwhere,Len(cwhere)-1)
			Set rs = db.RunSQLReturnRS_RW("SELECT * FROM " & sqltable & " WHERE PK IN (" & cwhere & ")","")
			'Response.Write("SELECT * FROM " & sqltable & " WHERE PK IN (" & cwhere & ")")
			'Response.End
		Else
			Set rs = db.RunSQLReturnRS_RW("SELECT TOP 0 * FROM " & sqltable,"")
		End If

		If Not db.dok Then
			Exit Function
		End If

		n = 0
		eof = rs.EOF

		Do While (Not eof) or (n < ubound(newrecs))

			If eof Then
				theindex = Mid(newrecs(n),2)
				suffix = Left(newrecs(n),1)
				rs.AddNew
				isinsert = True
			Else
				isinsert = False
				theindexanddate = split(d(htmltable & NullCheck(rs("PK"))),"$")

				theindex = theindexanddate(0)
				RowVersionDate = theindexanddate(1)

				'Response.Write RowVersionDate
				'Response.Write "<br>"
				'Response.Write rs("RowVersionDate")
				'Response.Write "<br>"
				'Response.Write Trim(RowVersionDate) = Trim(rs("RowVersionDate"))

				suffix = "1" ' default value

				' ENSURE THE ROWVERSIONDATE MATCHES THE ONE FROM THE POST OR IT IS A STALE CHILD RECORD
				'--------------------------------------------------------------------------------------------------------------------------------------------------------------
				On Error Resume Next
				If Not Trim(RowVersionDate) = Trim(NullCheck(rs("RowVersionDate"))) Then
					If Err.Number <> 0 Then
						Err.Clear
					Else
						db.dok = False
						db.derror = "Another user has made modifications to a detail record while you were working with it. Since you could potentially overwrite changes the other user made, you will need to Cancel your changes and start again. Please click the CANCEL button to cancel your changes."
						Exit Function
					End If
				End If
				If ErrorHandler Then
					On Error Resume Next
				Else
					On Error Goto 0
				End If
				'--------------------------------------------------------------------------------------------------------------------------------------------------------------

			End If

			Execute("Call db_" & htmltable & "(rs,isinsert,suffix,htmltable,theindex)")

			If rs.EOF or eof Then
				n = n + 1
				eof = True
			Else
				rs.MoveNext()
				If rs.Eof Then
					eof = True
				End If
			End If

		Loop

		' Put this in if we run into problems with the following error:

		' rs.properties("update criteria") = adCriteriaKey

		' "The specified row could not be located for updating: Some values may have
		' been changed since it was last read."

		' -2147217864 ADO error


		'Response.ContentType = "application/xml"
		'Response.AddHeader "Content-Disposition", "attachment;filename=ReportOutput.xml"
		''Response.Write "<?xml version=""1.0"" ?>" & Chr(13) & Chr(10)
		'rs.Save Response, adPersistXML
		'Response.End

		db.dobatchupdate rs

	    Set rs.ActiveConnection = Nothing
		rs.close
		Set rs = Nothing

	End If

	' Perform Deletes
	If Not dwhere = "" Then

		dwhere = Left(dwhere,Len(dwhere)-1)
		Call db.RunSQL("DELETE FROM " & sqltable & " WITH (ROWLOCK) WHERE PK IN (" & dwhere & ")","")

	End If

End Function

'******************************************************************************************************

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
			Exit Function
		End If

	End If

End Function

'******************************************************************************************************

Sub SaveUDF(rs)

	If Not IsEmpty(Request.Form("txtUDFChar1")) Then
		If Len(Trim(Request.Form("txtUDFChar1").Item)) > 0 Then
			rs("UDFChar1") = Trim(Mid(Request.Form("txtUDFChar1"),1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("UDFChar1") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtUDFChar2")) Then
		If Len(Trim(Request.Form("txtUDFChar2").Item)) > 0 Then
			rs("UDFChar2") = Trim(Mid(Request.Form("txtUDFChar2"),1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("UDFChar2") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtUDFChar3")) Then
		If Len(Trim(Request.Form("txtUDFChar3").Item)) > 0 Then
			rs("UDFChar3") = Trim(Mid(Request.Form("txtUDFChar3"),1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("UDFChar3") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtUDFChar4")) Then
		If Len(Trim(Request.Form("txtUDFChar4").Item)) > 0 Then
			rs("UDFChar4") = Trim(Mid(Request.Form("txtUDFChar4"),1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("UDFChar4") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtUDFChar5")) Then
		If Len(Trim(Request.Form("txtUDFChar5").Item)) > 0 Then
			rs("UDFChar5") = Trim(Mid(Request.Form("txtUDFChar5"),1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("UDFChar5") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtUDFDate1")) Then
		If Not Request.Form("txtUDFDate1") = "" Then
			rs("UDFDate1") = Request.Form("txtUDFDate1")	' Nullable: YES Type: datetime
		Else
			rs("UDFDate1") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtUDFDate2")) Then
		If Not Request.Form("txtUDFDate2") = "" Then
			rs("UDFDate2") = Request.Form("txtUDFDate2")	' Nullable: YES Type: datetime
		Else
			rs("UDFDate2") = Null
		End If
	End If
	rs("UDFBit1") = Not Request.Form("txtUDFBit1") = ""	' Nullable: No Type: bit
	rs("UDFBit2") = Not Request.Form("txtUDFBit2") = ""	' Nullable: No Type: bit

End Sub

'******************************************************************************************************

Sub db_version(rs)

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	If DemoMode() and NewRecord Then

		On Error Resume Next

		' This makes it so that only the records the demo user adds
		' will show up in their Explorer Lists. It also prevents them from
		' seeing other demo user's records - which may contain
		' private email addresses and / or other data.
		rs("DemoLaborPK") = GetSession("UserPK")

		If ErrorHandler Then
			On Error Resume Next
		Else
			On Error Goto 0
		End If

	End If
	'--------------------------------------------------------------------------------------------------------------

	If Not GetSession("UserIPAddress") = "" Then
		rs("RowVersionIPAddress") = GetSession("UserIPAddress")	' Nullable: YES Type: int
	Else
		rs("RowVersionIPAddress") = Null
	End If
	rs("RowVersionUserPK") = GetSession("UserPK")	' Nullable: YES Type: int
	'rs("RowVersionUserID") = Trim(Mid(GetSession("UserID"),1,25))	' Nullable: YES Type: varchar
	If Not GetSession("UserInitials") = "" Then
		rs("RowVersionInitials") = Trim(Mid(GetSession("UserInitials"),1,5))									' Nullable: No Type: nvarchar
	Else
		rs("RowVersionInitials") = Null
	End If
	rs("RowVersionDate") = ServerDate

	'If Not IsEmpty(Request.Form("txtRowVersionUserPK")) Then
	'	If Len(Trim(Request.Form("txtRowVersionUserPK"))) > 0 Then
	'		rs("RowVersionUserPK") = Request.Form("txtRowVersionUserPK")	' Nullable: YES Type: int
	'	End If
	'End If
	'If Not IsEmpty(Request.Form("txtRowVersionUser")) Then
	'	rs("RowVersionUserID") = Trim(Mid(GetSession("UserID"),1,25))	' Nullable: YES Type: varchar
	'End If
	'If Not IsEmpty(Request.Form("txtRowVersionInitials")) Then
	'	rs("RowVersionInitials") = Trim(Mid(Request.Form("txtRowVersionInitials"),1,5))									' Nullable: No Type: nvarchar
	'End If

End Sub

'******************************************************************************************************

Function IsValidTime(whatTime)
	On Error Resume Next
	Dim w
	w = TimeValue(whatTime)
	if Err.Number <> 0 Then
		IsValidTime = False
	Else
		IsValidTime = w
	End If
End Function

'******************************************************************************************************

Function GetValidTime(whatTime)
	On Error Resume Next
	Dim w
	w = TimeValue(whatTime)
	if Err.Number <> 0 Then
		GetValidTime = False
	Else
		GetValidTime = w
	End If
End Function

'******************************************************************************************************

Sub DemoNoEditMsg(db,rs)

	' Let the users change the sample records
	' 3/1/03
	Exit Sub
	Dim newjs,playsound,warntext

	If DemoMode() and Not NewRecord Then
		If IsNull(rs("DemoLaborPK")) Then

			warntext = "In demo mode, you can not save changes to sample records. You can, however, create new records and make changes to your new records."
			newjs = newjs + "   top.mcalert('info','Demo Message','" & warntext & "','bg_okprint',600,240,'sounds/error.wav');" + nl
			newjs = newjs + "	top.cancelchanges();" + nl
			db.CloseClientConnection
			Call CloseObj(rs)
			Set db = Nothing
			playsound = False
			Call DoScript(newjs,False)

		End If
	End If

End Sub

'******************************************************************************************************

Function DemoNoActionMultiMsg(db,sql)

	' Let the users change the sample records
	' 3/1/03
	DemoNoActionMultiMsg = ""
	Exit Function

	Dim rs,newjs,playsound,warntext

	newjs = ""

	Set rs = db.RunSQLReturnRS(sql,"")
	Call dok_check(db,"Demo Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

	If Not rs.eof Then

		warntext = "The selected action was only performed on non-sample data records. In demo mode, you can not perform actions to sample records. You can, however, create new records and perform actions on your new records."
		newjs = "   top.mcalert('info','Demo Message','" + warntext + "','bg_okprint',700,255,'sounds/error.wav');" + nl

	End If

	Call CloseObj(rs)
	DemoNoActionMultiMsg = newjs

End Function

'******************************************************************************************************

Sub DemoNoActionMsg(db,sql)

	' Let the users change the sample records
	' 3/1/03
	Exit Sub
	Dim rs,newjs,playsound,warntext

	Set rs = db.RunSQLReturnRS(sql,"")
	Call dok_check(db,"Demo Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

	If rs.eof Then

		warntext = "In demo mode, you can not perform actions to sample records. You can, however, create new records and perform actions on your new records."
		newjs = newjs + "   top.mcalert('info','Demo Message','" & warntext & "','bg_okprint',600,240,'sounds/error.wav');" + nl
		'newjs = newjs + "	top.cancelchanges();" + nl
		newjs = newjs + "	top.showactions('" & LastAction & "');" + nl
		newjs = newjs + "	top.removemessage();" + nl
		newjs = newjs + "	top.playsound('sounds/done.wav');" + nl

		db.CloseClientConnection
		Call CloseObj(rs)
		Set db = Nothing
		playsound = False
		Call DoScript(newjs,False)

	End If

End Sub

'******************************************************************************************************

Sub dok_check(db,errortitle,errortext)

	Dim dok,derror,newjs,playsound,aok

	dok = db.dok
	derror = db.derror

	If Not dok Then

		db.CloseClientConnection
		Set db = Nothing

		newjs = newjs + "   top.mcalert('warning','" & errortitle & "','" & errortext & "<br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
		playsound = False
		aok = False
		On Error Resume Next
		returnmessage = ""
		On Error Goto 0
		Response.Clear
		Call DoResponse("",newjs,playsound,True,False,"")

	End If

End Sub

'******************************************************************************************************

Sub dok_check_afterflush(db,errortitle,errortext)

	Dim dok,derror,newjs,playsound

	dok = db.dok
	derror = db.derror

	If Not dok Then

		db.CloseClientConnection
		Set db = Nothing

		newjs = newjs + "   top.mcalert('warning','" & errortitle & "','" & errortext & "<br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
		newjs = newjs + "	if (document && document.body) {document.body.style.backgroundImage = '';}" + nl
		newjs = newjs + "   if (top.name.toUpperCase() != 'MCMAIN') {top.close();}" + nl

		%>
		<script language="JavaScript">
			if (top.doError) {window.onerror = top.doError;}
			<% =formjs %>
			if (top.endprocess) {top.endprocess();}
			<% =newjs %>
		</script>
		<%
		Response.End
	End If

End Sub

'******************************************************************************************************

Sub dok_check_afterflush_noinfo(db,errortitle,errortext)

	Dim dok,derror,newjs,playsound

	dok = db.dok
	derror = db.derror

	If Not dok Then

		db.CloseClientConnection
		Set db = Nothing

		newjs = newjs + "   top.mcalert('warning','" & errortitle & "','" & errortext & "<div id=""errorinfo"" style=""display:none;""><br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "<br><br></div>','bg_okprint',700,255,'sounds/error.wav');" + nl
		newjs = newjs + "	if (document && document.body) {document.body.style.backgroundImage = '';}" + nl
		newjs = newjs + "   if (top.name.toUpperCase() != 'MCMAIN') {top.close();}" + nl

		%>
		<script language="JavaScript">
			if (top.doError) {window.onerror = top.doError;}
			<% =formjs %>
			top.endprocess();
			<% =newjs %>
		</script>
		<%
		Response.End
	End If

End Sub

'******************************************************************************************************

Sub dok_check_popup(db,errortitle,errortext)

	Dim dok,derror,newjs,playsound

	dok = db.dok
	derror = db.derror

	If Not dok Then

		db.CloseClientConnection
		Set db = Nothing

		newjs = newjs + "   thetop.mcalert('warning','" & errortitle & "','" & errortext & "<br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
		newjs = newjs + "	if (document && document.body) {document.body.style.backgroundImage = '';}" + nl
		newjs = newjs + "   if (top.name.toUpperCase() != 'MCMAIN') {top.close();}" + nl

		%>
		<script language="JavaScript">
			var thetop = top.dialogArguments.caller;
			if (thetop.doError) {window.onerror = thetop.doError;}
			<% =formjs %>
			if (thetop.endprocess) {thetop.endprocess();}
			<% =newjs %>
		</script>
		<%
		Response.End
	End If

End Sub

'******************************************************************************************************

Function OutputArray(alldata)

	Dim BlankSpace,shownull,cols,rows,field,rowcounter,colcounter

	BlankSpace="&nbsp;"
	shownull="-null-"

	cols=ubound(alldata,1)
	rows=ubound(alldata,2)

	response.write "<table cellspacing=""0"" cellpadding=""4"">"
	FOR rowcounter= 0 TO rows
	   response.write "<tr>" & vbcrlf
	   FOR colcounter=0 to cols
	      field=alldata(colcounter,rowcounter)
	      if isnull(field) then
	         field=shownull
	      end if
	      if trim(field)="" then
	         field=BlankSpace
	      end if
	      response.write "<td nowrap valign=top>"
	      response.write field
	      response.write "</td>" & vbcrlf
	   NEXT
	   response.write "</tr>" & vbcrlf
	NEXT
	response.write "</table>"

End Function

'******************************************************************************************************

Function URLDecode(strToDecode)
    Dim strIn
    Dim strOut
    Dim strLeft
    Dim strRight
    Dim intPos
    Dim intLoop

    strIn = strToDecode
    strOut = ""
    intPos = Instr(strIn, "+")   'Look for + signs and replace with space
    Do While intPos
        strLeft = ""
        strRight = ""
        If intPos > 1 Then
            strLeft = Left(strIn, intPos - 1)
        End If
        If intPos < Len(strIn) Then
           strRight = Mid(strIn, intPos + 1)
        End If
        strIn = strLeft & " " & strRight
        intPos = Instr(strIn, "+")
        intLoop = intLoop + 1
    Loop
    intPos = Instr(strIn, "%")  'Look for ASCII coded characters
    Do While intPos
        If intPos > 1 Then strOut = strOut & Left(strIn, intPos - 1)
        strOut = strOut & Chr(CInt("&H" & Mid(strIn, intPos + 1, 2)))
        If intPos > (Len(strIn) - 3) Then
            strIn = ""
        Else
            strIn = Mid(strIn, intPos + 3)
        End If
        intPos = Instr(strIn, "%")   'look for the next one
    Loop
    URLDecode = strOut & strIn
End Function

'******************************************************************************************************

Function getclientimage(img)
	getclientimage = GetSession("webHTTP") & GetWebServer() & Application("ImageServer") & img
End Function

'******************************************************************************************************

Function showclientimg(imgsrc)
	showclientimg = Application("ImageServer") & GetSession("db") & "/" & imgsrc
End Function
'******************************************************************************************************

Function QuoteList(ByVal ThisValue)

	Dim NewValue,i
	NewValue=""

	ThisValue=split(ThisValue,",")

	If IsArray(ThisValue) Then
		For i=0 to ubound(ThisValue)
			NewValue = NewValue & "'" & Trim(ThisValue(i)) & "'"

			if i < ubound(ThisValue) Then
				NewValue = NewValue & ","
			End If
		Next
	Else
		NewValue = ThisValue
	End If

	QuoteList = NewValue

End Function

'******************************************************************************************************

Function BinToText(varBinData, intDataSizeInBytes)
   Dim objRS

   Set objRS = Server.CreateObject("ADODB.Recordset")

   objRS.Fields.Append "txt", adVarChar, intDataSizeInBytes, adFldLong
   objRS.Open

   objRS.AddNew
   objRS.Fields("txt").AppendChunk varBinData
   BinToText = objRS("txt").Value

   objRS.Close
   Set objRS = Nothing
End Function

'******************************************************************************************************

Function RemotePageGrab(inRemoteURL)
   Dim tempPos1
   Dim timeoutResolve, timeoutConnect, timeoutSend, timeoutReceive
   Dim strRetVal, objHTTPGet

   strRetVal = ""

   'Response.Write "Getting information...<br>"

   'ON ERROR RESUME NEXT
   Set objHTTPGet = Server.CreateObject("MSXML2.ServerXMLHTTP")

   ' Sets the timeout for the http connection
   timeoutResolve = 5 * 1000
   timeoutConnect = 5 * 1000
   timeoutSend = 5 * 1000
   timeoutReceive = 12 * 1000
   objHTTPGet.setTimeouts timeoutResolve, timeoutConnect, timeoutSend, timeoutReceive

   objHTTPGet.open "GET", inRemoteURL, False
   If Err.Number <> 0 Then
      Response.Write "error on HTTP Open, Detail: " & Err.Description
      Response.Write "<br>"
   Else
      objHTTPGet.Send
      If Err.Number <> 0 Then
         Response.Write "error on HTTP Send, Detail: " & Err.Description
         Response.Write "<br>"
      Else
         'strRetVal = BinToText(objHTTPGet.responseBody, 4800)
         strRetVal = objHTTPGet.responseText
      End If
   End If

   Set objHTTPGet = Nothing
   ON ERROR GOTO 0

   RemotePageGrab = strRetVal
End Function

'******************************************************************************************************

Sub DemoModeCheck()

	Dim dok,derror,newjs,playsound

	If DemoMode() Then

		Dim errortext

		errortext = "Since you are in demo mode, the selected action has been disabled. Please call our sales department at 1-888-567-3434 to have your organization setup with a trial solution that would have access to the selected action."
		newjs = newjs + "   top.mcinfo(null,;" & errortext & "');" + nl

		playsound = False
		aok = False
		On Error Resume Next
		returnmessage = ""
		On Error Goto 0
		Response.Clear
		Call DoResponse("",newjs,playsound,True,False,"")

	End If

End Sub

'******************************************************************************************************

Sub DemoModeCheck_AfterFlush()

	Dim dok,derror,newjs,playsound

	If DemoMode() Then

		Dim errortext

		errortext = "Since you are in demo mode, the selected action has been disabled. Please call our sales department at 1-888-567-3434 to have your organization setup with a trial solution that would have access to the selected action"
		newjs = newjs + "   top.mcinfo(null,;" & errortext & "');" + nl
		newjs = newjs + "	if (document && document.body) {document.body.style.backgroundImage = '';}" + nl
		newjs = newjs + "   if (top.name.toUpperCase() != 'MCMAIN') {top.close();}" + nl

		%>
		<script language="JavaScript">
			<% =formjs %>
			top.endprocess();
			<% =newjs %>
		</script>
		<%
		Response.End

	End If

End Sub

'******************************************************************************************************

Function DemoMode()

	DemoMode = False

	If Trim(UCase(GetSession("DM"))) = "Y" Then
		DemoMode = True
	End If

End Function

'******************************************************************************************************

Sub SendApprovalEmail(db,PersonApprovedEmail,FullName,FirstName)

	Dim mailsubject, strBody, strBodyHTML, mailerror

	mailsubject = "Approved: Your Connection to " & Trim(GetSession("en"))

	strBody = "Dear " & FirstName & "," & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
		& "This email has been sent to you to notify you that you have been " _
		& "approved to access " & Trim(GetSession("en")) & ". " _
		& "If you have any " _
        & "questions, please send an email to " & Trim(GetSession("an")) & " at " & Trim(GetSession("ae")) & "." & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
		& "Thank you," & Chr(13) & Chr(10) & Chr(13) & Chr(10)
		If Trim(UCase(GetSession("an"))) = Trim(UCase(Application("ProductShortName"))) Then
		strBody = strBody _
		& Application("ProductProducer") & Chr(13) & Chr(10) _
		& Trim(GetSession("an")) & Chr(13)
		Else
		strBody = strBody _
		& Application("ProductProducer")  & " on behalf of" & Chr(13) & Chr(10) _
		& Trim(GetSession("an")) & Chr(13)
		End If

	strBodyHTML = _
		"<html>" & vbCrLf & _
		"<head>" & vbCrLf & _
		"</head>" & vbCrLf & _
		"<body link=""#0000FF"" vlink=""#0000FF"">" & vbCrLf & _
		"<p><img border=""0"" src=""http://" & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/mc_logonewsmall2.gif"" width=""190"" height=""45""></p>" & vbCrLf & _
		"<p><font face=""Arial"">Dear " & FirstName & ",<br>" & vbCrLf & _
		"<br>" & vbCrLf & _
		"This email has been sent to you to notify you that you have been approved to access " & Trim(GetSession("en")) & ".<br><br><img border=""0"" src=""http://" & GetWebServer() & showclientimg(GetSession("el")) & """><br><br> " & _
		"If you have any questions, please send an email to " & Trim(GetSession("an")) & " at <a href=""mailto:" & Trim(GetSession("ae")) & """>" & Trim(GetSession("ae")) & "</a>." & vbCrLf & _
		"<br><br>" & vbCrLf & _
		"Thank you,<br><br>" & vbCrLf
		If Trim(UCase(GetSession("an"))) = Trim(UCase(Application("ProductShortName"))) Then
		strBodyHTML = strBodyHTML & _
		Application("ProductProducer") & _
		"&nbsp;</font></p>" & vbCrLf & _
		"</body>" & vbCrLf & _
		"</html>" & vbCrLf
		Else
		strBodyHTML = strBodyHTML & _
		Application("ProductProducer") & " on behalf of<br>" & _
		Trim(GetSession("an")) & _
		"&nbsp;</font></p>" & vbCrLf & _
		"</body>" & vbCrLf & _
		"</html>" & vbCrLf
		End If

	mailerror = SendMail(FullName,PersonApprovedEmail,Application("ProductProducer"),Trim(GetSession("ae")),mailsubject,strBody,strBodyHTML)
	If Not mailerror = "OK" Then
		db.warn = True
		db.warntext = Trim(db.warntext + "The record saved successfully, but there was a problem encountered while sending out the notification email. (" & mailerror & ").")
	End If

End Sub

'******************************************************************************************************

Sub OutputAttachments(rs)

	Dim docmodid,PrintWithWO,SendWithEmail,attachments,docloc
	For attachments = 1 To 2
		Do While Not rs.eof
			docmodid = Trim(RS("ModuleID"))
			Response.Write("	top.builddatarow(myframe.oat1body,3,null,'NOGUID$NOVERSION','','',true,'',null,null,'" & docmodid & "','" & JSEncode(RS("TitleforDocumentList")) & "');")+nl
			Do While (Not rs.eof)
				If Not docmodid = NullCheck(RS("ModuleID")) Then
					Exit Do
				End If
				If RS("LocationType") = "LIBRARY" Then
					docloc = JSEncode(RS("LocationTypeDesc"))
				Else
					docloc = JSEncode(RS("Location"))
					'If Len(docloc) > 65 Then
					'	docloc = Left(docloc,65) & "..."
					'End If
				End If
				If RS("PrintWithWO") Then
					PrintWithWO = "<img src=""../../images/taskchecked.gif"">"
				Else
					PrintWithWO = "<img src=""../../images/taskline.gif"">"
				End If
				If RS("SendWithEmail") Then
					SendWithEmail = "<img src=""../../images/taskchecked.gif"">"
				Else
					SendWithEmail = "<img src=""../../images/taskline.gif"">"
				End If
				If attachments = 1 Then
					Response.Write("	top.builddatarow(myframe.oat1body,2,null,'NOGUID$NOVERSION','" & RS("DocumentPK") & "','DO',true,'" & JSEncode(RS("Photo")) & "',null,null,'" & docmodid & "','" & JSEncode(RS("LocationTypeDesc")) & "','" & JSEncode(RS("DocumentTypeDesc")) & "','" & JSEncode(RS("DocumentID")) & " (" & JSEncode(RS("DocumentName")) & ")','" & docloc & "','','" & PrintWithWO & "','" & SendWithEmail & "');")+nl
				Else
					Response.Write("	top.builddatarow(myframe.oat1body,2,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("DocumentPK") & "','DO',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & docmodid & "','" & JSEncode(RS("LocationTypeDesc")) & "','" & JSEncode(RS("DocumentTypeDesc")) & "','" & JSEncode(RS("DocumentID")) & " (" & JSEncode(RS("DocumentName")) & ")','" & docloc & "','','" & PrintWithWO & "','" & SendWithEmail & "');")+nl
				End If
				rs.MoveNext
			Loop
		Loop
		If attachments = 1 Then
			Set rs = rs.NextRecordset
		End If
	Next
	Response.Write(nl)

End Sub

'******************************************************************************************************

Sub SetScriptTimeoutTo(timeinminutes)
	Server.ScriptTimeout = 60 * timeinminutes
End Sub

Sub ResetScriptTimeOut()
	If Application("DefaultScriptTimeout") = "" Then
		Server.ScriptTimeout = 90
	Else
		Server.ScriptTimeout = Application("DefaultScriptTimeout")
	End If
End Sub

'----------------------------------------------------------------------
' XMLEncode
'
' Encodes the given variant for storage inside an XML string.  If the
' variant is a boolean, this function returns "1" or "0".  Otherwise
' XMLEncode treats the variant as a string and replaces all invalid
' XML tokens with their escaped counterparts.
'----------------------------------------------------------------------
Function XMLEncode(var)
	On Error Resume Next

	Dim strTmp

	If (IsNull(var)) Then
		var = ""
	End If

	If (VarType(var) = adBoolean) Then
		If (var) Then
			strTmp = "1"
		Else
			strTmp = "0"
		End If
	Else
		strTmp = CStr(var)
		strTmp = Replace(strTmp, "&", "&amp;")
		strTmp = Replace(strTmp, "<", "&lt;")
		strTmp = Replace(strTmp, ">", "&gt;")
		strTmp = Replace(strTmp, """", "&quot;")
		strTmp = Replace(strTmp, "'", "&apos;")
	End If

	XMLEncode = strTmp

End Function

'******************************************************************************************************

Function GetAccessRight(db, ActionID, byref MaxAmount)

	Dim rs

	GetAccessRight = True
	MaxAmount = 0

	Set rs = db.RunSPReturnRS("MC_GetLaborAccessRight",Array(Array("@accessgroupPK", adInteger, adParamInput, 4, GetSession("AGPK")),Array("@actionID", adVarChar, adParamInput, 25, ActionID)),"")
	If db.dok Then
		If Not rs.eof Then
			GetAccessRight = rs("Enabled")
			MaxAmount = rs("MaxAmount")
		End If
	End If
	Call CloseObj(rs)

End Function

'******************************************************************************************************

Function GetPreference(db,ispk,rcpk,prefname,byref prefvalue,byref prefdesc,byref prefpk)

	Dim OutArray
	OutArray = ""

	Call db.RunSP("MC_GetLaborPrefs",Array(Array("@LaborPK", adInteger, adParamInput, 4, GetSession("USERPK")),Array("@RepairCenterPK", adInteger, adParamInput, 4, rcpk),Array("@PreferenceName", advarchar, adParamInput, 50, prefname),Array("@NoRecordSet", adBoolean, adParamInput, 1, True),Array("@ModuleID", advarchar, adParamInput, 2, vbNull),Array("@PreferenceOutput", advarchar, adParamOutput, 1000, ""),Array("@PreferenceDescOutput", advarchar, adParamOutput, 50, ""),Array("@PreferencePKOutput", adInteger, adParamOutput, 4, 0)),outarray)

	If db.dok Then
		GetPreference = True
		prefvalue = Trim(NullCheck(OutArray(5)))
		prefdesc = Trim(NullCheck(OutArray(6)))
		prefpk = NullCheck(OutArray(7))
		If (ispk and prefpk = "") or _
		   (prefvalue = "") Then
			GetPreference = False
			prefvalue = ""
			prefdesc = ""
			prefpk = ""
		End If
	Else
		GetPreference = False
		prefvalue = ""
		prefdesc = ""
		prefpk = ""
	End If

End Function

'******************************************************************************************************

Function GetDefaultPreference(db,ispk,prefname,byref prefvalue,byref prefdesc,byref prefpk)

	Dim OutArray
	OutArray = ""

	Call db.RunSP("MC_GetDefaultPrefs",Array(Array("@PreferenceName", advarchar, adParamInput, 50, prefname),Array("@NoRecordSet", adBoolean, adParamInput, 1, True),Array("@ModuleID", advarchar, adParamInput, 2, vbNull),Array("@PreferenceOutput", advarchar, adParamOutput, 50, ""),Array("@PreferenceDescOutput", advarchar, adParamOutput, 50, ""),Array("@PreferencePKOutput", adInteger, adParamOutput, 4, 0)),outarray)

	If db.dok Then
		GetDefaultPreference = True
		prefvalue = Trim(NullCheck(OutArray(3)))
		prefdesc = Trim(NullCheck(OutArray(4)))
		prefpk = NullCheck(OutArray(5))
		If (ispk and prefpk = "") or _
		   (prefvalue = "") Then
			GetDefaultPreference = False
			prefvalue = ""
			prefdesc = ""
			prefpk = ""
		End If
	Else
		GetDefaultPreference = False
		prefvalue = ""
		prefdesc = ""
		prefpk = ""
	End If

End Function

'******************************************************************************************************
' Called from all modules after SAVE to see if the explorer should be refreshed.

Function RefreshExplorerIsTurnedOn(db,Module,NewRecord)

	Dim prefvalue, prefdesc, prefpk, prefname

	RefreshExplorerIsTurnedOn = False

	' The explorer refresh after save can be programatically overrided
	' by setting the hidden variable noexplorerrefresh to the value of 'Y'.
	If Request("noexplorerrefresh") = "Y" Then
		Exit Function
	End If

	If NewRecord Then
		prefname = Module & "_EXPLORER_ONSAVE_NEW"
	Else
		prefname = Module & "_EXPLORER_ONSAVE_EDT"
	End If

	If GetPreference(db,False,prefname,prefvalue, prefdesc, prefpk) Then
		If UCase(prefvalue) = "YES" Then
			RefreshExplorerIsTurnedOn = True
		End If
	End If

End Function

'===================================================================================================================

Sub DoListLine(pos)

Dim h1,h2
If pos = "TOP" Then
	h1 = 0
	h2 = 4
Else
	h1 = 4
	h2 = 7
End If
%>
  	<tr>
  		<td></td>
		<td colspan="<% =fieldtotal %>" height="<% =h1 %>"></td>
	</tr>
  	<tr>
  		<td height="1" bgcolor="#D7D7D7"></td>
		<td colspan="<% =fieldtotal %>" height="1" bgcolor="#D7D7D7"></td>
	</tr>
  	<tr>
  		<td></td>
		<td colspan="<% =fieldtotal %>" height="<% =h2 %>"></td>
	</tr>
<%
End Sub

'===================================================================================================================

Sub OutputBottomTabs(TabNumber)

	Select Case TabNumber

		Case 2

			Response.Write("<div id=""pagetabs"" style=""display:none;"">" & vbNewline  & _
			vbNewLine & _
			"<map name=""BottomTabsMap"">" & vbNewline  & _
			"<area  onclick=""top.changepage(self.pagetabs_21);"" shape=""rect"" coords=""5, 0, 56, 14"">" & vbNewline  & _
			"<area  onclick=""top.changepage(self.pagetabs_22);"" shape=""rect"" coords=""59, 0, 143, 14"">" & vbNewline  & _
			"</map>" & vbNewline  & _
			"<img name=""pagetabs_img"" style=""cursor:hand; position:relative; top:-12px;"" border=""0"" src=""../../images/pagetabs_21.gif"" usemap=""#BottomTabsMap"">" & vbNewline  & _
			vbNewLine & _
			"</div>" & vbNewline)

		Case 3

			Response.Write("<div id=""pagetabs"" style=""display:none;"">" & vbNewline  & _
			vbNewLine & _
			"<map name=""BottomTabsMap"">" & vbNewline  & _
			"<area onclick=""top.changepage(self.pagetabs_31);"" shape=""rect"" coords=""3, 0, 55, 14"">" & vbNewline  & _
			"<area onclick=""top.changepage(self.pagetabs_32);"" shape=""rect"" coords=""58, 0, 104, 14"">" & vbNewline  & _
			"<area onclick=""top.changepage(self.pagetabs_33);"" shape=""rect"" coords=""108, 0, 192, 14"">" & vbNewline  & _
			"</map>" & vbNewline  & _
			vbNewLine & _
			"<img name=""pagetabs_img"" style=""cursor:hand; position:relative; top:-12px;"" border=""0"" src=""../../images/pagetabs_31.gif"" usemap=""#BottomTabsMap"">" & vbNewline  & _
			vbNewLine & _
			"</div>" & vbNewline)

		End Select

End Sub

'===================================================================================================================

Sub OutputUDFHTML(tablename)
%>
<table border="0" cellspacing="0" width="100%" cellpadding="2">
  <tr>
    <td valign="top">
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
      <tr>
        <td valign="top" style="padding-right:25;">
        <table border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td colspan="2" class="fieldsheader" style="border-bottom:1px solid #A2A2A2;">
            Text Fields</td>
          </tr>
          <tr>
            <td nowrap valign="top" style="padding-top:10;">
            <label FOR="txtUDFChar1" ACCESSKEY>
            <a href="javascript:void(0);" class="fieldlabel" onclick="return top.showhelp('<% =tablename %>','UDFChar1','User-Defined',txtUDFChar1);this.blur();" TabIndex="-1">
            <font face="arial"><span id="lblUDFChar1">Field 1</span>:</font></a> </label>&nbsp;</td>
            <td valign="top" style="padding-bottom:8; padding-top:10;">
            <input class="normal" mcType="C" maxlength="50" mcRequired="N" type="text" name="txtUDFChar1" id="txtUDFChar1" tabindex="100" size="41" onChange="top.codevalidUDF('<% =LCase(tablename) %>_udfchar1',this,self.lblUDFChar1.innerText, self.document.images.imgUDFChar1);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img style="display:none;" id="imgUDFChar1" src="../../images/lookupiconxp3.gif" border="0" align="absbottom" onclick="top.showpopup('<% =LCase(tablename) %>_udfchar1',self.lblUDFChar1.innerText,266,129,this,txtUDFChar1)" class="lookupicon" WIDTH="16" HEIGHT="20">
            <span mchidden="y" style="display:none;" id="txtUDFChar1Desc" class="mc_lookupdesc"></span></td>
          </tr>
          <tr>
            <td nowrap valign="top">
            <label FOR="txtUDFChar2" ACCESSKEY>
            <a href="javascript:void(0);" class="fieldlabel" onclick="return top.showhelp('<% =tablename %>','UDFChar2','User-Defined',txtUDFChar2);this.blur();" TabIndex="-1"><font face="arial">
            <span id="lblUDFChar2">Field 2</span>:</font></a> </label>&nbsp;</td>
            <td valign="top" style="padding-bottom:8">
            <input class="normal" mcType="C" maxlength="50" mcRequired="N" type="text" name="txtUDFChar2" id="txtUDFChar2" tabindex="101" size="41" onChange="top.codevalidUDF('<% =LCase(tablename) %>_udfchar2',this,self.lblUDFChar2.innerText, self.document.images.imgUDFChar2);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img style="display:none;" id="imgUDFChar2" src="../../images/lookupiconxp3.gif" border="0" align="absbottom" onclick="top.showpopup('<% =LCase(tablename) %>_udfchar2',self.lblUDFChar2.innerText,266,129,this,txtUDFChar2)" class="lookupicon" WIDTH="16" HEIGHT="20">
            <span mchidden="y" style="display:none;" id="txtUDFChar2Desc" class="mc_lookupdesc"></span></td>
          </tr>
          <tr>
            <td nowrap valign="top">
            <label FOR="txtUDFChar3" ACCESSKEY><a href="javascript:void(0);" class="fieldlabel" onclick="return top.showhelp('<% =tablename %>','UDFChar3','User-Defined',txtUDFChar3);this.blur();" TabIndex="-1"><font face="arial">
            <span id="lblUDFChar3">Field 3</span>:</font></a> </label>&nbsp;</td>
            <td valign="top" style="padding-bottom:8">
            <input class="normal" mcType="C" maxlength="50" mcRequired="N" type="text" name="txtUDFChar3" id="txtUDFChar3" tabindex="102" size="41" onChange="top.codevalidUDF('<% =LCase(tablename) %>_udfchar3',this,self.lblUDFChar3.innerText, self.document.images.imgUDFChar3);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img style="display:none;" id="imgUDFChar3" src="../../images/lookupiconxp3.gif" border="0" align="absbottom" onclick="top.showpopup('<% =LCase(tablename) %>_udfchar3',self.lblUDFChar3.innerText,266,129,this,txtUDFChar3)" class="lookupicon" WIDTH="16" HEIGHT="20">
            <span mchidden="y" style="display:none;" id="txtUDFChar3Desc" class="mc_lookupdesc"></span></td>
          </tr>
          <tr>
            <td nowrap valign="top">
            <label FOR="txtUDFChar4" ACCESSKEY><a href="javascript:void(0);" class="fieldlabel" onclick="return top.showhelp('<% =tablename %>','UDFChar4','User-Defined',txtUDFChar4);this.blur();" TabIndex="-1"><font face="arial">
            <span id="lblUDFChar4">Field 4</span>:</font></a> </label>&nbsp;</td>
            <td valign="top" style="padding-bottom:8">
            <input class="normal" mcType="C" maxlength="50" mcRequired="N" type="text" name="txtUDFChar4" id="txtUDFChar4" tabindex="103" size="41" onChange="top.codevalidUDF('<% =LCase(tablename) %>_udfchar4',this,self.lblUDFChar4.innerText, self.document.images.imgUDFChar4);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img style="display:none;" id="imgUDFChar4" src="../../images/lookupiconxp3.gif" border="0" align="absbottom" onclick="top.showpopup('<% =LCase(tablename) %>_udfchar4',self.lblUDFChar4.innerText,266,129,this,txtUDFChar4)" class="lookupicon" WIDTH="16" HEIGHT="20">
            <span mchidden="y" style="display:none;" id="txtUDFChar4Desc" class="mc_lookupdesc"></span></td>
          </tr>
          <tr>
            <td nowrap valign="top">
            <label FOR="txtUDFChar5" ACCESSKEY><a href="javascript:void(0);" class="fieldlabel" onclick="return top.showhelp('<% =tablename %>','UDFChar5','User-Defined',txtUDFChar5);this.blur();" TabIndex="-1"><font face="arial">
            <span id="lblUDFChar5">Field 5</span>:</font></a> </label>&nbsp;</td>
            <td valign="top" style="padding-bottom:8">
            <input class="normal" mcType="C" maxlength="50" mcRequired="N" type="text" name="txtUDFChar5" id="txtUDFChar5" tabindex="104" size="41" onChange="top.codevalidUDF('<% =LCase(tablename) %>_udfchar5',this,self.lblUDFChar5.innerText, self.document.images.imgUDFChar5);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img style="display:none;" id="imgUDFChar5" src="../../images/lookupiconxp3.gif" border="0" align="absbottom" onclick="top.showpopup('<% =LCase(tablename) %>_udfchar5',self.lblUDFChar5.innerText,266,129,this,txtUDFChar5)" class="lookupicon" WIDTH="16" HEIGHT="20">
            <span mchidden="y" style="display:none;" id="txtUDFChar5Desc" class="mc_lookupdesc"></span></td>
          </tr>
          <tr>
            <td colspan="2" class="fieldsheader" style="border-bottom:1px solid #A2A2A2;padding-top:10">
            Date Fields</td>
          </tr>
          <tr>
            <td nowrap valign="top" style="padding-top:10;">
            <label FOR="txtUDFDate1" ACCESSKEY>
            <a href="javascript:void(0);" class="fieldlabel" onclick="return top.showhelp('<% =tablename %>','UDFDate1','User-Defined',txtUDFDate1);this.blur();" TabIndex="-1">
            <font face="arial"><span id="lblUDFDate1">Field 6</span>:</font></a> </label>&nbsp;</td>
            <td valign="top" style="padding-bottom:8; padding-top:10;">
            <input class="normal" mcType="D" maxlength="10" mcRequired="N" type="text" name="txtUDFDate1" id="txtUDFDate1" tabindex="105" size="15" onChange="top.fieldvalid(this);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img id="imgUDFDate1" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calendar','Calendar',172,160,this,txtUDFDate1)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
            <span id="txtUDFDate1Err" class="mc_lookupdesc"></span></td>
          </tr>
          <tr>
            <td nowrap valign="top"><label FOR="txtUDFDate2" ACCESSKEY>
            <a href="javascript:void(0);" class="fieldlabel" onclick="return top.showhelp('<% =tablename %>','UDFDate2','User-Defined',txtUDFDate2);this.blur();" TabIndex="-1">
            <font face="arial"><span id="lblUDFDate2">Field 7</span>:</font></a> </label>&nbsp;</td>
            <td valign="top" style="padding-bottom:8">
            <input class="normal" mcType="D" maxlength="10" mcRequired="N" type="text" name="txtUDFDate2" id="txtUDFDate2" tabindex="106" size="15" onChange="top.fieldvalid(this);" onfocus="top.fieldfocus(this);" onblur="top.fieldblur(this);" onkeypress="return top.checkKey(this);"><img id="imgUDFDate2" src="../../images/lookupiconxp3.gif" border="0" onclick="top.showpopup('calendar','Calendar',172,160,this,txtUDFDate2)" align="absbottom" class="lookupicon" WIDTH="16" HEIGHT="20">
            <span id="txtUDFDate2Err" class="mc_lookupdesc"></span></td>
          </tr>
          <tr>
            <td colspan="2" class="fieldsheader" style="border-bottom:1px solid #A2A2A2;padding-top:10">
            Checkbox Fields</td>
          </tr>
		  <tr>
			<td colspan="2" style="padding-top:5px;">
			    <table cellspacing="0" cellpadding="0" border="0">
			    <tr>
				<td>
					<input class="mccheckbox" mcType="B" mcRequired="N" type="checkbox" value="ON" name="txtUDFBit1" id="txtUDFBit1" tabindex="99" onclick="top.togglecb(this);">
				</td>
				<td style="padding-top:3;padding-left:6;">
					<a href="javascript:void(0);" class="fieldlabel" onclick="top.togglecheckbox(this);" title="Right-Click for Help" oncontextmenu="return top.showhelp('<% =tablename %>','UDFBit1','User-Defined',txtUDFBit1);this.blur();" TabIndex="-1">
					<font face="arial"><span id="lblUDFBit1">Field 8</span></font></a>
				</td>
				</tr>
				</table>
			</td>
		  </tr>
		  <tr>
			<td colspan="2">
			    <table cellspacing="0" cellpadding="0" border="0">
			    <tr>
				<td>
					<input class="mccheckbox" mcType="B" mcRequired="N" type="checkbox" value="ON" name="txtUDFBit2" id="txtUDFBit2" tabindex="99" onclick="top.togglecb(this);">
				</td>
				<td style="padding-top:3;padding-left:6;">
					<a href="javascript:void(0);" class="fieldlabel" onclick="top.togglecheckbox(this);" title="Right-Click for Help" oncontextmenu="return top.showhelp('<% =tablename %>','UDFBit2','User-Defined',txtUDFBit2);this.blur();" TabIndex="-1">
					<font face="arial"><span id="lblUDFBit2">Field 9</span></font></a>
				</td>
				</tr>
				</table>
			</td>
		  </tr>
        </table>
        </td>
      </tr>
      <tr>
        <td nowrap valign="top"></td>
        <td valign="top"></td>
      </tr>
      <tr>
        <td nowrap valign="top"></td>
        <td valign="top"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<%
End Sub

'===================================================================================================================

Sub OutputUDFData(rs)

		Response.Write("	myform.txtUDFChar1.value = '" & JSEncode(RS("UDFChar1")) & "';")+nl
		Response.Write("	myform.txtUDFChar2.value = '" & JSEncode(RS("UDFChar2")) & "';")+nl
		Response.Write("	myform.txtUDFChar3.value = '" & JSEncode(RS("UDFChar3")) & "';")+nl
		Response.Write("	myform.txtUDFChar4.value = '" & JSEncode(RS("UDFChar4")) & "';")+nl
		Response.Write("	myform.txtUDFChar5.value = '" & JSEncode(RS("UDFChar5")) & "';")+nl
		Response.Write("	if (myframe.txtUDFChar1Desc) myframe.txtUDFChar1Desc.innerText = '';")+nl
		Response.Write("	if (myframe.txtUDFChar2Desc) myframe.txtUDFChar2Desc.innerText = '';")+nl
		Response.Write("	if (myframe.txtUDFChar3Desc) myframe.txtUDFChar3Desc.innerText = '';")+nl
		Response.Write("	if (myframe.txtUDFChar4Desc) myframe.txtUDFChar4Desc.innerText = '';")+nl
		Response.Write("	if (myframe.txtUDFChar5Desc) myframe.txtUDFChar5Desc.innerText = '';")+nl
		Response.Write("	myform.txtUDFDate1.value = '" & DateNullCheck(RS("UDFDate1")) & "';")+nl
		Response.Write("	myform.txtUDFDate2.value = '" & DateNullCheck(RS("UDFDate2")) & "';")+nl
		If RS("UDFBit1") = 0 Then
			Response.Write("	myform.txtUDFBit1.checked = false;")+nl
		Else
			Response.Write("	myform.txtUDFBit1.checked = true;")+nl
		End If
		If RS("UDFBit2") = 0 Then
			Response.Write("	myform.txtUDFBit2.checked = false;")+nl
		Else
			Response.Write("	myform.txtUDFBit2.checked = true;")+nl
		End If

End Sub

'===================================================================================================================

Sub OutputUDFLabels(rs,moduletable)

	Dim db

	' If Initial Module Load (IML) Then Write the Field Labels
	If Request.QueryString("iml") = "Y" Then
		If Not moduletable = "" Then
			Set db = New ADOHelper
			Set rs = db.RunSPReturnRS("MC_GetUDFLabels",Array(Array("@TableName", advarchar, adParamInput, 255, moduletable)),"")
			Call dok_check(db,"UDF Message","There was a problem retrieving the User Defined Labels. The details of the problem are described below. You can try to create the Work Order again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
		End If
		Response.Write("")+nl
		Response.Write("// Write UDF Field Label Data")+nl
		Response.Write("// -------------------------------------------------------------------------")+nl
		Do While Not rs.Eof
			Response.Write("if (myframe.lbl" & JSEncode(rs("column_name")) & ") myframe.lbl" & JSEncode(rs("column_name")) & ".innerText = '" & JSEncode(rs("field_label")) & "';")+nl
			If UCase(Left(rs("column_name"),7)) = "UDFCHAR" Then
				If rs("lookup_table") Then
					Response.Write("if (mydoc.images.img" & JSEncode(rs("column_name")) & ") mydoc.images.img" & JSEncode(rs("column_name")) & ".style.display = '';")+nl
				Else
					Response.Write("if (mydoc.images.img" & JSEncode(rs("column_name")) & ") mydoc.images.img" & JSEncode(rs("column_name")) & ".style.display = 'none';")+nl
				End If
			End If
			rs.MoveNext()
		Loop
	End If

End Sub

'===================================================================================================================

Sub OutputInTextArea(t,name,hide)
	If name = "" Then
		If hide Then
			Response.write "<textarea style=""display:none;font-family:arial; font-size:10pt; background-color:#FFFFCC; width:100%;height:80;"" id=textarea1 name=textarea1>" & t & "</textarea>"
		Else
			Response.write "<textarea style=""font-family:arial; font-size:10pt; background-color:#FFFFCC; width:100%;height:80;"" id=textarea1 name=textarea1>" & t & "</textarea>"
		End If
	Else
		If hide Then
			Response.write "<textarea name=""" & name & """ id=""" & name & """ style=""display:none;font-family:arial; font-size:10pt; background-color:#FFFFCC; width:100%;height:80;"" id=textarea1 name=textarea1>" & t & "</textarea>"
		Else
			Response.write "<textarea name=""" & name & """ id=""" & name & """ style=""font-family:arial; font-size:10pt; background-color:#FFFFCC; width:100%;height:80;"" id=textarea1 name=textarea1>" & t & "</textarea>"
		End If
	End If
End Sub

'===================================================================================================================

Function Shorten(ByVal str,l)

    Dim oldstr

	On Error Resume Next

	If l = -1 Then
		l = 75
	End If

    oldstr = str

    If IsNull(str) Or str = "" Then
		str=""
    Else
		If Len(str) > l Then
			str = Left(str,l) & "..."
		End If
    End If

	If Err.Number <> 0 Then
		str = oldstr
	End If

    Shorten = Trim(str)

End Function

'===================================================================================================================

Function ReadINI(file, section, key, value_default)

	Dim FileSysObj, ini, reSection, reKey, line

	Set FileSysObj = Server.CreateObject("Scripting.FileSystemObject")

	ReadIni=value_default

	If FileSysObj.FileExists(file) Then

	     Set ini = FileSysObj.OpenTextFile( file, 1, False)

	     Set reSection = new RegExp
	     reSection.Global =False
	     reSection.IgnoreCase=True
	     reSection.Pattern ="\s*\[\s*" & section & "\s*\]"

	     Set reKey = new RegExp
	     reKey.Global =False
	     reKey.IgnoreCase=True
	     reKey.Pattern="\s*" & key & "\s*=\s*"

	     Do While ini.AtEndofStream = False

	          line = ini.ReadLine

	               if reSection.Test(line) = True then

	                    line=ini.ReadLine

	                    do while instr(line,"[")=0 and ini.AtEndofStream = False

	                         if reKey.Test(line) then

	                              ReadINI=trim(mid(line,instr(line,"=")+1))
	                              exit do

	                         end if

	                         line=ini.ReadLine
	                    Loop

                         if reKey.Test(line) then

                              ReadINI=trim(mid(line,instr(line,"=")+1))
                              exit do

                         end if

	          exit do
	          end if
	     loop

	     ini.Close
	     Set reSection=nothing
	     Set reKey =nothing

	End If ' If FileSysObj

End Function

'===================================================================================================================

Function FixMapPath(p)
	Dim temp_path
	FixMapPath = Trim(p)
	If Application("IsCompiled") Then
		FixMapPath = Replace(FixMapPath,"/","\")
		FixMapPath = Replace(FixMapPath,"\\","\")
		temp_path = Replace(Application("mapp_path"),"/","\")
		FixMapPath = Replace(FixMapPath,temp_path & temp_path,temp_path)
		temp_path = Replace(Application("rapp_path"),"/","\")
		FixMapPath = Replace(FixMapPath,temp_path & temp_path,temp_path)
		FixMapPath = Replace(FixMapPath,"mcadmin\mcadmin\","mcadmin\")
		FixMapPath = Replace(FixMapPath,"online\online\","online\")
		FixMapPath = Replace(FixMapPath,"onsite\onsite\","onsite\")
	End If
End Function

'===================================================================================================================

Function DoNotify(ByVal eids, ByVal PK, db)

'Not in this release
Exit Function

On Error Resume Next

Dim rs, rs2, sql, OutArray, EmailAddresses
Dim mailsubject, strBody, strBodyHTML, emailfrom, emailfromaddress, mailerror
Dim NotifyA, i, PKLookup

If eids = "" Then
	Exit Function
End If

DoNotify = False
OutArray = ""

NotifyA = split(eids & ",",",")
For i = 0 to ubound(NotifyA)
	If Not Trim(NotifyA(i)) = "" Then

		If InStr(NotifyA(i),"#") > 0 Then
			PKLookup = Mid(NotifyA(i),InStr(NotifyA(i),"#")+1)
			NotifyA(i) = Mid(NotifyA(i),1,InStr(NotifyA(i),"#")-1)
		Else
			PKLookup = PK
		End If

		sql = "SELECT * FROM MCEvent WITH (NOLOCK) WHERE EventID = '" & NotifyA(i) & "' AND IsNull(RepairCenterPK," & GetSession("RCPK") & ") = " & GetSession("RCPK")
		SET rs = db.RunSQLReturnRS(SQL,"")

		If Not db.dok Then
			db.warn = True
			db.warntext = Trim(db.warntext + "The record saved successfully, but there was a problem encountered while sending out the notification email. (" & db.derror & ").")
			Exit Function
		End If
		If rs.eof Then
			db.warn = True
			db.warntext = Trim(db.warntext + "The record saved successfully, but there was a problem encountered while sending out the notification email. (Event '" & NotifyA(i) & "' Not Found).")
			Exit Function
		End If

		Do Until rs.eof

			If rs("active") Then

				Set rs2 = db.RunSPReturnMultiRS("MC_GetNotifyInfo",Array(Array("@EventPK", adInteger, adParamInput, 4, rs("EventPK")),Array("@PK", adInteger, adParamInput, 4, PK),Array("@PKLookup", adInteger, adParamInput, 4, PKLookup)),OutArray)

				If Not db.dok Then
					db.warn = True
					db.warntext = Trim(db.warntext + "The record saved successfully, but there was a problem encountered while sending out the notification email. (" & db.derror & ").")
					Exit Function
				End If

				EmailAddresses = NullCheck(rs2("emailaddresses"))

				'Response.Write EmailAddresses
				'Response.End

				If Not EmailAddresses = "" Then

					Set rs2 = rs2.NextRecordset()

					mailsubject = "Notification: " & Trim(NullCheck(rs("EventName")))

					strBody = NullCheck(rs("EmailTemplate"))
					strBody = Replace(strBody,"##ENTITYNAME##",GetSession("en"))
					strBody = Replace(strBody,"##PRODUCT##",Application("ProductProducer"))

					strBodyHTML = NullCheck(rs("EmailTemplateHTML"))
					strBodyHTML = Replace(strBodyHTML,"##ENTITYNAME##",GetSession("en"))
					strBodyHTML = Replace(strBodyHTML,"##PRODUCT##",Application("ProductProducer"))
					strBodyHTML = Replace(strBodyHTML,"##ENTITYLOGO##","<img border=""0"" src=""" & GetSession("webHTTP") & GetWebServer() & showclientimg(GetSession("el")) & """><br clear=""all"">")
					strBodyHTML = Replace(strBodyHTML,"##MCLOGO##","<img border=""0"" src=""http://" & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/mc_logonewsmall2.gif"">")
					strBodyHTML = strBodyHTML & "<p><img border=""0"" src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/mcrptgenlogo.jpg""></p>"

					Select Case NullCheck(rs("ModuleID"))
						Case "AS"
							strBody = Replace(strBody,"@@ASSET.ASSETID@@",NullCheck(rs2("AssetID")))
							strBody = Replace(strBody,"@@ASSET.ASSETNAME@@",NullCheck(rs2("AssetName")))
							strBodyHTML = Replace(strBodyHTML,"@@ASSET.ASSETID@@",NullCheck(rs2("AssetID")))
							strBodyHTML = Replace(strBodyHTML,"@@ASSET.ASSETNAME@@",NullCheck(rs2("AssetName")))

							If NotifyA(i) = "ASSubStatusChange" Then
								strBody = Replace(strBody,"@@ASSET.STATUS@@",NullCheck(rs2("Status")))
								strBody = Replace(strBody,"@@ASSET.STATUSDESC@@",NullCheck(rs2("StatusDesc")))
								strBodyHTML = Replace(strBodyHTML,"@@ASSET.STATUS@@",NullCheck(rs2("Status")))
								strBodyHTML = Replace(strBodyHTML,"@@ASSET.STATUSDESC@@",NullCheck(rs2("StatusDesc")))
							End If
							If NotifyA(i) = "ASSpecOutofRange" Then
								strBody = Replace(strBody,"@@ASSETSPECIFICATION.SPECIFICATIONNAME@@",NullCheck(rs2("SpecificationName")))
								strBody = Replace(strBody,"@@ASSETSPECIFICATION.VALUENUMERIC@@",NullCheck(rs2("ValueNumeric")))
								strBodyHTML = Replace(strBodyHTML,"@@ASSETSPECIFICATION.SPECIFICATIONNAME@@",NullCheck(rs2("SpecificationName")))
								strBodyHTML = Replace(strBodyHTML,"@@ASSETSPECIFICATION.VALUENUMERIC@@",NullCheck(rs2("ValueNumeric")))
							End If
						Case "PD"
						Case "PJ"
						Case "PO"
						Case "IN"
						Case "SY"
						Case "WO"
					End Select

					emailfrom = NullCheck(rs("EmailFromName"))
					If emailfrom = "" Then
						emailfrom = "Maintenance Connection"
					End If
					emailfromaddress = NullCheck(rs("EmailFromAddress"))
					If emailfromaddress = "" Then
						emailfromaddress = Application("FromMail")
					End If

					If err.number = 0 Then

						mailerror = SendMailWithAttachment("","",emailaddresses,emailfrom,emailfromaddress,mailsubject,strBody,strBodyHTML,"","")
						If Not mailerror = "OK" Then
							db.warn = True
							db.warntext = Trim(db.warntext + "The record saved successfully, but there was a problem encountered while sending out the notification email. (" & mailerror & ").")
						End If

					End If

				End If

			End If

			rs.MoveNext()

		Loop

	End If
Next

End Function

'===================================================================================================================

Function Build_ewherehtmlbox(title,sql_ewhere)

	Dim ewherehtmlbox
	ewherehtmlbox = ""
	ewherehtmlbox = ewherehtmlbox & "<div style=""margin-bottom:10px;""><fieldset style=""height:162px; border:1 solid #C0C0C0;"">"
	ewherehtmlbox = ewherehtmlbox & "<legend style=""font-size:10pt; color:royalblue;"" class=""legendHeader"">" & title & "</legend>"
	ewherehtmlbox = ewherehtmlbox & "<div style=""height:132px; overflow-y:auto; margin-top:2px; padding-left:10px; padding-right:10px; padding-top:10px; padding-bottom:10px;""><table cellpadding=""2"" cellspacing=""0"" border=""0"">"
	If Len(sql_ewhere) > 6 Then
	  sql_ewhere = sql_ewhere & "</td></tr>"
	  If InStr(sql_ewhere,"WHERE") > 0 Then
	 	ewherehtmlbox = ewherehtmlbox & Mid(sql_ewhere,7)
	  Else
	 	ewherehtmlbox = ewherehtmlbox & sql_ewhere
	  End If
	Else
	 	ewherehtmlbox = ewherehtmlbox & "<span style=""padding-left:5px; font-family:arial; font-size:10pt;"">None</span>"
	End If
	ewherehtmlbox = ewherehtmlbox & "</table></div>"
	ewherehtmlbox = ewherehtmlbox & "</fieldset></div>"

	Build_ewherehtmlbox = ewherehtmlbox

End Function

'===================================================================================================================

Function RCCheck(PK)

    ' Below check is for Stock Rooms (which do not have to have an RC)
    If IsNull(PK) Then
        RCCheck = True
        Exit Function
    End If

    Dim RCList
    RCList = GetSession("RCDeny")
    'Response.Write RCList
    If RCList = "" Then
        RCCheck = True
        Exit Function
    End If
    ' Sending "" means we are checking "All Repair Centers" and if we are only granting access to some - then this returns False
    If PK = "" Then
        RCCheck = False
        Exit Function
    End If

    'Response.Write "<br/>HERE"

    RCList = "," & RCList & ","

    If InStr(RCList,","&PK&",") > 0 Then
        RCCheck = True
    Else
        RCCheck = False
    End If

End Function

'===================================================================================================================

Function RSConcat(ors,pos)
    RSConcat = ""
    Do While Not ors.Eof
        If RSConcat = "" Then
            RSConcat = RSConcat & ors(pos)
        Else
            RSConcat = RSConcat & "," & ors(pos)
        End If
        ors.MoveNext()
    Loop
End Function

'===================================================================================================================

Function FixInt(xString)
     dim rx,rValue
     Set rx = new RegExp
     rx.Global = True
     rx.IgnoreCase = True
     rx.Pattern = "[^0-9]"
     rValue = Trim(rx.Replace(xString, ""))
     If rValue = "" Then
        FixInt = 0
     Else
        FixInt = rValue
     End If
End Function
%>