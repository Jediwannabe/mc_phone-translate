<%
Sub Authenticate()

	Call SetSession("logintime",Now())

	If Not Request("SearchBy") = "" Then
		txtmembername = GetSession("m")
		txtpassword = GetSession("p")
	Else
		If txtmembername = "" Then
			txtmembername = Server.HTMLEncode(Request("membername"))
			txtpassword = Server.HTMLEncode(Request("password"))
		End If
		txtmembername = Trim(Mid(txtmembername,1,50))
		txtpassword = Trim(Mid(txtpassword,1,20))
		Call SetSession("m",txtmembername)
		Call SetSession("p",txtpassword)
	End If

	If Not txtmembername = "" and Not txtpassword = "" and Not Application("loginsdisabled") Then

		' Validation against DB goes HERE

		Retval = -1
		Retval2 = -1
		rscounter = 1
		retvalcounter = 1

		Call sp_start("PSPT_WebAuthenticate","NEWCONNECTION")
		cmd.Parameters.Append cmd.CreateParameter("@p_logon",  adVarchar, adParamInput, 50, txtmembername)
		cmd.Parameters.Append cmd.CreateParameter("@p_password",  adVarchar, adParamInput, 20, txtpassword)
		cmd.Parameters.Append cmd.CreateParameter("@p_resource",  adVarchar, adParamInput, 36, p_resource)
		cmd.Parameters.Append cmd.CreateParameter("@p_ipaddress",  adVarchar, adParamInput, 15, Request.ServerVariables ("remote_addr"))
		cmd.Parameters.Append cmd.CreateParameter("@p_returnemailaccounts", adBoolean, adParamInput, , 0)
		RS.Open cmd, , adOpenForwardOnly, adLockReadOnly
		Call CheckForError("A",cmd.ActiveConnection,RS)
		Call scattertoarrays("KEEPALIVE")

		'Response.Write "Retval = " & retval & "<br>"

		If Retval = 0 Then

			' now we need to call PSPT_GetUserRights

			Call sp_start("PSPT_GetUserRights","EXISTINGCONNECTION")
			cmd.Parameters.Append cmd.CreateParameter("@p_userGuid",  adVarchar, adParamInput, 36, rsdata1(0,0))
			cmd.Parameters.Append cmd.CreateParameter("@p_resource",  adVarchar, adParamInput, 36, p_resource)
			RS.Open cmd, , adOpenForwardOnly, adLockReadOnly
			Call CheckForError("A",cmd.ActiveConnection,RS)
			Call scattertoarrays("KILLONCOMPLETE")

			Call CloseObj( con )

			If Retval2 = 0 and rsdata2recs Then

				'Call outputarray(rsdata4)
				'Response.End

				If rsdata2(0,0) = 1 and ubound(rsdata2,2) = 0 Then
				' If they are only connected or associated to 1 entity - then launch that entity
				' The UBound is to ensure they are only connected to 1 resource as well

					If rsdata4(12,0) and rsdata4(11,0) and rsdata4(20,0) Then
					' Ensure they are enabled
						'Response.Write "<script>alert('" & rsdata4(11,0) & "');</script>"

						Call SetSession("USERIPADDRESS",Request.ServerVariables("REMOTE_HOST"))
						Call SetSession("ug",rsdata1(0,0))
						Call SetSession("ut",rsdata1(5,0))
						Call SetSession("mi",rsdata1(6,0))
						Call SetSession("mf",Trim(rsdata1(1,0)))
						Call SetSession("ml",Trim(rsdata1(2,0)))
						Call SetSession("me",rsdata4(15,0))
						Call SetSession("eg",rsdata4(1,0))
						Call SetSession("en",rsdata4(5,0))
						Call SetSession("el",rsdata4(16,0))
						Call SetSession("et",rsdata4(3,0))

						If Not Application("ds") = "" Then
							Call SetSession("ds",Application("ds"))
						Else
							Call SetSession("ds",rsdata4(7,0))
						End If
						If rsdata4(22,0) Then
							If Application("Use_SSL") Then
								Call SetSession("webHTTP","https://")
							Else
								Call SetSession("webHTTP","http://")
							End If
						Else
							Call SetSession("webHTTP","http://")
						End If
						If Not Application("db") = "" Then
							Call SetSession("db",Application("db"))
						Else
							If UCase(Trim(rsdata4(3,0))) = "DEMO" Then
								' HARDCODE UNTIL WE GET ALL DEMO ENTITIES BUILT 9/17/2003
								'Call SetSession("db",Trim(rsdata4(3,erow)) & Trim(rsdata4(4,erow)))
								Call SetSession("db",Trim(rsdata4(3,0)) & "AMI")
							Else
								Call SetSession("db",Trim(rsdata4(3,0)) & Trim(rsdata4(4,0)))
							End If
						End If
						Call SetSession("an",Trim(rsdata4(18,0)))
						Call SetSession("ae",rsdata4(19,0))
						Call SetSession("rl",Trim(rsdata4(21,0)))
						If UCase(Trim(rsdata1(5,0))) = "DEMO" or UCase(Trim(rsdata1(5,0))) = "SALES" Then
							Call SetSession("dm","Y")
						Else
							Call SetSession("dm","N")
						End If
						Call SetSession("ut",UCase(Trim(rsdata1(5,0))))
					    Call SetSession("lc",rsdata4(24,0))

						' Set the Session Timeout Factor
						If Application("Onsite") = 1 Then
							If Not Application("SessionTimeout") = "0" Then
								Call SetSession("tf",Application("SessionTimeout"))
							Else
								Call SetSession("tf",rsdata4(23,0))
							End If
						Else
							If Not NullCheck(rsdata4(23,0)) = "" Then
								If rsdata4(23,0) > 0 Then
									Call SetSession("tf",rsdata4(23,0))
								Else
									Call SetSession("tf",Application("SessionTimeout"))
								End If
							Else
								Call SetSession("tf",Application("SessionTimeout"))
							End If
							If GetSession("tf") = 0 or GetSession("tf") = "" Then
								Call SetSession("tf","20")
							End If
						End If

						Call SetSession("tf","9999")

						SessionID = CreateSessionAndContext(Trim(rsdata4(0,0)), _
									rsdata4(1,0), _
									Request.ServerVariables("REMOTE_HOST"), _
									GetSession("tf"), _
									"TVO", _
									"", _
									rsdata1(0,0))

						If Not SessionID = "" Then
							LaunchIt = True
						Else
							errortext = "A session could not be established. Please try again."
						End If

					Else
						If Not rsdata4(11,0) Then
							errortext = "The " & Trim(rsdata4(5,0)) & " Company / Organization is currently disabled. You can contact the Maintenance Administrator (" & Trim(rsdata4(18,0)) & ") at <a href=""mailto:" & Trim(rsdata4(19,0)) & """>" & Trim(rsdata4(19,0)) & "</a>."
						Elseif Not rsdata4(20,0) Then
							errortext = "The " & Trim(rsdata4(5,0)) & " Company / Organization is currently disabled from the use of Maintenance Connection. You can contact the Maintenance Administrator (" & Trim(rsdata4(18,0)) & ") at <a href=""mailto:" & Trim(rsdata4(19,0)) & """>" & Trim(rsdata4(19,0)) & "</a>."
						Elseif Trim(rsdata1(5,0)) = "DEMO" Then
							errortext = "Your trial period has ended for using the " & Trim(rsdata4(5,0)) & " Company / Organization. Please contact our Sales Department at <a href=""mailto:sales@maintenanceconnection.com"">sales@maintenanceconnection.com</a>."
						Else
							errortext = "You are not currently approved to use the " & Trim(rsdata4(5,0)) & " Company / Organization. You can contact the Maintenance Administrator (" & Trim(rsdata4(18,0)) & ") at <a href=""mailto:" & Trim(rsdata4(19,0)) & """>" & Trim(rsdata4(19,0)) & "</a>."
						End If
					End If

				Else

					Call SetSession("USERIPADDRESS",Request.ServerVariables("REMOTE_HOST"))
					Call SetSession("ug",rsdata1(0,0))
					Call SetSession("ut",rsdata1(5,0))
					Call SetSession("mi",rsdata1(6,0))
					Call SetSession("mf",Trim(rsdata1(1,0)))
					Call SetSession("ml",Trim(rsdata1(2,0)))
					If UCase(Trim(rsdata1(5,0))) = "DEMO" or UCase(Trim(rsdata1(5,0))) = "SALES" Then
						Call SetSession("dm","Y")
					Else
						Call SetSession("dm","N")
					End If
					Call SetSession("ut",UCase(Trim(rsdata1(5,0))))

					If Not NullCheck(rsdata4(23,0)) = "" Then
						If rsdata4(23,0) > 0 Then
							Call SetSession("tf",rsdata4(23,0))
						Else
							Call SetSession("tf",Application("SessionTimeout"))
						End If
					Else
						Call SetSession("tf",Application("SessionTimeout"))
					End If
					If GetSession("tf") = 0 or GetSession("tf") = "" Then
						Call SetSession("tf","20")
					End If

					Call SetSession("tf","9999")

					If Not Request("SearchBy") = "" Then
						If UCase(Request("SearchBy")) = "NONE" Then
							Call SetSession("SearchBy","")
						Else
							Call SetSession("SearchBy",Request("SearchBy"))
						End If
						SetContext
						EntitySelect
						Response.End
					Else

						SessionID = CreateSessionAndContext(p_resource, _
									rsdata4(1,0), _
									Request.ServerVariables("REMOTE_HOST"), _
									GetSession("tf"), _
									"TVO", _
									SessionVars, _
									rsdata1(0,0))

						If Not SessionID = "" Then

							Response.Cookies("m") = txtmembername
							Response.Cookies("m").Path   = "/"
							Response.Cookies("m").Expires = Now()

							'Response.Cookies("p") = txtpassword
							'Response.Cookies("p").Path   = "/"
							'Response.Cookies("p").Expires = Now()

							EntitySelect
							Response.End
						Else
							If errortext = "" Then
								errortext = "A session could not be established. Please try again."
							End If
						End If

					End If

				End If

			Else
				Select Case Retval2

					 Case -400  ' incorrect parameter which means the parameters were null
							errortext = "There appears to be a database problem. Please try again later."
			         Case -2300 ' now not found cannot find row in table users
							errortext = "There is an internal data problem (-2300). Please contact your Maintenance Administrator."
			         Case -2400 ' row found in users table, but no resource/container sets found
							errortext = "You have successfully entered your Member ID and Password but you are not connected to a Company / Resource. Please contact your Maintenance Administrator."

				End Select
			End If

		Else

			Call CloseObj( con )

			Select Case CLng(rsdata1(0,0))

				Case -400  ' incorrect parameters passed
					errortext = "There appears to be a database problem. Please try again later."

				Case -500  ' business logic error which means there is a problem in the stored procedure.
					errortext = "There appears to be a database problem (-500). Please try again later."

				Case -600  ' invalid logon and password combination which means the logon exists in users table.

				Case -800  ' correct logon and password provided, however the is_enabled row in users is set to false.
					errortext = "We are sorry, but you have been disabled from access to the Maintenance Connection."

				Case -880  ' correct logon and password provided, however the user has never gone through 1st time logon.
					errortext = "Before logging on, please use the Sign Up process by clicking the Sign Up link."

				Case -890  ' correct logon and password provided, however the user must change their expired password.
					errortext = "Your password has expired. Please contact your Maintenance Administrator."

				Case -2300 ' row not found which means no row exists in the users table

			End Select

		End If

	Else

		If Application("loginsdisabled") Then
			errortext = Application("DownText")
		Else
			errortext = "Either the Member ID or Password has been left blank. Please try again."
		End If

	End If

	If launchit Then

		Response.Cookies("m") = txtmembername
		Response.Cookies("m").Path   = "/"
		Response.Cookies("m").Expires = Now()

		'Response.Cookies("p") = txtpassword
		'Response.Cookies("p").Path   = "/"
		'Response.Cookies("p").Expires = Now()

		InitConnection
		MainMenu
		Response.End

	Else

		If ErrorText = "" Then
			errortext = "We were unable to verify your Member ID and/or Password."
		End If

	End If

	Call OutputWAPError(ErrorText)

End Sub

'====================================================================================================================================

Sub EntitySelect()

	pagesize = 100

	If Request("pagepos") = "" Then
		pagepos = 0
	Else
		pagepos = CLng(Request("pagepos"))
	End If
	errortext = "There appears to be a database problem. Please try again later."
	Retval = -1
	rscounter = 1
	retvalcounter = 1

	If GetSession("ug") = "" Then
		ErrorText = "A User was not found. Please try again."
		Call OutputWAPError(ErrorText)
	End If

	Call sp_start("PSPT_GetUserRights","NEWCONNECTION")
	cmd.Parameters.Append cmd.CreateParameter("@p_userGuid",  adVarchar, adParamInput, 36, GetSession("ug"))
	cmd.Parameters.Append cmd.CreateParameter("@p_resource",  adVarchar, adParamInput, 36, p_resource)
	If Not GetSession("SearchBy") = "" Then
	'cmd.Parameters.Append cmd.CreateParameter("@p_searchby",  adVarChar, adParamInput, 25, GetSession("SearchBy"))
	End If
	RS.Open cmd, , adOpenForwardOnly, adLockReadOnly
	Call CheckForError("A",cmd.ActiveConnection,RS)
	Call scattertoarrays("KILLONCOMPLETE")

	If Not Retval = 0 Then

		Select Case Retval

			Case -400  ' incorrect parameter which means the parameters were null
			   	errortext = "There appears to be a database problem. Please try again later."
		    Case -2300 ' now not found cannot find row in table users
			   	errortext = "There is an internal data problem (-2300). Please contact your Maintenance Administrator."
		    Case -2400 ' row found in users table, but no resource/container sets found
			   	errortext = "There is an internal data problem (-2400). Please contact your Maintenance Administrator."
			Case Else
			   	errortext = "There was a problem loading the list of companies you have access to. Please try again."

		End Select

		Call OutputWAPError(ErrorText)

	End If

	ei = Request("ei")

    'If rsdata3(1,ei) = "5938C631-B16C-11D5-9E09-00104BCA94F0" Then
    'Redirect MC to MCSupport
    If Not ei = "" and Trim(UCase(rsdata3(4,ei))) = "MC" Then
        Dim p
        For p = 0 to UBound(rsdata3,2)
            If Trim(UCase(rsdata3(4,p))) = "MCSUPPORT" Then
                ei = p
                Exit For
            End If
        Next
    End If

	If Trim(UCase(rsdata3(4,ei))) = "MCSUPPORT" and Trim(UCase(rsdata3(3,ei))) = "ADM" Then
        Call SetSession("mcsupport","Y")
	End If

	If Not ei = "" Then

		' Bounce this stuff in the session and redirect to mc_login.htm

		If rsdata3(12,ei) and rsdata3(11,ei) and rsdata3(20,ei) Then
		' Ensure they are enabled

			Call SetSession("me",rsdata3(15,ei))
			Call SetSession("eg",rsdata3(1,ei))
			Call SetSession("en",rsdata3(5,ei))
			Call SetSession("el",rsdata3(16,ei))
			Call SetSession("et",rsdata3(3,ei))

			If Not Application("ds") = "" Then
				Call SetSession("ds",Application("ds"))
			Else
				Call SetSession("ds",rsdata3(7,ei))
			End If
			If rsdata3(22,ei) Then
				If Application("Use_SSL") Then
					Call SetSession("webHTTP","https://")
				Else
					Call SetSession("webHTTP","http://")
				End If
			Else
				Call SetSession("webHTTP","http://")
			End If
			If Not Application("db") = "" Then
				Call SetSession("db",Application("db"))
			Else
				Call SetSession("db",Trim(rsdata3(3,ei)) & Trim(rsdata3(4,ei)))
			End If
			Call SetSession("an",Trim(rsdata3(18,ei)))
			Call SetSession("ae",rsdata3(19,ei))
			Call SetSession("rl",Trim(rsdata3(21,ei)))
		    Call SetSession("lc",rsdata3(24,ei))
		    ' Multi-Entity = True
		    Call SetSession("ed","Y")

			' Set the Session Timeout Factor
			'If Application("Onsite") = 1 Then
			'	If Not Application("SessionTimeout") = "0" Then
			'		Call SetSession("tf",Application("SessionTimeout"))
			'	Else
			'		Call SetSession("tf",rsdata3(23,ei))
			'	End If
			'Else
				If Not NullCheck(rsdata3(23,ei)) = "" Then
					If rsdata3(23,ei) > 0 Then
						Call SetSession("tf",rsdata3(23,ei))
					Else
						Call SetSession("tf",Application("SessionTimeout"))
					End If
				Else
					Call SetSession("tf",Application("SessionTimeout"))
				End If
				If GetSession("tf") = 0 or GetSession("tf") = "" Then
					Call SetSession("tf","20")
				End If

				Call SetSession("tf","9999")

			'End If

			InitConnection
			MainMenu
			Response.End

		End If
	End If

	Dim resrows,resrowcounter
	choiceindex = 0
	tabindex = 1

	Call StartMobileDocument("Select Company")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=entitysearch&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:6px;'><img src='images/icons/48/Find.png' alt='Search' title='Search' style='border:none; cursor:pointer; width:42px; height:42px;' /></div>"
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Logout' title='Logout' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, False)
	%>
	<div class='Font2' style='text-align:left; width:100%;'>
	<%
	If rsdata1recs Then

		resrows = ubound(rsdata1,2)

		If resrows > 0 Then
		' There is more than just the Maintenance Connection Application they have access to


			FOR resrowcounter= 0 TO resrows
				If (pagepos > 0 and resrowcounter > 0) or (pagepos = 0) Then
					If Not GetSession("SearchBy") = "" and UCase(rsdata2(6,resrowcounter)) = "MAINTENANCE CONNECTION ADMINISTRATION" Then
					Else
					%>
						<div class='Font2' style='padding:5px;font-size:16pt; font-weight:normal; background-color:#fddcbd; color:#323232;'><% =rsdata2(6,resrowcounter) %></div>
					<%
					End If
				End If
				%>
				<%Call OutputEntities(Trim(rsdata1(1,resrowcounter)),resrowcounter,choiceindex,tabindex,pagepos,pagesize)%>
			<%
			NEXT

		Else
			%><div class='Font2' style='font-size:18pt;padding-bottom:20px;'>What company would you like to work with?</div>
			<%
			Call OutputEntities(Trim(rsdata1(1,0)),0,choiceindex,tabindex,pagepos,pagesize)
		End If

	End If
	%>
	<div style='width:100%;padding-top:20px; padding-bottom:10px;'>
	<div style='float:right;padding-right:20px; padding-top:20px;'>
	<%
	If PagePos >= PageSize Then
	%><a href="default.asp?card=entityselect&amp;pagepos=<% =CStr(PagePos-PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Left Green.png' alt='Back' title='Back' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	If TabIndex = PageSize+1 Then
	%><a href="default.asp?card=entityselect&amp;pagepos=<% =CStr(PagePos+PageSize) %>&amp;s=<% =SessionID %>"><img src='images/icons/48/Navigation 1 Right Green.png' alt='Back' title='Back' style='border:none; cursor:pointer;' /></a>
	<%
	End If
	%></div><div style="clear:both;"></div>
	</div>

	<div style='padding-bottom:20px; width:100%; text-align:center; display:none;'>
		<input style="width:300px;" type="button" name="Search" value="Search" onclick="self.location.href='default.asp?card=entitysearch&s=<% =SessionID %>';"/>
		<input style="width:300px;" type="button" name="LogOff" value="Log-Off" onclick="self.location.href='default.asp?card=logoff&s=<% =SessionID %>';"/>
	</div>
	<%
	HTMLbuttonsEnd
	SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Sub OutputEntities(resource,rescount,ByRef choiceindex,ByRef tabindex,pagepos,pagesize)


	If rsdata3recs Then

		Dim cols,rows,field,rowcounter

		cols=ubound(rsdata3,1)
		rows=ubound(rsdata3,2)
		Dim iCount, altStyle
		iCount = 0

		FOR rowcounter= 0 TO rows
			if iCount Mod 2 = 0 Then
				altStyle = "#f0f7fe"
			else
				altStyle = "#ffffff"
			end if

		   If Trim(rsdata3(0,rowcounter)) = resource Then
				'Response.Write choiceindex & "-" & pagepos & "<br/>"
			    If choiceindex >= pagepos Then
					If (rsdata3(11,rowcounter) = False) or (rsdata3(20,rowcounter) = False) Then
						'do not output if container is disabled or container/resource is disabled
						'choiceindex = choiceindex - 1

					'ElseIf (Not UCase(Left(rsdata3(5,rowcounter),1)) = UCase(GetSession("SearchBy"))) and (Not GetSession("SearchBy") = "") Then
						'choiceindex = choiceindex - 1
					Else
						If rsdata3(12,rowcounter) and rsdata3(11,rowcounter) and rsdata3(20,rowcounter) Then
						' If the User is enabled for the Resource / Container and the Container is enabled and the Container is enabled for the Resource
						%>
						<div class='Font1' style='padding:5px;background-color:<%=altStyle%>; font-size:14pt; cursor:pointer; width:100%; clear:both;' onclick="location.href = 'default.asp?card=entityselect&amp;ei=<% =CStr(choiceindex) %>&amp;s=<% =SessionID %>';">
							<div style='float:left;'><img src="images/icons/48/Hospital 2.png" alt="" title=""></div>
							<div class='Font1 RowElisp' style='padding-left:10px;float:left;font-size:14pt; max-width:75%; cursor:pointer;'>
								<% =WAPValidate(mcval(rsdata3(5,rowcounter))) %>
						  	</div>
						  	<div style='clear:both;'></div>
						</div>
						<%
						Else
							If False Then
								If (rsdata3(11,rowcounter) = False) or (rsdata3(20,rowcounter) = False) Then
								' The Entity is disabled from all Resources or the Entity is disabled for this Resource
								%>
								<% =mcval(rsdata3(5,rowcounter)) %> (Disabled)
								<br/>
								<%
								ElseIf GetSession("ut") = "DEMO" Then
								%>
								<% =mcval(rsdata3(5,rowcounter)) %> (Trial Period Ended)
								<br/>
								<%
								Else
								%>
								<% =mcval(rsdata3(5,rowcounter)) %> (Not Approved)
								<br/>
								<%
								End If
							End If
						End If
						tabindex = tabindex + 1
						If tabindex > pagesize Then
							Exit For
						End If
					End If
				End If
				choiceindex = choiceindex + 1
			End If
			iCount = iCount + 1
		NEXT

	End If

End Sub

'====================================================================================================================================

Sub EntitySearch()
	searchby = UCase(Request("searchby"))
	Call StartMobileDocument("Company Search")
	headerString = GetBackButton("history.back(1);")
	headerString = headerString & "<div onclick=""self.location.href='default.asp?card=logoff&s=" & SessionID & "';"" style='float:left; padding-right:10px; padding-top:3px;'><img src='images/icons/48/Logout.png' alt='Logout' title='Logout' style='border:none; cursor:pointer;' /></div><div style='clear:both;'></div>"
	Call IPadHeader(headerString, False)

		%>
		<p align="center">
		<b>Company Search</b>
		</p>
		<p align="center">
		<% Select Case searchby
		   Case ""
		%>
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=NONE">ALL</a>&nbsp;
		<a href="default.asp?card=EntitySearch&amp;s=<% =SessionID %>&amp;searchby=a">A-F</a>&nbsp;
		<a href="default.asp?card=EntitySearch&amp;s=<% =SessionID %>&amp;searchby=g">G-K</a>&nbsp;
		<a href="default.asp?card=EntitySearch&amp;s=<% =SessionID %>&amp;searchby=l">L-P</a>&nbsp;
		<a href="default.asp?card=EntitySearch&amp;s=<% =SessionID %>&amp;searchby=q">Q-U</a>&nbsp;
		<a href="default.asp?card=EntitySearch&amp;s=<% =SessionID %>&amp;searchby=v">V-Z</a>&nbsp;
		<% Case "A" %>
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=NONE">ALL</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=a">A</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=b">B</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=c">C</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=d">D</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=e">E</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=f">F</a>&nbsp;
		<% Case "G" %>
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=g">G</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=h">H</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=i">I</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=j">J</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=k">K</a>&nbsp;
		<% Case "L" %>
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=l">L</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=m">M</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=n">N</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=o">O</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=p">P</a>&nbsp;
		<% Case "Q" %>
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=q">Q</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=r">R</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=s">S</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=t">T</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=u">U</a>&nbsp;
		<% Case "V" %>
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=v">V</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=w">W</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=x">X</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=y">Y</a>&nbsp;
		<a href="default.asp?card=authenticate&amp;s=<% =SessionID %>&amp;searchby=z">Z</a>&nbsp;
		<% End Select %>
		</p>

	<%
		SetContext
	EndWMLDocument

End Sub

'====================================================================================================================================

Function getResourceItem(resitem,resourceguid)

	Dim resrowcounter
	getResourceItem = ""

	FOR resrowcounter= 0 TO  ubound(rsdata2,2)
		If rsdata2(0,resrowcounter) = resourceguid Then
			Select Case resitem
				Case "RESOURCECODE"
					getResourceItem = Trim(rsdata2(5,resrowcounter))
				Case "RESOURCEURL"
					getResourceItem = Trim(rsdata2(9,resrowcounter))
			End Select
			Exit For
		End If
	NEXT

End Function

'====================================================================================================================================

Function InitConnection()

	Dim db,rs,rs_modlist,rs_access,sql,sql2
	Dim errormessage,errormessage2,returnmessage

	errormessage = ""
	errormessage2 = ""

	If GetSession("UT") = "MC" Then
		Call SetSession("IsAdmin","Y")
	End If

	If (GetSession("webHTTP") = "") Then
		If Trim(UCase(Request.ServerVariables("HTTPS"))) = "ON" Then
			Call SetSession("webHTTP","https://")
		Else
			Call SetSession("webHTTP","http://")
		End If
	End If

	Set db = New ADOHelper

	If GetSession("LaborPK") = "" Then
		sql = "SELECT L.LaborPK, L.LaborID, L.LaborName, L.Access, L.AccessGroupPK, L.Initials, L.PhoneWork, L.RepairCenterPK, L.RepairCenterID, L.RepairCenterName, L.ShopPK, L.ShopID, L.ShopName, L.CraftPK, L.CraftID, L.CraftName, RC.StockRoomPK, RC.StockRoomID, RC.StockRoomName, RC.ToolRoomPK, RC.ToolRoomID, RC.ToolRoomName " + _
			  "FROM Labor L WITH (NOLOCK) Left Outer Join RepairCenter RC WITH (NOLOCK) ON RC.RepairCenterPK = L.RepairCenterPK " + _
			  "WHERE (L.User_Guid = '" & GetSession("UG") & "')"
	Else
		sql = "SELECT L.LaborPK, L.LaborID, L.LaborName, L.Access, L.AccessGroupPK, L.Initials, L.PhoneWork, L.RepairCenterPK, L.RepairCenterID, L.RepairCenterName, L.ShopPK, L.ShopID, L.ShopName, L.CraftPK, L.CraftID, L.CraftName, RC.StockRoomPK, RC.StockRoomID, RC.StockRoomName, RC.ToolRoomPK, RC.ToolRoomID, RC.ToolRoomName " + _
			  "FROM Labor L WITH (NOLOCK) Left Outer Join RepairCenter RC WITH (NOLOCK) ON RC.RepairCenterPK = L.RepairCenterPK " + _
			  "WHERE (L.LaborPK = '" & GetSession("LaborPK") & "')"
	End If

	Set rs = db.RunSQLReturnRS(sql,"")

	'Call OutputWAPError(sql)
	'Call OutputWAPError(db.GetConnectionString())
	'Call OutputWAPError("HERE")


	If DemoMode() Then

		If Not db.dok Then
			errortext = "Your login failed to initialize: " & db.derror
			Set db = Nothing
			Call OutputWAPError(ErrorText)
		Else
			If rs.eof Then

				Dim laborid, laborname, craftpk, craftid, craftname

				laborid = UCase(Mid(GetSession("MF"),1,1) + GetSession("ML"))
				laborname = GetSession("ML") + ", " + GetSession("MF")

				sql2 = _
				"INSERT INTO Labor " + _
				"                      (LaborID, LaborName, FirstName, MiddleName, LastName, Initials, CraftPK, CraftID, CraftName, LaborType, LaborTypeDesc, Active, AccessGroupPK, " + _
				"                      AccessGroupID, AccessGroupName, Access, RepairCenterPK, RepairCenterID, RepairCenterName, ShopPK, ShopID, ShopName, Email, EmailNotify, EmailNotifyDesc, CostRegular, " + _
				"                      CostOvertime, CostOther, ChargePercentage, ChargeRate, HireDate, User_Guid, role_code, role_desc, memberid, DemoLaborPK, RowVersionInitials, RowVersionAction) " + _
				"SELECT     '" & laborid & "' AS Expr1,'" & laborname & "' AS Expr2,'" & GetSession("MF") & "' AS Expr3,'' AS Expr4,'" & GetSession("ML") & "' AS Expr5,'" & Mid(GetSession("MF"),1,1) & Mid(GetSession("ML"),1,1) & "' AS Expr6," & _
				"			craftpk,craftid,craftname,LaborType, LaborTypeDesc, Active, AccessGroupPK, " & _
				"                      AccessGroupID, AccessGroupName, 1, RepairCenterPK, RepairCenterID, RepairCenterName, ShopPK, ShopID, ShopName,'" & GetSession("ME") & "' AS Expr7, EmailNotify, EmailNotifyDesc, CostRegular, " & _
				"                      CostOvertime, CostOther, ChargePercentage, ChargeRate,getDate() AS Expr8,'" & GetSession("UG") & "' AS Expr9,'MGR','MRO WorkCenter','" &  GetSession("MI") & "', 0 AS Expr10, 'MC' AS Expr11, 'INSERT' AS Expr12 " & _
				"FROM Labor WITH (NOLOCK) " & _
				"WHERE LaborID = "

				If UCase(GetSession("ET")) = "DEMO" Then
					sql2 = sql2 & "'BSQUIRES' "
				Else
					sql2 = sql2 & "'A' "
				End If

				If Not db.RunSQL(sql2,"") Then

					db.derror = ""
					db.dok = True

					laborid = UCase(Mid(GetSession("MF"),1,1) + GetSession("ML")) + "-" + CStr(GetRandomNumber(1,1000))

					sql2 = _
					"INSERT INTO Labor " + _
					"                      (LaborID, LaborName, FirstName, MiddleName, LastName, Initials, CraftPK, CraftID, CraftName, LaborType, LaborTypeDesc, Active, AccessGroupPK, " + _
					"                      AccessGroupID, AccessGroupName, Access, RepairCenterPK, RepairCenterID, RepairCenterName, ShopPK, ShopID, ShopName, Email, EmailNotify, EmailNotifyDesc, CostRegular, " + _
					"                      CostOvertime, CostOther, ChargePercentage, ChargeRate, HireDate, User_Guid, role_code, role_desc, memberid, DemoLaborPK, RowVersionInitials, RowVersionAction) " + _
					"SELECT     '" & laborid & "' AS Expr1,'" & laborname & "' AS Expr2,'" & GetSession("MF") & "' AS Expr3,'' AS Expr4,'" & GetSession("ML") & "' AS Expr5,'" & Mid(GetSession("MF"),1,1) & Mid(GetSession("ML"),1,1) & "' AS Expr6," & _
					"			craftpk,craftid,craftname,LaborType, LaborTypeDesc, Active, AccessGroupPK, " & _
					"                      AccessGroupID, AccessGroupName, 1, RepairCenterPK, RepairCenterID, RepairCenterName, ShopPK, ShopID, ShopName,'" & GetSession("ME") & "' AS Expr7, EmailNotify, EmailNotifyDesc, CostRegular, " & _
					"                      CostOvertime, CostOther, ChargePercentage, ChargeRate,getDate() AS Expr8,'" & GetSession("UG") & "' AS Expr9,'MGR','MRO WorkCenter','" & GetSession("MI") & "', 0 AS Expr10, 'MC' AS Expr11, 'INSERT' AS Expr12 " & _
					"FROM Labor WITH (NOLOCK) " & _
					"WHERE LaborID = "

					If UCase(GetSession("ET")) = "DEMO" Then
						sql2 = sql2 & "'BSQUIRES' "
					Else
						sql2 = sql2 & "'A' "
					End If

					If Not db.RunSQL(sql2,"") then
						Set db = Nothing
						errortext = "Your login failed to initialize."
						Call OutputWAPError(ErrorText)
					End If

				End If

				Set rs = db.RunSQLReturnRS(sql,"")

				If Not db.dok Then
					Set db = Nothing
					errortext = "Your login failed to initialize."
					Call OutputWAPError(ErrorText)
				End If

				If Not rs.eof Then

					If False Then
						sql2 = _
						"INSERT INTO LaborAsset " + _
						"                      (LaborPK, AssetPK, RowVersionUserPK, RowVersionInitials, RowVersionDate) " + _
						"SELECT     '" & rs("LaborPK").Value & "', LaborAsset.AssetPK, Null, 'MC', getDate() " + _
						"FROM         LaborAsset WITH (NOLOCK) INNER JOIN " + _
						"                      Labor WITH (NOLOCK) ON LaborAsset.LaborPK = Labor.LaborPK " + _
						"WHERE     Labor.LaborID = "

						If UCase(GetSession("ET")) = "DEMO" Then
							sql2 = sql2 & "'BSQUIRES' "
						Else
							sql2 = sql2 & "'A' "
						End If

						If Not db.RunSQL(sql2,"") then
							Set db = Nothing
							errortext = "Your login failed to initialize."
							Call OutputWAPError(ErrorText)
						End If
					End If

					sql2 = _
					"INSERT INTO LaborAddressBook " + _
					"          (LaborPK, EmailName, EmailAddress) " + _
					"			VALUES     (" + CStr(rs("LaborPK").Value) + ",'" + rs("LaborName").Value + "','" + GetSession("ME") + "') "

					If Not db.RunSQL(sql2,"") then
						Set db = Nothing
						errortext = "Your login failed to initialize."
						Call OutputWAPError(ErrorText)
					End If

				End If

			End If
		End If

	End If

	If Not db.dok Then
		errortext = "Your login failed to initialize: "	& db.derror
		Call OutputWAPError(ErrorText)
	ElseIf rs.eof Then
		errortext = "Your login was authenticated but a labor record was not found."
		Call OutputWAPError(ErrorText)
	ElseIf rs("access") = False Then
		errortext = "Your login failed to initialize due to an access violation."
		Call OutputWAPError(ErrorText)
	ElseIf NullCheck(rs("AccessGroupPK")) = "" Then
		errortext = "Your login failed to initialize because an access group was not found."
		Call OutputWAPError(ErrorText)
	End If

	If Not errormessage = "" Then
		errortext = errormessage
		Call OutputWAPError(ErrorText)
	End If

	Call SetSession("RCPK",NullCheck(rs("RepairCenterPK")))
	Call SetSession("RCID",NullCheck(rs("RepairCenterID")))
	Call SetSession("RCNM",NullCheck(rs("RepairCenterName")))

	Call SetSession("SHPK",NullCheck(rs("ShopPK")))
	Call SetSession("SHID",NullCheck(rs("ShopID")))
	Call SetSession("SHNM",NullCheck(rs("ShopName")))

	' Below Added for PDA Version
	'----------------------------------------------------------------------------------
	Dim prefvalue, prefdesc, prefpk, rs2, shoppk
	If GetPreference(db,True,GetSession("RCPK"),"SU_DefaultShop",prefvalue, prefdesc, prefpk) Then
		shoppk = prefpk
		sql = "SELECT * FROM Shop WITH (NOLOCK) WHERE ShopPK = '" & NullCheck(Request("ShopPK")) & "' "
		Set rs2 = db.RunSQLReturnRS(sql,"")
		Call CheckDB(db)
		If Not rs2.eof Then
			Call SetSession("SHPK",rs2("ShopPK"))
			Call SetSession("SHID",rs2("ShopID"))
			Call SetSession("SHNM",rs2("ShopName"))
		End If
	End If
	'----------------------------------------------------------------------------------

	Call SetSession("SRPK",NullCheck(rs("StockRoomPK")))
	Call SetSession("SRID",NullCheck(rs("StockRoomID")))
	Call SetSession("SRNM",NullCheck(rs("StockRoomName")))

	Call SetSession("TMPK",NullCheck(rs("ToolRoomPK")))
	Call SetSession("TMID",NullCheck(rs("ToolRoomID")))
	Call SetSession("TMNM",NullCheck(rs("ToolRoomName")))

	Call SetSession("AGPK",NullCheck(rs("AccessGroupPK")))

	Call SetSession("USERPK",NullCheck(rs("LaborPK")))
	Call SetSession("USERID",NullCheck(rs("LaborID")))
	Call SetSession("USERNAME",NullCheck(rs("LaborName")))
	Call SetSession("USERINITIALS",NullCheck(rs("Initials")))

	Call SetSession("CraftPK",NullCheck(rs("CraftPK")))
	Call SetSession("CraftID",NullCheck(rs("CraftID")))
	Call SetSession("CraftName",NullCheck(rs("CraftName")))

	' GET ACCESS GROUP INFO
	'===================================================================================
	Set rs_access = db.RunSPReturnMultiRS("MC_GetLaborAccess",Array(Array("@accessgroupPK", adInteger, adParamInput, 4, rs("AccessGroupPK").Value)),"")
	If Not db.dok Then
		Set db = Nothing
		errortext = "Your login failed to initialize."
		Call OutputWAPError(ErrorText)
	End If

	Dim RCDENY
	RCDENY = ""
	Do Until rs_access.eof
		RCDENY = RCDENY & rs_access("RepairCenterPK")
		rs_access.movenext
		If Not rs_access.eof Then
			RCDENY = RCDENY & ","
		End If
	Loop

	Call SetSession("AGPK",rs("AccessGroupPK").value)
	Call SetSession("RCDENY",RCDENY)

	Set rs_access = rs_access.NextRecordset

	If rs_access("WOAuth") Then
		Call SetSession("WOAuth","Y")
	Else
		Call SetSession("WOAuth","N")
	End If
	Call SetSession("WOAuthAmount",rs_access("WOAuthAmount"))
	Call SetSession("WOAuthReq",rs_access("WOAuthReq"))
	Call SetSession("WOAuthReqAmount",rs_access("WOAuthReqAmount"))

	If rs_access("PDAuth") Then
		Call SetSession("PDAuth","Y")
	Else
		Call SetSession("PDAuth","N")
	End If
	Call SetSession("PDAuthAmount",rs_access("PDAuthAmount"))
	Call SetSession("PDAuthReq",rs_access("PDAuthReq"))
	Call SetSession("PDAuthReqAmount",rs_access("PDAuthReqAmount"))

	If rs_access("POAuth") Then
		Call SetSession("POAuth","Y")
	Else
		Call SetSession("POAuth","N")
	End If
	Call SetSession("POAuthAmount",rs_access("POAuthAmount"))
	Call SetSession("POAuthReq",rs_access("POAuthReq"))
	Call SetSession("POAuthReqAmount",rs_access("POAuthReqAmount"))

	If rs_access("PJAuth") Then
		Call SetSession("PJAuth","Y")
	Else
		Call SetSession("PJAuth","N")
	End If
	Call SetSession("PJAuthAmount",rs_access("PJAuthAmount"))
	Call SetSession("PJAuthReq",rs_access("PJAuthReq"))
	Call SetSession("PJAuthReqAmount",rs_access("PJAuthReqAmount"))

	Set rs_access = rs_access.NextRecordset

	Set rs_modlist = db.RunSPReturnRS("MC_GetModuleList","","")
	If Not db.dok Then
		Set db = Nothing
		errortext = "Your login failed to initialize."
		Call OutputWAPError(ErrorText)
	End If

	Call CloseObj(rs)

End Function
%>