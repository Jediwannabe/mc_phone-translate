<%
Function CreateSessionAndContext(ByVal p_resource, _
                                 ByVal p_containerguid, _
                                 ByVal p_IPAddress, _
                                 ByVal p_timeout, _
                                 ByVal p_formatCode, _
                                 ByVal p_context, _
                                 ByVal p_userguid)
	Dim pdb, OutArray
	
    CreateSessionAndContext = ""

	Set pdb = New ADOHelper 
	pdb.oledbstr = Application("app_dsn")		
		                                                    
    Call pdb.RunSP("SES_CreateSessionAndContext", _
    Array( _
    Array("RETURN_VALUE", adInteger, adParamReturnValue, 4, 0), _
    Array("@p_IPAddress", adVarChar, adParamInput, 15, p_IPAddress), _
    Array("@p_timeout", adInteger, adParamInput, 4, p_timeout), _
    Array("@p_containerguid", adVarChar, adParamInput, 36, p_containerguid), _
    Array("@p_resource", adVarChar, adParamInput, 36, p_resource), _
    Array("@p_formatCode", adChar, adParamInput, 4, p_formatCode), _
    Array("@p_context", adVarChar, adParamInput, 7500, p_context), _
    Array("@p_userguid", adVarChar, adParamInput, 36, p_userguid), _
    Array("@p_sessionid", adVarChar, adParamOutput, 100, ""), _
    Array("@p_errortext", adVarChar, adParamOutput, 1000, "") _
    ),OutArray)
	
	If pdb.dok Then               
		If OutArray(0) = 0 Then		
            CreateSessionAndContext = OutArray(8)
        Else
			If OutArray(9) = "" Then
				errortext = "There was a problem establishing your Session ID. Please try again."
			Else
				errortext = "There was a problem establishing your Session ID. (" & OutArray(9) & ") Please try again."
			End If
        End If
    Else
		errortext = pdb.derror
    End If
                                 
End Function

Function SetContext()
                                     
	Dim pdb, OutArray

	If SessionID = "" Then
		SetContext = False
		Exit Function
	End If
		
    SetContext = True

	Set pdb = New ADOHelper 
	pdb.oledbstr = Application("app_dsn")		
		                                                    
    Call pdb.RunSP("SES_SetContext", _
    Array( _
    Array("RETURN_VALUE", adInteger, adParamReturnValue, 4, 0), _
    Array("@p_sessionKey", adVarChar, adParamInput, 60, SessionID), _
    Array("@p_resource", adVarChar, adParamInput, 36, p_resource), _
    Array("@p_context", adVarChar, adParamInput, 7500, SessionVars) _
    ),OutArray)

	If pdb.dok Then               
		If Not OutArray(0) = 0 Then		
		errortext = "There was a problem writing to your session: " & pdb.derror
        End If
    Else
		errortext = pdb.derror
	End If

    If Not SetContext Then
		Call OutputWAPError(errortext)
    End If
                                
End Function

Function GetContext()
                      
	Dim pdb, rs, OutArray

    If Trim(Request.QueryString("s")) <> "" Then
	    SessionID = Request.QueryString("s")
	Else
	    SessionID = ""
    End If
		
	If SessionID = "" Then
		GetContext = False
		Exit Function
	End If

	If UCase(Request.QueryString("card")) = "LOGOFF" Then	
		GetContext = True
		Exit Function
	End If
	
    GetContext = True

	Set pdb = New ADOHelper 
	pdb.oledbstr = Application("app_dsn")		
		                                                    
    Set rs = pdb.RunSPReturnRS("SES_GetContext", _
    Array( _
    Array("RETURN_VALUE", adInteger, adParamReturnValue, 4, 0), _
    Array("@p_sessionKey", adVarChar, adParamInput, 60, SessionID), _
    Array("@p_resource", adVarChar, adParamInput, 36, p_resource) _
    ),OutArray)
	
	If pdb.dok Then                               
		If OutArray(0) = 0 Then		
			If Not RS.EOF Then
		        SessionVars = RS.Fields("session_context")
			End If
		Else
			errortext = "Your session timed out or is invalid."
			GetContext = False
		End If
	    RS.Close
	Else
		errortext = "There was a problem reading your session: " & pdb.derror
		GetContext = False
    End If
    
    Set RS = Nothing
	    
    If Not GetContext Then
		Call OutputWAPError(errortext)
    End If
                                                  
End Function

Function DeleteSessionSQL()
                                
	Dim pdb, OutArray

    DeleteSessionSQL = False

	If SessionID = "" Then
		Call OutputWAPError(SessionID)
		DeleteSessionSQL = True
		Exit Function
	End If
	
	Set pdb = New ADOHelper 
	pdb.oledbstr = Application("app_dsn")		
		                                                                                      
    Call pdb.RunSP("SES_DeleteSession", _
    Array(_
    Array("RETURN_VALUE", adInteger, adParamReturnValue, 4, 0), _
    Array("@p_sessionKey", adVarChar, adParamInput, 60, SessionID) _
    ),OutArray)
  	
	If pdb.dok Then               
	
		If OutArray(0) = 0 Then	
			DeleteSessionSQL = True
        Else
			errortext = "There was a problem removing the Licensed Session. Please try again."
        End If

		pdb.CloseClientConnection
		Set pdb = Nothing  
        
    Else
		errortext = pdb.derror
    End If

    If Not DeleteSessionSQL Then
		Call OutputWAPError(errortext)
    End If
	                             
End Function
%>