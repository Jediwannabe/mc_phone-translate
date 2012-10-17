<%
Class ADOHelper

private AppConn,inTransaction,pdok,pderror,pisduplicate,pisdeletecolref,pwarn,pwarntext,poledbstr

property get con()
	   If isObject(AppConn) Then
		  Set con = AppConn
	   Else
		  con = null
	   End If
end property

property get oledbstr()
	oledbstr = poledbstr
end property

property get dok()
	dok = pdok
end property

property get isduplicate()
	isduplicate = pisduplicate
end property

property get isdeletecolref()
	isdeletecolref = pisdeletecolref
end property

property get warn()
	warn = pwarn
end property

property get derror()
	derror = pderror
end property

property get warntext()
	warntext = pwarntext
end property

property let isduplicate(b)
	pisduplicate = b
end property

property let isdeletecolref(b)
	pisdeletecolref = b
end property

property let dok(b)
	pdok = b
end property

property let warn(b)
	pwarn = b
end property

property let derror(s)
	pderror = s
end property

property let warntext(s)
	pwarntext = s
end property

property let oledbstr(s)
	poledbstr = s
end property

Function GetConnectionString()
	Dim ds,ic,u,p

	If Not poledbstr = "" Then
		GetConnectionString = poledbstr
	Else
		If Not Application("entity_dsn") = "" Then
			GetConnectionString = Application("entity_dsn")
		Else
			ds = GetSession("ds")
			ic = GetSession("db")
			u = Application("u")
			p = Application("p")
			GetConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID="&u&";Password="&p&";Initial Catalog="&ic&";Data Source="&ds&";Locale Identifier=1033;Connect Timeout=15;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
		End If
	End If
End Function

Function OpenClientConnection()

	'On Error Resume Next

	pdok = True
	pderror	= ""

	GetConnection

	If dberrors(AppConn) Then
		OpenClientConnection = False
	Else
		OpenClientConnection = True
	End If

End Function

Function CloseClientConnection()

	pdok = True
	pderror	= ""

	Call CloseObj(AppConn)

	inTransaction = False

End Function

Function GetConnection()

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

		pdok = True
		pderror	= ""

		Set AppConn = OpenConnection(GetConnectionString())

	End If

End Function

Function OpenConnection(constr)

	On Error Resume Next
	Err.Clear

	Dim MaxCount,Count

	MaxCount = 3
	Count = 0

	pdok = True
	pderror	= ""

	Do While (Count < MaxCount)
		Err.Clear
		pderror = ""

		Set OpenConnection = Server.CreateObject("ADODB.Connection")
		OpenConnection.ConnectionTimeout = Application("app_dsn_ConnectionTimeout")
		OpenConnection.CommandTimeout = Application("app_dsn_CommandTimeout")
		OpenConnection.Open constr

		If Not dberrors(OpenConnection) Then
			Exit Do
		End If

		OpenConnection.Close
		Set OpenConnection = Nothing
		Count = Count + 1
	Loop

	If Count = MaxCount Then
		Set OpenConnection = Nothing
	End If

End Function

Function OpenTransaction()

	'On Error Resume Next

	pdok = True
	pderror	= ""

	AppConn.BeginTrans

	'If dberrors(AppConn) Then
	'	OpenTransaction = False
	'	inTransaction = True
	'Else
		OpenTransaction = True
		inTransaction = True
	'End If

End Function

Function CloseTransaction()

	'On Error Resume Next

	If pdok = False	Then
		CloseTransaction = False
		RollbackTransaction
		Exit Function
	End If

	If inTransaction Then
		AppConn.CommitTrans
	End If

	'If dberrors(AppConn) Then
	'	CloseTransaction = False
	'Else
		CloseTransaction = True
		inTransaction = False
	'End If

End Function

Function RollbackTransaction()

	'On Error Resume Next

	If inTransaction Then
		AppConn.RollbackTrans
	End If

	'If dberrors(AppConn) Then
	'	RollbackTransaction = False
	'Else
		RollbackTransaction = True
		inTransaction = False
	'End If

End Function

Function CloseObj(byref the_object )

	On Error Resume Next

	If IsNull ( the_object ) Then
		Err.Clear
		Exit Function
	End If

	If IsObject( the_object ) then
		If not the_object Is Nothing Then
			If Not the_object.state = 0 Then
				the_object.Close
				'Response.Write "HERE"
			End If
			Set the_object = Nothing
		End If
	End If

	Err.Clear

End Function

Function RunSP(ByVal strSP, params, byRef OutArray)
        'On Error Resume Next

        ' Create the ADO objects
        Dim cmd, OutPutParms
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSP
			cmd.CommandType = adCmdStoredProc

	        collectParams cmd, params, OutPutParms

			' Execute the query without returning a recordset
			cmd.Execute , , adExecuteNoRecords

			if OutPutParms then OutArray = collectOutputParms(cmd, params)
			' Disconnect the recordset and clean up

			If dberrors(cmd.ActiveConnection) Then
				RollbackTransaction
				RunSP = False
			Else
				RunSP = True
			End If

		Else
			RunSP = False
		End If

		On Error Resume Next

        'cmd.ActiveConnection.close()
        Set cmd.ActiveConnection = Nothing
        set cmd = Nothing
        'RunSP = 0
        'Set cmd = Nothing
        'CloseClientConnection
End Function

Function RunSPAsync(ByVal strSP, params, byRef OutArray)
        'On Error Resume Next

        ' Create the ADO objects
        Dim cmd, OutPutParms
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSP
			cmd.CommandType = adCmdStoredProc
	        collectParams cmd, params, OutPutParms

			' Execute the query without returning a recordset
			cmd.Execute , , adAsyncExecute

			if OutPutParms then OutArray = collectOutputParms(cmd, params)
			' Disconnect the recordset and clean up

			If dberrors(cmd.ActiveConnection) Then
				RollbackTransaction
				RunSPAsync = False
			Else
				RunSPAsync = True
			End If

		Else
			RunSPAsync = False
		End If

		On Error Resume Next

        'cmd.ActiveConnection.close()
        'Set cmd.ActiveConnection = Nothing
        set cmd = Nothing
        'RunSPAsync = 0
        'Set cmd = Nothing
        'CloseClientConnection
End Function

Function RunSPReturnRS(ByVal strSP, params, byRef OutArray)

        'On Error Resume Next

        ' Create the ADO objects
        Dim rs, cmd, OutPutParms
        Set rs = Server.CreateObject("adodb.Recordset")
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSP
			cmd.CommandType = adCmdStoredProc

			collectParams cmd, params, OutPutParms

			' Execute the query for readonly
			rs.CursorLocation = adUseClient

			rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
			if OutPutParms then OutArray = collectOutputParms(cmd, params)
			' Disconnect the recordset and clean up
			' Disconnect the recordset

			Call dberrors(rs.ActiveConnection)

		End If

		On Error Resume Next

        'cmd.ActiveConnection.close()
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
        'rs.ActiveConnection.close()
        Set rs.ActiveConnection = Nothing
		'CloseClientConnection

        ' Return the resultant recordset
        Set RunSPReturnRS = rs
End Function

Function RunSPReturnMultiRS(ByVal strSP, params, byRef OutArray)

        On Error Resume Next

        ' Create the ADO objects
        Dim rs, cmd, OutPutParms
        Set rs = Server.CreateObject("adodb.Recordset")
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSP
			cmd.CommandType = adCmdStoredProc

			collectParams cmd, params, OutPutParms

			' Execute the query for readonly
			rs.CursorLocation = adUseClient
			rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
			if OutPutParms then OutArray = collectOutputParms(cmd, params)
			' Disconnect the recordset and clean up
			' Disconnect the recordset

			Call dberrors(rs.ActiveConnection)

		End If

		On Error Resume Next

        'Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing

        ' Can not disconnect because there are multiple recordsets!
        'Set rs.ActiveConnection = Nothing

        ' Return the resultant recordset
        Set RunSPReturnMultiRS = rs
End Function

Function RunSQLReturnMultiRS(ByVal strSQL,ByVal params)

        On Error Resume Next

        ' Set up Command and Connection objects
        Dim rs, cmd, OutPutParms
        Set rs = Server.CreateObject("adodb.Recordset")
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSQL
			cmd.CommandType = adCmdText
			cmd.Prepared = true

			collectParams cmd, params, OutPutParms

			rs.CursorLocation = adUseClient

			rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
			if OutPutParms then OutArray = collectOutputParms(cmd, params)
			' Disconnect the recordset and clean up
			' Disconnect the recordset

			Call dberrors(rs.ActiveConnection)

		End If

		On Error Resume Next

        'Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing

        ' Can not disconnect because there are multiple recordsets!
        'Set rs.ActiveConnection = Nothing

        ' Return the resultant recordset
        Set RunSQLReturnMultiRS = rs
End Function

Function RunSPReturnRS_RW(ByVal strSP, params, byRef OutArray)

        'On Error Resume Next

        ' Create the ADO objects
        Dim rs, cmd, OutPutParms
        Set rs = Server.CreateObject("adodb.Recordset")
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSP
			cmd.CommandType = adCmdStoredProc

			collectParams cmd, params, OutPutParms

			' Execute the query for readonly
			rs.CursorLocation = adUseClient
			rs.Open cmd, , adOpenDynamic, adLockBatchOptimistic
			if OutPutParms then OutArray = collectOutputParms(cmd, params)
			' Disconnect the recordset and clean up
			' Disconnect the recordset
			'Set cmd.ActiveConnection = Nothing

			Call dberrors(rs.ActiveConnection)

		End If

		On Error Resume Next

        Set cmd = Nothing
        'Set rs.ActiveConnection = Nothing

        ' Return the resultant recordset
        Set RunSPReturnRS_RW = rs
End Function

'Function Name: RunSQLReturnRS
'Purpose:  Run a pre-prepared SQL statement with zero or more parameters,
'           return an ADO recordset to caller.

Function RunSQLReturnRS(ByVal strSQL,ByVal params)

        On Error Resume Next

        ' Set up Command and Connection objects
        Dim rs, cmd, OutPutParms
        Set rs = Server.CreateObject("adodb.Recordset")
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSQL
			cmd.CommandType = adCmdText
			cmd.Prepared = true

			collectParams cmd, params, OutPutParms

			rs.CursorLocation = adUseClient

			'Note: These are optimal cursor types and lock levels for performance
			'       In general, if you are not going to update a recordset, then do not
			'       return an updatable recordset --its more overhead.  Also, open forward
			'      only and use ADO pagination functions to re-page the recordset.  This
			'       allows all recordsets to remain completely stateless and disconnected (not
			'       persisted on per-session basis between pages.  Its still quite easy to
			'       put page forward/backward functonality into the application, although the
			'       logic to do so is not included here. I will be happy to show you how
			'       to do so when I review the complete Nile code you are working from

			rs.Open cmd, , adOpenForwardOnly, adLockReadOnly

			' Disconnect the recordsets and cleanup

			Call dberrors(rs.ActiveConnection)

		End If

		On Error Resume Next

        'cmd.ActiveConnection.close()
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
        'rs.ActiveConnection.close()
        Set rs.ActiveConnection = Nothing
		'CloseClientConnection
        Set RunSQLReturnRS = rs

End Function

Function RunSQLReturnRS_RW(ByVal strSQL,ByVal params)

        'On Error Resume Next

        ' Set up Command and Connection objects
        Dim rs, cmd, OutPutParms
        Set rs = Server.CreateObject("adodb.Recordset")
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSQL
			cmd.CommandType = adCmdText
			cmd.Prepared = true

			collectParams cmd, params, OutPutParms

			rs.CursorLocation = adUseClient

			'Note: These are optimal cursor types and lock levels for performance
			'       In general, if you are not going to update a recordset, then do not
			'       return an updatable recordset --its more overhead.  Also, open forward
			'      only and use ADO pagination functions to re-page the recordset.  This
			'       allows all recordsets to remain completely stateless and disconnected (not
			'       persisted on per-session basis between pages.  Its still quite easy to
			'       put page forward/backward functonality into the application, although the
			'       logic to do so is not included here. I will be happy to show you how
			'       to do so when I review the complete Nile code you are working from

			rs.Open cmd, , adOpenDynamic, adLockBatchOptimistic

			Call dberrors(rs.ActiveConnection)

		End If

		On Error Resume Next

        ' Disconnect the recordsets and cleanup
        'Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
        'Set rs.ActiveConnection = Nothing
        Set RunSQLReturnRS_RW = rs

End Function

'Function Name: RunSQL
'Purpose:  Run a pre-prepared SQL statement with no returned recordset.  This is what
'           should be used to run updates, deletes, inserts, for example.

Function RunSQL(ByVal strSQL, byRef params)
        On Error Resume Next

        ' Create the ADO objects
        Dim cmd, outPutParms
        Set cmd = Server.CreateObject("adodb.Command")

        GetConnection

        If pdok Then

			' Init the ADO objects  & the stored proc parameters
	        cmd.ActiveConnection = AppConn

       		cmd.CommandTimeout = Application("app_dsn_CommandTimeout")

			cmd.CommandText = strSQL
			cmd.CommandType = adCmdText
			collectParams cmd, params, OutPutParms

			' Execute the query without returning a recordset
			' Specifying adExecuteNoRecords reduces overhead and improves performance
			cmd.Execute , , adExecuteNoRecords

			If dberrors(cmd.ActiveConnection) Then
				RollbackTransaction
				RunSQL = False
			Else
				RunSQL = True
			End If

		Else
			RunSQL = False
		End If

		On Error Resume Next

        ' Cleanup
        'cmd.ActiveConnection.close()
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
		'CloseClientConnection

End Function


'Function Name: PageRecords
'Purpose:  Paginate an ADO recordset, returning a 2-dimensional array
'           consisting of the database rows for the page requested

Function PageRecords(rs, byRef intPage, byVal PageSize, NumPages, NumRecords)

	If intPage < 1 Then intPage = 1

	If PageSize < 0 Then
		PageSize = NumRecords
		NumPages = 1
		PageRecords = rs.GetRows()
		NumRecords = rs.RecordCount
	Else
		rs.PageSize = PageSize
		NumPages = rs.PageCount
		If intPage > NumPages Then intPage = NumPages
		rs.AbsolutePage = intPage
		PageRecords = rs.GetRows(PageSize, adBookmarkCurrent)
		If (intPage < NumPages) Or (rs.RecordCount Mod PageSize = 0) Then
		    NumRecords = PageSize
		Else
		    NumRecords = rs.RecordCount Mod PageSize
		End If
	End If
End Function

'Function Name: PageRecords
'Purpose:  Paginate an ADO recordset, returning a recordset
'           consisting of the database rows for the page requested

Function PageRecordsReturnRS(byRef rs, byRef intPage, byVal PageSize, NumPages, NumRecords)

   If intPage < 1 Then intPage = 1
   If PageSize < 0 Then PageSize = 10
   rs.PageSize = PageSize
   NumPages = rs.PageCount
   If intPage > NumPages Then intPage = NumPages
   rs.AbsolutePage = intPage
   'PageRecords = rs.GetRows(PageSize, adBookmarkCurrent)
   If (intPage < NumPages) Or (rs.RecordCount Mod PageSize = 0) Then
       NumRecords = PageSize
   Else
       NumRecords = rs.RecordCount Mod PageSize
   End If

End Function

'Sub Name: collectParams
'Purpose:  collect parameters from array of parameters, for binding to pre-
'           prepared SQL statements

Private Sub collectParams(ByRef cmd, ByVal argparams, ByRef OutPutParms)
        Dim params, v
        Dim i, l, u
        'if argparams is empty

        If Not IsArray(argparams) Then Exit Sub

        OutPutParms = false
        params = argparams
        For i = LBound(params) To UBound(params)
            l = LBound(params(i))
            u = UBound(params(i))
            ' Check for nulls.
            If u - l >= 3 Then
                If VarType(params(i)(4)) = vbString Then
                    if params(i)(4) = "" then
                        v=null
                    else
                        v=params(i)(4)
                    end if
                Else
                    v = params(i)(4)
                End If
                if params(i)(2) = adParamOutput or params(i)(2) = adParamReturnValue then OutPutParms = true
                cmd.Parameters.Append cmd.CreateParameter(params(i)(0), params(i)(1), params(i)(2), params(i)(3), v)
            Else
                err.raise m_modName, "collectParams(...): incorrect # of parameters"
            End If
        Next

End Sub

Private Function collectOutputParms(ByRef cmd, argparams)
        Dim params, v, OutArray(30)
        Dim i, l, u
        'if argparams is empty

        'If Not IsArray(argparams) Then Exit Sub

        params = argparams
        For i = LBound(params) To UBound(params)
            OutArray(i) = cmd.Parameters(i).Value
        Next
        collectOutputParms = OutArray
End Function

Function dobatchupdate(rs)

	On Error Resume Next

	rs.UpdateBatch

	If dberrors(rs.ActiveConnection) Then
		rs.ActiveConnection.Errors.Clear
		rs.CancelUpdate
		RollbackTransaction
	End If

End Function

Function dberrors(con)

	Dim objError

	dberrors = True

	If Err.Number <> 0 Then

		pdok = False
		pderror = Trim(Err.Description) 'Trim(Err.Source)

		If InStr(UCase(pderror),"DUPLICATE") > 0 Then
			pisduplicate = True
		End If

		If InStr(UCase(pderror),"DELETE STATEMENT CONFLICTED") > 0 and Not GetSession("IsAdmin") = "Y" Then
			pisdeletecolref = True
			pderror = "The selected record(s) can not be deleted because there are references in other modules to the record(s) you are trying to delete. You must first remove all references to the record(s) you are trying to delete. As an alternative, you can set the status of the record(s) to inactive. "
		End If

	ElseIf IsObject(con) and con.Errors.Count > 0 Then

		pdok = False

		For Each objError in con.Errors
			'Response.Write("Error " & objError.SQLState & ": " & _
			'objError.Description & " | " & objError.NativeError)
			pderror	= pderror & objError.Description

			' Do Special ADO Error Handling Here
			' ======================================================

			' If duplicate then set duplicate flag to then later
			' write appropriate error message
			' ======================================================
			'If objError.Number = -2147217873 Then
			'	pisduplicate = True
			'End If

			If InStr(UCase(objError.Description),"DUPLICATE") > 0 Then
				pisduplicate = True
				Exit For
			End If

			' If trying to delete with column reference constraint then set isdeletecolref flag to then later
			' write appropriate error message
			' ======================================================
			If InStr(UCase(objError.Description),"DELETE STATEMENT CONFLICTED") > 0 and Not GetSession("IsAdmin") = "Y" Then
				pisdeletecolref = True
				pderror = "The selected record(s) can not be deleted because there are references in other modules to the record(s) you are trying to delete. You must first remove all references to the record(s) you are trying to delete. As an alternative, you can set the status of the record(s) to inactive. "
				Exit For
			End If

			'Use the below to check what the SQLState # is
			'=================================================
			'Response.Write objError.Number
			'Response.Write "<BR>"
			'Response.Write objError.Description
			'Response.Write "<BR>"
		Next

		'Response.Write "HERE"
		'Response.End

		con.Errors.Clear

	Else

		dberrors = False

	End If

End Function

End class

%>