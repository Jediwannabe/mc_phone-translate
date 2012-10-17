<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>
<!--#INCLUDE FILE="includes/mc_all.asp" -->
<%
	Dim db, WOPK
	Set db = New ADOHelper
	Dim Action

	Action = Request("action")

	Select Case LCase(Action)
		Case "wotaskcomplete"
			WOPK = Request("wopk")
			If WOPK <> "" Then
				Call db.RunSQL("UPDATE WOTASK WITH ( ROWLOCK ) SET Complete = 1 WHERE PK = " & Request("taskid") & " AND WOPK = " & WOPK,"")
			End If
	End Select


%>
