<%
Sub aspdebug()
%>
	<b>Session	Variables: <%=Now()%></b><br/>
	<% =Replace(Replace(GetSession(""),"^^","<br/>"),"^","") %>
	</p>
<%
End Sub

Function GetVarType( sess_type )

	dim a_obj

	select case VarType( sess_type )
		case 0
			a_obj = "Empty"
		case 1
			a_obj = "Null"
		case 2
			a_obj = "Integer"
		case 3
			a_obj = "Long Integer"
		case 4
			a_obj = "Single-precision floating-point number"
		case 5
			a_obj = "Double-precision floating-point number"
		case 6
			a_obj = "Currency"
		case 7
			a_obj = "Date"
		case 8
			a_obj = "String"
		case 9
			a_obj = "Automation object"
		case 10
			a_obj = "Error"
		case 11
			a_obj = "Boolean"
		case 12
			a_obj = "Variant"
		case 13
			a_obj = "Data-access object"
		case 17
			a_obj = "Byte"
		case 8192
			a_obj = "Array"
		case else
			a_obj = "Unknown"
	end select

	GetVarType = a_obj

End Function

Function GetContents( the_contents )

	dim sess_contents

	if IsNull( the_contents ) then
		sess_contents = ""
	else
		sess_contents = Server.HTMLEncode( the_contents )
	end if

	if Err.Number <> 0 OR IsNull(sess_contents) OR IsEmpty(sess_contents) then
		GetContents = ""
	else
		GetContents = sess_contents
	end if

End Function

Sub showformvars()
Dim x
%>

		<table border="1">
					<% For Each x In Request.Form %>
					<tr>
					<td><font size="1" face="Arial"><% = x %></font></td>
					<td><font size="1" face="Arial">
					<%= Request.Form(x) %>
					</td>
					</font>
			</tr>
		<% Next %>
		</table>

<%
End Sub
%>


