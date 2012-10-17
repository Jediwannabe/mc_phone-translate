<html>
<head><title>Maintenance Connection</title>

<style type='text/css'>
html, form, body{
	margin:0px;
	padding:0px;
}

.Font1{
    font-family:tahoma,verdana,arial,sans-serif;
    font-size:9pt;
    color:#525252;
}

.Font2{
    font-family:tahoma,verdana,arial,sans-serif;
    font-size:9pt;
    color: #ff6500;
}
</style>
</head>
<body>
<%
    folder = Server.MapPath("images/icons/48/")

    set fso = CreateObject("Scripting.fileSystemObject")
    set fold = fso.getFolder(folder)
    dim iCount
    iCount = 0
	Response.Write("<table class='Font1' cellpadding='2' cellspacing='2' style='width:100%;'>")
    for each file in fold.files
    	if iCount = 0 Then
	    	Response.Write("<tr>")
    	End If
    	If iCount Mod 6 = 0 Then
			Response.Write("</tr><tr>")
    	End If
		Response.Write("<td align='center'>")
    	Response.Write("<img src='images/icons/48/" & file.name & "' /><br />")
    	Response.Write(REPLACE(file.name, ".png", "") & "<br /><br />")
		Response.Write("</td>")
    	iCount = iCount + 1
    next
    	Response.Write("</tr></table>")
    set fold = nothing: set fso = nothing
%>
</body>
</html>