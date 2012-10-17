<%
If InStr(LCase(Replace(Request.ServerVariables("HTTP_USER_AGENT"),"""","",1)),"blackberry") > 0 Then
	Response.Redirect("default.asp")
End If
%>
<html>
<head>
<title>MC Mobile</title>
<style type="text/css">
	html,form,body{
		background-color:#6475d7;
		margin:0px;
		padding:0px;
	}
</style>
</head>
<body>
<center>
	<div style="background-image:url('bbe.jpg'); background-repeat:no-repeat;width:450px; height:800px;">
		<div style="position:relative;top:172px;left:-1px;">
			<iframe width="378" height="259" src="default.asp" frameborder="0" scrolling="auto" style="border:none; overflow:auto; overflow-x:hidden; overflow-y:auto;"></iframe>
		</div>
	</div>
</center>
</body>
</html>