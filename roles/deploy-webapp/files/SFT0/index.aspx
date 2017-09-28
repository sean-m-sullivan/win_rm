<html>
<head>
<title>Hello World!</title>
</head>
<body>
	<h1>Hello World!</h1>
	<p>
		It is now
		<% Response.Write(DateTime.Now.ToString("ddd, MMM dd yyyy hh:mm:ss tt")) %></p>
	<p>
		You are coming from 
		<% Response.Write(Request.ServerVariables("remote_addr")) %>
	</p>
	<p>
		You are accessing
		<% Response.Write(System.Environment.MachineName) %>
	</p>

</html>