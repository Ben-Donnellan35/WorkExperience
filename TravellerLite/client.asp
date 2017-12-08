<!doctype html>
<%
 		dim conn
        dim cmd
		dim cmd2
		dim i
		
		dim adInteger : adInteger = 3
		dim adDate	: adDate = 7
		dim adWChar : adWChar = 202
		dim adParamInput : adParamInput = 1
		
		dim rst
		
		dim arr
		
		dim firstName
		dim lastName
		dim dateOfBirth
		dim address
		dim postcode
		dim telephone
		dim email
		
		set rst = Server.CreateObject("ADODB.Recordset")
		
		if (Request("firstName") <> "") then
		
			set conn = Server.CreateObject("ADODB.Connection")
	    	conn.Open "Provider=sqloledb;Data Source=localhost\SQL2017;Initial Catalog=Traveller;User ID=sa;Password=Password123"
		    set cmd = Server.CreateObject("ADODB.Command")
		    set cmd.ActiveConnection = conn
			
		    cmd.CommandType = 4
				
			'* Save details of the auto collect transaction
	        cmd.CommandText = "spClientAdd"
	                    
			cmd.Parameters.Append cmd.CreateParameter("@FirstName", adWChar, adParamInput, 30, Request("firstname"))
	        cmd.Parameters.Append cmd.CreateParameter("@LastName", adWChar, adParamInput, 30, Request("lastname"))
	        cmd.Parameters.Append cmd.CreateParameter("@DOB", adDate, adParamInput, , CDate(Request("dateofbirth")))
			cmd.Parameters.Append cmd.CreateParameter("@Address", adWChar, adParamInput, 30, Request("address"))
	        cmd.Parameters.Append cmd.CreateParameter("@Postcode", adWChar, adParamInput, 30, Request("postcode"))
	        cmd.Parameters.Append cmd.CreateParameter("@Telephone", adWChar, adParamInput, 30, Request("telephone"))
	      	cmd.Parameters.Append cmd.CreateParameter("@Email", adWChar, adParamInput, 30, Request("email"))
			
			cmd.Execute
			
			conn.close
		
		end if
		
		set conn = Server.CreateObject("ADODB.Connection")
	    conn.Open "Provider=sqloledb;Data Source=localhost\SQL2017;Initial Catalog=Traveller;User ID=sa;Password=Password123"
		set cmd2 = Server.CreateObject("ADODB.Command")
		set cmd2.ActiveConnection = conn
		
		cmd2.CommandType = 4
		
		cmd2.CommandText = "spClientGet"
		
		rst.CursorLocation = 1
    	rst.Open cmd2, , 0
		
		 if not rst.EOF then
                    arr = rst.getRows
         end if	        
%>

<html>
<head>
	<title>Client.asp</title>
	<link href="Index.css" type="text/css" rel="stylesheet">
</head>
<body>
	<h1>Traveller</h1>
	<div class=buttons>
		<ul>
			<li><a href="index.asp" target="_self"> Home</a></li>
			<li><a href="client.asp" target="_self">Client</a></li>
			<li><a href="supplier.asp" target="_self">Supplier</a></li>
			<li><a href="booking.asp" target="_self">Booking</a></li>
			<li><a href="other.asp" target="_self">Other</a></li>
		</ul>
	</div>
	<h2>Client</h2>
	<table align="center">
		<tr>
			<th>First name</th>
			<th>Last name</th>
			<th>Date of birth</th>
			<th>Address</th>
			<th>Postcode</th>
			<th>Telephone</th>
			<th>Email</th>
		</tr>

		<%
		For i = 0 To UBound(arr,2)
  		%>
	
		<tr>
			<td><%Response.Write(arr(0, i))%></td>
			<td><%Response.Write(arr(1, i))%></td>
			<td><%Response.Write(arr(2, i))%></td>
			<td><%Response.Write(arr(3, i))%></td>
			<td><%Response.Write(arr(4, i))%></td>
			<td><%Response.Write(arr(5, i))%></td>
			<td><%Response.Write(arr(6, i))%></td>
		</tr>
		
	<%next%>
		
	</table><br><br>
	<form method="post">
		First name:
		<input type="text" name="firstname">
		Last name:
		<input type="text" name="lastname">
		Date of birth:
		<input type="month" name="dateofbirth">
		Address:
		<input type="text" name="address"><br><br>
		Postcode:
		<input type="text" name="postcode">
		Telephone:
		<input type="text" name="telephone">
		Email:
		<input type="text" name="email">
		<input type="submit" value="submit">
	</form>
</body>
</html>