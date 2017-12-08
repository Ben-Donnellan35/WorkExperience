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
		
		dim arrClientName
		dim arrServiceName
		
		dim firstName
		dim lastName
		dim dateOfBirth
		dim address
		dim postcode
		dim telephone
		dim email
		
		set rst = Server.CreateObject("ADODB.Recordset")
		
		set conn = Server.CreateObject("ADODB.Connection")
    	conn.Open "Provider=sqloledb;Data Source=localhost\SQL2017;Initial Catalog=Traveller;User ID=sa;Password=Password123"
	    set cmd = Server.CreateObject("ADODB.Command")
	    set cmd.ActiveConnection = conn
		
	    cmd.CommandType = 4
			
		'* Save details of the auto collect transaction
        cmd.CommandText = "spPopulateBookingPage"
                 
		rst.CursorLocation = 1
    	rst.Open cmd, , 0
		
		 if not rst.EOF then
		 	arrClientName = rst.getRows		
         end if
		 
		set rst = rst.NextRecordset

        if not rst.EOF then
            arrServiceName = rst.getRows
        end if	          
		
		cmd.Execute
		
		conn.close
%>

<html>
<head>
	<title>Booking.asp</title>
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
	<h2>Booking</h2>
	<table align="center">
		<tr>
			<th>Client Name</th>
			<th>Service Name</th>
			<th>Nights</th>
			<th>Special Requests</th>
			<th>Date From</th>
			<th>Date To</th>
		</tr>
	</table><br><br>
	<form>
		Client:
		<select>
				<%
				For i = 0 To UBound(arrClientName, 2)
  				%>
					<option><%Response.Write(arrClientName(0,i))%></option>
				<%next%>
		</select>
		Service:
		<select>
				<%
				For i = 0 To UBound(arrServiceName, 2)
  				%>
					<option><%Response.Write(arrServiceName(0,i))%></option>
				<%next%>
		</select>
		Nights:
		<input type="text">
		Special Requests:
		<input type="text">
	</form>
</body>
</html>