
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title>USDF Membership Card</title>


</head>

<body>

<img height="285px" width="452px" src="http://www.usdf.org/images/pdf/member_card_online.jpg">
<table background="" height="270px" width="452px" style="position:fixed;float:left;top:20px;left:30px">
<tr><td height="80px"></td></tr>
<%
	Dim SQL
	strSQL = "select * " & _
	"from tblhorseinformation" & _
	" where usdfhorsenumber = " & request.querystring("membernumber")

	
	'	strSQL = "select * " & _
'	"from tblhorseinformation inner join tblpersontohorserelationship " & _
'	"on tblhorseinformation.usdfhorsenumber = tblpersontohorserelationship.USDFHorsenumber " & _
'	"where tblhorseinformation.usdfhorsenumber = " & request.querystring("membernumber") & " and " & _
'	"tblpersontohorserelationship.expirationdate > '" & date() & "'"

	set rs = Server.CreateObject("ADODB.Recordset")	
	Application("wConnectString") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=webuser;Password=webuser;Initial Catalog=Webdb;Data Source=Poison"

	rs.Open strSQL, _
	Application("wConnectString"), _
	3, _
	4
	
	if rs.fields("registrationdescription") = "Life" then
		response.write "<tr><td height=""15px""><font size=""6px"">Lifetime Registration</font></td></tr>"
	end if
%>
<tr><td height="15px"><font size="6px">
<%		
	response.write rs.fields("usdfhorsename")
%>
</font></td></tr><tr><td height="5px"><font size="4px">
<%
	response.write request.querystring("membernumber")
%>
</font></td></tr><tr><td height="14px"><font size="4px">
<%
		if rs.fields("registrationdescription") = "Life" then
			datereg = rs.fields("LifetimeRegistereddate")
		else
			datereg = rs.fields("hidregistereddate")
		end if
		response.write "Date registered with USDF: " & datereg

%>
</font></td></tr><tr><td height="24px"><font size="4px">
<%
'owners
	strSQL = "select * from" & _
	" tblpersontohorserelationship inner join tblpersoninformation" & _
	" on tblpersontohorserelationship.usdfpersonnumber = tblpersoninformation.usdfpersonnumber" & _
	" where usdfhorsenumber = " & request.querystring("membernumber") & " and tblpersontohorserelationship.relationshipdescription = 'Owner' and tblpersontohorserelationship.expirationdate > '" & date() & "' order by tblpersoninformation.usdfpersonnumber;"
	set rs = Server.CreateObject("ADODB.Recordset")	
	Application("wConnectString") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=webuser;Password=webuser;Initial Catalog=Webdb;Data Source=Poison"

	rs.Open strSQL, _
	Application("wConnectString"), _
	3, _
	4
	Do While Not rs.EOF
		ownername = rs.fields("firstname") & " " & rs.fields("lastname") & ", " & rs.fields("relationshipdescription")
		response.write ownername
		if Life then
			response.write "<br />Relationship date: " & rs.fields("relationshipdate") & "<br />"
		else
			response.write "<br />Relationship date: " & rs.fields("relationshipdate") & "<br />"
		end if
		rs.movenext
	Loop
	rs.close

'Business owners
	strSQL = "select * from" & _
	" tblbusinesstohorserelationship inner join tblbusinessinformation" & _
	" on tblbusinesstohorserelationship.usdfbusinessnumber = tblbusinessinformation.usdfbusinessnumber" & _
	" where usdfhorsenumber = " & request.querystring("membernumber") & " and tblbusinesstohorserelationship.relationshipdescription = 'Owner' and tblbusinesstohorserelationship.expirationdate > '" & date() & "' order by tblbusinessinformation.usdfbusinessnumber;"
	set rs = Server.CreateObject("ADODB.Recordset")	
	Application("wConnectString") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=webuser;Password=webuser;Initial Catalog=Webdb;Data Source=Poison"

	rs.Open strSQL, _
	Application("wConnectString"), _
	3, _
	4
	Do While Not rs.EOF
		ownername = rs.fields("businessname") & ", " & rs.fields("relationshipdescription")
		if Life then
			response.write "<br />Relationship date: " & rs.fields("relationshipdate") & "<br />"
		else
			response.write "<br />Relationship date: " & rs.fields("relationshipdate") & "<br />"
		end if
		rs.movenext
	Loop
	rs.close
	
'lessee	
	strSQL = "select *" & _
	" from tblpersontohorserelationship inner join tblpersoninformation" & _
	" on tblpersontohorserelationship.usdfpersonnumber = tblpersoninformation.usdfpersonnumber" & _
	" where usdfhorsenumber = " & request.querystring("membernumber") & " and tblpersontohorserelationship.relationshipdescription = 'Lessee' and tblpersontohorserelationship.expirationdate > '" & date() & "' order by tblpersoninformation.usdfpersonnumber;"
	set rs = Server.CreateObject("ADODB.Recordset")	
	Application("wConnectString") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=webuser;Password=webuser;Initial Catalog=Webdb;Data Source=Poison"

	rs.Open strSQL, _
	Application("wConnectString"), _
	3, _
	4
	
	do While Not rs.EOF 
		ownername = rs.fields("firstname") & " " & rs.fields("lastname") & ", " & rs.fields("relationshipdescription")
		if Life then
			response.write "<br />Relationship date: " & rs.fields("relationshipdate") & "<br />"
		else
			response.write "<br />Relationship date: " & rs.fields("relationshipdate") & "<br />"
		end if
		rs.movenext
	loop
	rs.close
%>
</font></td></tr>
</table>
</body>
</html>


<%	
'lessee	
'	strSQL = "select *" & _
'	" from tblpersontohorserelationship inner join tblpersoninformation" & _
'	" on tblpersontohorserelationship.usdfpersonnumber = tblpersoninformation.usdfpersonnumber" & _
'	" where usdfhorsenumber = " & request.querystring("membernumber") & " and tblpersontohorserelationship.relationshipdescription = 'Lessee' and tblpersontohorserelationship.expirationdate > '" & date() & "' order by tblpersoninformation.usdfpersonnumber;"
'	set rs = Server.CreateObject("ADODB.Recordset")	
'	Application("wConnectString") = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=webuser;Password=webuser;Initial Catalog=Webdb;Data Source=Poison"

'	rs.Open strSQL, _
'	Application("wConnectString"), _
'	3, _
'	4
	
'	do While Not rs.EOF 
'		ownername = rs.fields("firstname") & " " & rs.fields("lastname") & ", " & rs.fields("relationshipdescription")
'		pdf.addtextpos 306-PDF.GetTextWidth(ownername)\2, HP, ownername
'		HP = HP + HPS - 5
'		PDF.SetProperty csPropTextSize, 14
'		if Life then
'			pdf.addtextpos 306-PDF.GetTextWidth("Relationship date: " & rs.fields("relationshipdate"))\2, HP, "Relationship date: " & rs.fields("relationshipdate")
'		else
'			pdf.addtextpos 306-PDF.GetTextWidth("Relationship date: " & rs.fields("relationshipdate"))\2, HP, "Relationship date: " & rs.fields("relationshipdate")
'		end if
'		PDF.SetProperty csPropTextSize, 25
'		HP = HP + HPS + 5
'		rs.movenext
'	loop
'	rs.close
%>