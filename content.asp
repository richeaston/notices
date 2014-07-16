<!--#include file="Connection.asp"-->
<link href="css/bootstrap.css" rel="stylesheet">
<link href="css/bootstrap-responsive.css" rel="stylesheet">
<%

Dim rsNotice    	'Holds the recordset for the records in the database
Dim strSQL          'Holds the SQL query for the database
Dim strcount           
Set rsNotice = Server.CreateObject("ADODB.Recordset")	
strSQL = "Select * from [Notices] where AND Date() >= [sdate] AND Date() <= [edate] AND [approved] = True;"
strcount = 0
rsNotice.Open strSQL, adoCon
%>
<div class="container">
<BR>
<div id="Carousel" class="carousel slide" data-interval="10000" data-pause='hover' rel='tooltip' title='Paused, to continue move the mouse.'>
  <!-- Carousel items -->
  <div class="well">
	<div class="carousel-inner" data-interval="10000" data-pause='hover'>
  	<%	 Do While not rsNotice.EOF
		 if strcount = 0 then
		   response.write("<div class='item active' data-pause='hover'>")
		 else  
		   response.write("<div class='item' data-pause='hover'>")
		 end if
		   Response.Write("<P><h2>" & rsNotice("Title") & "&nbsp;<i class='icon-comment'></i></h2>")
		   Response.Write("<Span class='label description'>Author</span>&nbsp;" & rsNotice("Author") & "</p>")
		   Response.Write("<p><Span class='label label-info'>Group</span>&nbsp;" & rsNotice("Group") & "</p>")
		   Response.Write("<p><Span class='label label-success'>Start Date</span>&nbsp;" & rsNotice("Sdate") & "&nbsp;<Span class='label label-important'>End Date</span>&nbsp;" & rsNotice("Edate") & "</p>" )
		   Response.Write("<Strong>" & rsNotice("Notice") & "</strong>")
		   Response.Write("</div>")
		   rsNotice.MoveNext
		   strcount=strcount+1
		 loop
rsNotice.close
Set rsNotice = Nothing
Set strSQL = Nothing
Set adocon = Nothing
	%>
</div>
</div>
</div>
</div>

