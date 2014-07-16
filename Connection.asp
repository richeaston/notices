<%
Dim adoCon          'Holds the Database Connection Object
set adoCon=Server.CreateObject("ADODB.Connection")
adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("notices.mdb")
%>