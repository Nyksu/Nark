<%
function CreateRecordSet(){
	var recs=Server.CreateObject("ADODB.Recordset")
	recs.ActiveConnection=Connect
	recs.CursorType = 0
	recs.CursorLocation = 2
	recs.LockType = 3
	return recs
}
%>