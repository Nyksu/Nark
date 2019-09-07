<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var id=parseInt(Request("pid"))
var setst=parseInt(Request("st"))
var usok=false
var sql=""

if (isNaN(setst)) {Response.Redirect("admarea.asp")}
if (isNaN(id)) {id=0}
if (id==0) {Response.Redirect("area.asp")}

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="area.asp"
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<3) {usok=true}
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
		}
		Records.Close()
	} else {
		usok=true
	}
}

if (!usok) {Response.Redirect("area.asp")}

var isok=true

sql="Select t1.* from publication t1, heading t2 where t1.id="+id+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("area.asp")
}
Records.Close()

sql="Update publication set state="+setst+" where id="+id
Connect.BeginTrans()
try{
	Connect.Execute(sql)
}
	catch(e){
	Connect.RollbackTrans()
	isok=false
}
if (isok){
	Connect.CommitTrans()
}
Response.Redirect("area.asp")
%>

