<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var usok=false
var sql=""
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var bk=parseInt(Request("bk"))
if (isNaN(bk)) {Response.Redirect("bloknews.asp")}
var pos=parseInt(Request("pos"))
if (isNaN(pos)) {pos=1}

var sminame=""
var ErrorMsg=""
var id=0

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="addposbloq.asp"
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")==1) {usok=true}
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

if (!usok) {Response.Redirect("index.asp")}

Records.Source="Select * from block_news where smi_id="+smi_id+" and id="+bk
Records.Open()
if (Records.EOF ) {
	Records.Close()
	Response.Redirect("bloknews.asp")
}
Records.Close()

Records.Source="Select * from news_pos where posit="+pos+" and block_news_id="+bk
Records.Open()
if (!Records.EOF ) {
	Records.Close()
	Response.Redirect("block.asp?bk="+bk)
}
Records.Close()

sql="Insert into news_pos(block_news_id,posit,coment,publication_id) values("+bk+","+pos+",'',null)"
Connect.BeginTrans()
try{
	Connect.Execute(sql)
}
catch(e){
	Connect.RollbackTrans()
	Response.Redirect("block.asp?bk="+bk)
}
Connect.CommitTrans()
Response.Redirect("block.asp?bk="+bk)
%>