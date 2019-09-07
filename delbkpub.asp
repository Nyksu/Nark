<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var bk=parseInt(Request("bk"))
var pid=parseInt(Request("pid"))
var ps=parseInt(Request("ps"))
var tp=parseInt(Request("tp"))
if (isNaN(bk)) {bk=0}
if (isNaN(tp)) {tp=0}

if (isNaN(pid)) {Response.Redirect("bloknews.asp")}
if (bk==0) {Response.Redirect("bloknews.asp")}

var backurl="bloknews.asp?pid="
if (tp==0) {backurl+=pid}
else {backurl="block.asp?bk="+bk}

var usok=false
var id_usr=0
var sql=""

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="delbkpub.asp?bk="+bk+"&pid="+pid+"&tp="+tp
		if (! isNaN(ps)) {Session("backurl")+="&ps="+ps}
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<2) {usok=true}
	id_usr=TextFormData(Session("id_mem_pub"),"0")
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

sql="Select t1.* from block_news t1, news_pos t2 where t1.id="+bk+" and t1.smi_id="+smi_id+" and t2.block_news_id=t1.id and t2.publication_id="+pid
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	usok=true
} else {
	usok=false
}
Records.Close()


if (!usok) {Response.Redirect("index.asp")}

sql="Update NEWS_POS set publication_id=null where block_news_id="+bk+" and publication_id="+pid
if (! isNaN(ps)) {sql+=" and posit="+ps}

Connect.BeginTrans()
try{
	Connect.Execute(sql)
}
catch(e){
	Connect.RollbackTrans()
	Response.Redirect(backurl)
}
Connect.CommitTrans()
Response.Redirect(backurl)

%>
