<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var pid=parseInt(Request("pid"))
if (isNaN(pid)) {pid=0}

var usok=false
var id_usr=0
var sql=""
var bid=0
var name=""
var sname=""
var sminame=""

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="bloknews.asp"
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<5) {usok=true}
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

if (!usok) {Response.Redirect("index.asp")}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

%>

<html>
<head>
<title>Информационные блоки новостей и публикаций</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF"> 
<tr> <td bgcolor="#CCCCCC" bordercolor="#333333"> <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На 
главную страницу</a> | <a href="area.asp">Кабинет администратора</a> | </font></p></td></tr> 
</table><h1 align="center">Сайт: <%=sminame%></h1>
<h1 align="center">Информационные блоки: 
  <%
var recs=CreateRecordSet()
Records.Source="Select * from block_news where smi_id="+smi_id+" order by id"
Records.Open()
while (!Records.EOF )
{
	bid=String(Records("ID").Value)
	name=String(Records("NAME").Value)
	sname=TextFormData(Records("SUBJ").Value,"")
%>
</h1>
<table width="780" border="1" bordercolor="#FFFFFF" align="center"> <tr bgcolor="#33CCFF" bordercolor="#333333"> 
<td width="150" bgcolor="#CCCCCC"> <p><b>Информационный Блок</b></p></td><td bgcolor="#FFFFFF"> 
      <p>|| <a href="block.asp?bk=<%=bid%>">Публикации в блоке</a> || <a href="edblock.asp?bk=<%=bid%>">Изменить 
        реквизиты блока</a> || </p>
    </td></tr> <tr> <td colspan="2" height="8"> <p><b>Видимое 
название бока:</b> <font color="#0000FF"><%=sname%><br> </font><B>Краткое описание:</B> 
<%=name%> </p></td></tr> </table>
<%if (pid>0) {%>
<p align="center">Разместить публикацию в позицию: || 
  <%
recs.Source="Select * from news_pos where block_news_id="+bid+" order by posit"
recs.Open()
while (!recs.EOF )
{
	pos=String(recs("POSIT").Value)
	isfull="*"
	if (String(recs("PUBLICATION_ID").Value)=="null") {isfull=""}
%>
  <a href="setposblok.asp?bk=<%=bid%>&pos=<%=pos%>&pid=<%=pid%>"><%=pos%><%=isfull%></a>|| 
  <%
	recs.MoveNext()
} 
recs.Close()%>
  <%}%>
<hr width="780">
<%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
<p>&nbsp; </p>
<p>&nbsp;</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor="#000000"> 
    <td width="300"> 
      <p>&nbsp;</p>
    </td>
    <td align="RIGHT"> 
      <p><b></b></p>
    </td>
  </tr>
</table>
<p align="center"><font size="1"><b>| </b><a href="index.asp"><font size="1">В 
  начало</font></a><b> | </b><a href="message.asp">Обратная связь</a><b> | </b><a href="org.asp">Реквизиты</a><b> |</b></font></p>
<p align="center"><font size="1">&copy; 2003 программирование <a href="http://www.rusintel.ru">Русинтел</a></font></p>
  <p align="center">&nbsp;</p>
  </body>
</html>
