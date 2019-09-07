<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\sql.inc" -->

<%
var usok=false
var sql=""
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var bk=parseInt(Request("bk"))
if (isNaN(bk)) {Response.Redirect("bloknews.asp")}

var sminame=""
var ErrorMsg=""
var ShowForm=true
var name=""
var subj=""
var id=0

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="edblock.asp"
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

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

Records.Source="Select * from block_news where  id="+bk+" and SMI_ID="+smi_id
Records.Open()
if (Records.EOF ) {
	Records.Close()
	{Response.Redirect("pubarea.asp")}
}
name=TextFormData(Records("NAME").Value,"")
subj=TextFormData(Records("SUBJ").Value,"")
Records.Close()

isFirst=String(Request.Form("Submit"))=="undefined"

if (!isFirst) {
	name=TextFormData(Request.Form("name"),"")
	subj=TextFormData(Request.Form("subj"),"")
	
	if (subj.length<1) {ErrorMsg+="Отсетствет наименование блока.<br>"}
	if (name.length<1) {ErrorMsg+="Отсутствует описание блока.<br>"}
	
	if (ErrorMsg=="") {
		sql=edblock
		sql=sql.replace("%ID",bk)
		sql=sql.replace("%NAM",name)
		sql=sql.replace("%SMI",smi_id)
		sql=sql.replace("%SUBJ",subj)
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
		}
		catch(e){
			Connect.RollbackTrans()
			ErrorMsg+=ListAdoErrors()
		}
		if (ErrorMsg==""){
		  Connect.CommitTrans()
		  ShowForm=false
		}
	}
}

%>

<html>
<head>
<title>Изменение блока новостей (<%=sminame%>)</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" bordercolor="#000000"> 
<tr> <td bordercolor="#CCCCCC"> <p><a href="area.asp"><b>Кабинет администратора</b></a> 
| <a href="index.asp">Редактировать сайт</a> | <a href="bloknews.asp">Информационные блоки</a> 
</p></td></tr> </table><%if(ErrorMsg!=""){%> <center> <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> 
<br><%=ErrorMsg%></p></center><%}%> <%if(ShowForm){%> 
<h1 align="center"><b><font size="3">Здесь можно изменить реквизиты информационного 
  блока </font><font size="3" color="#0000CC"><%=sminame%></font></b></h1>
<form name="form1" method="post" action="edblock.asp"> 
<p align="center"> <input type="hidden" name="bk" value="<%=bk%>"> Для добавления 
информационного блока, пожалуйста, заполните поля формы: </p><table width="780" border="1" bordercolor="#FFFFFF" align="center"> 
<tr> <td width="380" bgcolor="#CCCCCC" bordercolor="#333333"> <div align="center"> 
          <p><b>Реквизиты:</b></p>
        </div></td><td width="10"> <p>&nbsp;</p></td><td bgcolor="#CCCCCC" width="582" bordercolor="#333333"> 
<div align="center"> <p><b>Значения:</b></p></div></td></tr> <tr bordercolor="#333333"> 
<td width="380" height="14" valign="middle"> <div align="right"><font size="2" color="#FF0000">Видимое 
посетителям наименование блока:</font><font size="2">&nbsp;&nbsp;</font></div></td><td width="10" height="14" bordercolor="#FFFFFF"> 
<div align="center">-</div></td><td width="582" height="14" valign="top"> <p> 
<input type="text" name="subj" value="<%=isFirst?subj:Request.Form("subj")%>" maxlength="50" size="50"> 
до 50 символов</p></td></tr> <tr bordercolor="#333333"> <td width="380" valign="middle" height="18"> 
<div align="right"><font size="2" color="#FF0000">Краткое описание, назначение:</font><font size="2">&nbsp;&nbsp;</font></div></td><td width="10" height="18" bordercolor="#FFFFFF"> 
<div align="center">-</div></td><td width="582" valign="top" height="18"> <p><font size="2"> 
<input type="text" name="name" value="<%=isFirst?name:Request.Form("name")%>" maxlength="100" size="50"> 
до 100 символов</font></p></td></tr> </table><p align="center"><font color="#FF0000">Параметры 
выделенные красным цветом обязательны к заполнению!</font></p><p align="center"> 
<input type="submit" name="Submit" value="Сохранить"> </p></form><%} 
else 
{%> <center> <h1><font color="#3333FF">Блок изменен!</font></h1></center>
<%}%>
<p>&nbsp;</p>
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
