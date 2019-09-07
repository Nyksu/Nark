<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var ErrorMsg=""
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!
var Pass=""
var Nik=""


if (String(Session("backurl"))=="undefined"){Session("backurl")="index.asp"}

isFirst=String(Request.Form("login"))=="undefined"
if(!isFirst){	Pass=TextFormData(Request.Form("pass"),"")
	Nik=TextFormData(Request.Form("nik"),"")

	Nik=Nik.replace("/*","")
	Nik=Nik.replace("'","")
	Pass=Pass.replace("/*","")
	Pass=Pass.replace("'","")
	Records.Source="Select * from SMI_USR where PSW='"+Pass+"' and NIK_NAME='"+Nik+"' and SMI_ID="+smi_id
	Records.Open()
	if (Records.EOF){ErrorMsg+="Неверный 'Пароль или псевдоним'.<br>"}
	else { if (Records("STATE").Value==0){
		Session("id_mem_pub")=String(Records("ID").Value)
		Session("name_mem_pub")=String(Records("NAME").Value)
		Session("nik_mem_pub")=String(Records("NIK_NAME").Value)
		Session("tip_mem_pub")=String(Records("USR_TYPE_ID").Value)
		Session("email_mem_pub")=TextFormData(Records("EMAIL").Value,"")
		} else {
		Session("id_mem_pub")="undefined"
		Session("name_mem_pub")="undefined"
		Session("nik_mem_pub")="undefined"
		Session("tip_mem_pub")="undefined"
		ErrorMsg+="Неверный 'Пароль или псевдоним'.<br>"
		}
	}
	Records.Close()
	if (ErrorMsg==""){
		Response.Redirect(Session("backurl"))
	}
}

%>

<html>
<head>
<title>Аутентификация</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<LINK REL="stylesheet" HREF="/style.css" TYPE="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0">
<TABLE WIDTH="100%" BORDER="1" CELLSPACING="0" CELLPADDING="0" BORDERCOLOR="#FFFFFF"> 
<TR> <TD BGCOLOR="#CCCCCC" BORDERCOLOR="#333333"> <P><a href="index.asp">На главную страницу</A> 
| </P></TD></TR> </TABLE><p align="center"><%if(ErrorMsg!=""){%> </p><center> 
<h2> <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br> <%=ErrorMsg%></p></h2></center><%}%> 
<p align="center">Вход пользователя.</p><form name="form1" method="post" action="logpub.asp"> 
<table width="100%" border="0" cellspacing="4" cellpadding="0"> <tr valign="middle"> 
<td width="50%" align="right"> <p><b>Псевдоним (логин)</b></p></td><td width="50%"> 
<input type="text" name="nik" value="<%=isFirst?"":Request.Form("nik")%>" size="20" maxlength="20"> 
* </td></tr> <tr valign="middle"> <td width="50%" align="right"> <p><b>Пароль</b></p></td><td width="50%"> 
<input type="password" name="pass" size="20" maxlength="20"> * </td></tr> </table><p align="center"> 
<input type="submit" name="login" value="Ввод"> </p><hr size="1" width="400"> 
<p align="center"><b>* - Обязательные поля</b></p><p align="center">&nbsp;</p></form>
<p align="center"><a href="index.asp">Вернуться на главную страницу</a></p>
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
<p align="center">&nbsp;</p>
</body>
</html>
