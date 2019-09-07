<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var ErrorMsg=""

if (String(Session("backurl"))=="undefined"){Session("backurl")="/"}

isFirst=String(Request.Form("login"))=="undefined"
if(!isFirst){	Pass=TextFormData(Request.Form("pass"),"")
	Nik=TextFormData(Request.Form("nik"),"")

	Nik=Nik.replace("/*","")
	Nik=Nik.replace("'","")
	Pass=Pass.replace("/*","")
	Pass=Pass.replace("'","")
	Records.Source="Select * from USERS where PSW='"+Pass+"' and LOGNAME='"+Nik+"'"
	Records.Open()
	if (Records.EOF){ErrorMsg+="Неверный 'Пароль или псевдоним'.<br>"}
	else { if (Records("STATE").Value==1){
		Session("id_mem")=String(Records("ID").Value)
		Session("name_mem")=String(Records("FIO").Value)
		Session("nik_mem")=String(Records("LOGNAME").Value)
		Session("state_mem")=String(Records("STATE").Value)
		Session("is_adm_mem")=String(Records("ADM").Value)
		Session("email_mem")=String(Records("EMAIL").Value)
		} else {
		Session("id_mem")="undefined"
		Session("name_mem")="undefined"
		Session("is_adm_mem")="undefined"
		Session("is_host")="undefined"
		ErrorMsg+="Неверный 'Пароль или псевдоним'.<br>"
		}
	}
	Records.Close()
	if (ErrorMsg==""){
		Records.Source="Select * from ENTERPRISE where USERS_ID="+Session("id_mem")+" and ID="+company
		Records.Open()
		if   (!Records.EOF) {Session("is_host")=1}
		Records.Close()
		Response.Redirect(Session("backurl"))
	}
}

%>

<html>
<head>
<title>Аутентификация</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF"> 
<tr> <td bgcolor="#CCCCCC" bordercolor="#333333"> <p><a href="index.asp">На главную страницу</a> 
| </p></td></tr> </table><p align="center"><%if(ErrorMsg!=""){%> </p><center> 
<h2> <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br> <%=ErrorMsg%></p></h2></center><%}%> 
<p>&nbsp;</p><form name="form1" method="post" action="login.asp"> <table width="100%" border="0" cellspacing="4" cellpadding="0"> 
<tr valign="middle"> <td width="50%" align="right"> <p><b>Псевдоним (логин)</b></p></td><td width="50%"> 
<input type="text" name="nik" value="<%=isFirst?"":Request.Form("nik")%>" size="20" maxlength="20"> 
* </td></tr> <tr valign="middle"> <td width="50%" align="right"> <p><b>Пароль</b></p></td><td width="50%"> 
<input type="password" name="pass" size="20" maxlength="20"> * </td></tr> </table><p align="center"> 
<input type="submit" name="login" value="Ввод"> </p><hr size="1" width="300"> 
<p align="center"><b>* - Обязательные поля</b></p><p align="center">&nbsp;</p></form>
  

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
  начало</font></a><b> | </b><a href="message.asp">Обратная связь</a><b> | </b></font></p>
<p align="center"><font size="1">&copy; 2003 программирование <a href="http://www.rusintel.ru">Русинтел</a></font></p>
<p align="center">&nbsp;</p>
  </body>
</html>
