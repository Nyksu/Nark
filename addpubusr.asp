<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
var usok=false
var sql=""
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var sminame=""
var isFirst=false
var name=""
var nikname=""
var ps1=""
var ps2=""
var email=""
var phone=""
var usrtp=0
var ErrorMsg=""
var id=0
var ShowForm=true

if (String(Session("id_mem"))=="undefined") {
Session("backurl")="addpubusr.asp"
Response.Redirect("login.asp")
}

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

if (!usok) {Response.Redirect("index.asp")}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

isFirst=String(Request.Form("Submit"))=="undefined"

if (!isFirst) {
	name=TextFormData(Request.Form("name"),"")
	nikname=TextFormData(Request.Form("nikname"),"")
	ps1=TextFormData(Request.Form("ps1"),"")
	ps2=TextFormData(Request.Form("ps2"),"")
	email=TextFormData(Request.Form("email"),"")
	phone=TextFormData(Request.Form("phone"),"")
	usrtp=Request.Form("usrtp")
	
	if (email.length>0) {	 
	     if (!/(\w+)@((\w+).)*(\w+)$/.test(email)) {ErrorMsg=ErrorMsg+"Неверный формат поля 'E-mail'.<br>"}}
	if (name.length<3) {ErrorMsg+="Слишком короткое имя.<br>"}
	if (nikname.length==0) {ErrorMsg+="Пустой псевдоним.<br>"}
	if (ps1.length<3) {ErrorMsg+="Пароль должен быть больше 2-х символов! Повторите ввод пароля.<br>"}
	if (ps1 != ps2) {ErrorMsg+="Пароль введен с ошибкой! Повторите ввод пароля.<br>"}
	if (phone.length > 20) {ErrorMsg+="Для телефона выделено только 20 символов.<br>"}
	if (ErrorMsg=="") {
		Records.Source="Select * from SMI_USR where NIK_NAME='"+nikname+"'"
		Records.Open()
		if (!Records.EOF) {ErrorMsg+="Такой псевдоним уже существует! Выберите другой.<br>"}
		Records.Close()
	}
	if (ErrorMsg=="") {
		id=NextID("UNIVERSAL")
		sql=inspubusr
		sql=sql.replace("%ID",id)
		sql=sql.replace("%NAM",name)
		sql=sql.replace("%NIK",nikname)
		sql=sql.replace("%PS",ps1)
		sql=sql.replace("%SMI",smi_id)
		sql=sql.replace("%TP",usrtp)
		sql=sql.replace("%EML",email)
		sql=sql.replace("%PHN",phone)
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
<title>Добавление пользователя СМИ</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr>
    <td bgcolor="#CCCCCC" bordercolor="#333333">
      <p><a href="index.asp">Вернуться на главную страницу</a> | <a href="area.asp">Кабинет администратора</a></p>
    </td>
  </tr>
</table>
<h1 align="center">Добавляем пользователя СМИ.</h1>
<p align="center"><font color="#0000FF"><b><%=sminame%></b></font></p>

<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 
<%if(ShowForm){%> 
<form action="addpubusr.asp" method="post" name="form1">
  <div align="center">
    <table width="780" border="1" bordercolor="#FFFFFF">
      <tr> 
        <td width="250" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <p align="center"><b>Параметры</b></p>
        </td>
        <td width="6"> 
          <p align="center"><b><font color="#FF0000"></font></b></p>
        </td>
        <td bgcolor="#CCCCCC" bordercolor="#333333"> 
          <p align="center"><b>Значения</b></p>
        </td>
      </tr>
      <tr> 
        <td width="250" bordercolor="#333333"> 
          <p align="right"><font color="#FF0000">ФИО (или наименование организации) 
            :</font> </p>
        </td>
        <td width="6"> 
          <p align="center">-</p>
        </td>
        <td bordercolor="#333333"> 
          <p>
            <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" size="40" maxlength="90">
          </p>
        </td>
      </tr>
      <tr> 
        <td width="250" bordercolor="#333333"> 
          <p align="right"><font color="#FF0000">Псевдоним для входа :</font> 
          </p>
        </td>
        <td width="6"> 
          <p align="center">-</p>
        </td>
        <td bordercolor="#333333"> 
          <p>
            <input type="text" name="nikname" size="40" maxlength="20" value="<%=isFirst?"":Request.Form("nikname")%>">
          </p>
        </td>
      </tr>
      <tr> 
        <td width="250" bordercolor="#333333"> 
          <p align="right"><font color="#FF0000">Пароль :</font> </p>
        </td>
        <td width="6"> 
          <p align="center">-</p>
        </td>
        <td bordercolor="#333333"> 
          <p>
            <input type="password" name="ps1" maxlength="20">
          </p>
        </td>
      </tr>
      <tr> 
        <td width="250" bordercolor="#333333"> 
          <p align="right"><font color="#FF0000">Повторение пароля :</font> </p>
        </td>
        <td width="6"> 
          <p align="center">-</p>
        </td>
        <td bordercolor="#333333"> 
          <p>
            <input type="password" name="ps2" maxlength="20">
          </p>
        </td>
      </tr>
      <tr> 
        <td width="250" bordercolor="#333333"> 
          <p align="right">E-mail : </p>
        </td>
        <td width="6"> 
          <p align="center">-</p>
        </td>
        <td bordercolor="#333333"> 
          <p>
            <input type="text" name="email" size="40" maxlength="40" value="<%=isFirst?"":Request.Form("email")%>">
          </p>
        </td>
      </tr>
      <tr> 
        <td width="250" bordercolor="#333333"> 
          <p align="right">Телефон : </p>
        </td>
        <td width="6"> 
          <p align="center">-</p>
        </td>
        <td bordercolor="#333333"> 
          <p>
            <input type="text" name="phone" size="40" maxlength="20" value="<%=isFirst?"":Request.Form("phone")%>">
          </p>
        </td>
      </tr>
      <tr> 
        <td width="250" bordercolor="#333333"> 
          <p align="right"><font color="#FF0000">Статус пользователя :</font> 
          </p>
        </td>
        <td width="6"> 
          <p align="center">-</p>
        </td>
        <td bordercolor="#333333"> 
          <p>
            <select name="usrtp">
			<%
			Records.Source="Select ID, NAME from USR_TYPE order by NAME"
			Records.Open()
			while(!Records.EOF){%>
				<option value="<%=Records("ID").Value%>" 
				<%=isFirst&&(Records("ID").Value==5)?"selected":""%>
				<%=!isFirst&&(Records("ID").Value==Request.Form("usrtp"))?"selected":""%>> 
				<%=Records("NAME").Value%> </option>
				<%	Records.MoveNext()
			}
			Records.Close()
			%>
			</select>
          </p>
        </td>
      </tr>
      <tr> 
        <td width="250"> 
          <p>&nbsp;</p>
        </td>
        <td width="6"> 
          <p align="center">&nbsp;</p>
        </td>
        <td> 
          <p>&nbsp;</p>
        </td>
      </tr>
      <tr> 
        <td width="250" height="16"> 
          <p>&nbsp;</p>
        </td>
        <td width="6" height="16"> 
          <p align="center">&nbsp;</p>
        </td>
        <td height="16"> 
          <p><font color="#FF0000">Позиции выделенные красным цветом - обязательны 
            к заполнению</font></p>
          </td>
      </tr>
    </table>
    <p>
      <input type="submit" name="Submit" value="Сохранить">
    </p>
  </div>
 </form>
<hr width="780">
<%}
else{%>
<p align="center"><b><font color="#FF0000" size="3">Пользователь добавлен!!</font></b></p>
<%}%> 
<p>&nbsp;</p>
<p align="center"><a href="index.asp">Вернуться к публикациям</a></p>
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
