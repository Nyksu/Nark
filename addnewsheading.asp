<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\sql.inc" -->

<%
var hid=parseInt(Request("hid"))
if (isNaN(hid)) {hid=0}

var usok=false
var sql=""
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var sminame=""
var tit=""
var hiname=""
var ErrorMsg=""
var ShowForm=true
var name=""
var url=""
var period=0
var pglen=0
var redactor=0
var id=0
var isnews=1

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="addnewsheading.asp?hid="+hid
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

tit=sminame

if (hid>0) {
	Records.Source="Select * from heading where id="+hid
	Records.Open()
	if (Records.EOF) {
		Records.Close()
		Response.Redirect("index.asp")
	}
	hiname=String(Records("NAME").Value)
	Records.Close()
	tit+=" | "+hiname
}

isFirst=String(Request.Form("Submit"))=="undefined"

if (!isFirst) {
	name=TextFormData(Request.Form("name"),"")
	url=TextFormData(Request.Form("url"),"")
	period=parseInt(Request.Form("period"))
	pglen=parseInt(Request.Form("pglen"))
	redactor=parseInt(Request.Form("redactor"))
	if (parseInt(Request.Form("isnews"))==1) {isnews=1} else {isnews=0}
	
	if (name.length<3) {ErrorMsg+="Слишком короткое наименование рубрики.<br>"}
	if (isNaN(period)) {ErrorMsg+="Не корректный период  хранения.<br>"}
	if (isNaN(pglen)) {ErrorMsg+="Не корректная длина страницы.<br>"}
	if (pglen<1) {ErrorMsg+="Не корректная длина страницы.<br>"}
	if (isNaN(redactor)) {ErrorMsg+="Не выбран редактор рубрики.<br>"}
	
	if (ErrorMsg=="") {
		id=NextID("PUBHEADINGID")
		sql=insheading
		sql=sql.replace("%ID",id)
		sql=sql.replace("%NAM",name)
		sql=sql.replace("%SMI",smi_id)
		sql=sql.replace("%HID",hid)
		sql=sql.replace("%URL",url)
		sql=sql.replace("%PER",period)
		sql=sql.replace("%PL",pglen)
		sql=sql.replace("%USR",redactor)
		sql=sql.replace("%ISNEWS",isnews)
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
<title>Добавление рубрики: (<%=tit%>)</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На главную страницу</a> 
        | <a href="area.asp">Кабинет администратора</a> | <a href="bloknews.asp">Информационные блоки</a> </font></p>
    </td>
  </tr>
</table>
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 

<%if(ShowForm){%>
<h1 align="center"><b>Вы добавляете рубрику в : <%=tit%></b></h1>
 
<form name="form1" method="post" action="addnewsheading.asp">
  <p align="center">Для добавления рубрики, пожалуйста, заполните поля формы: 
    <input type="hidden" name="hid" value="<%=hid%>">
    <font size="2">
    <input type="hidden" name="url" value="<%=isFirst?"":Request.Form("url")%>" maxlength="50">
    </font> </p>
  <table width="780" border="1" bordercolor="#FFFFFF" align="center">
    <tr> 
      <td width="300" bgcolor="#CCCCCC" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>Параметры:</b></p>
        </div>
      </td>
      <td width="6"> 
        <p>&nbsp;</p>
      </td>
      <td bgcolor="#CCCCCC" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>Значения:</b></p>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="300" bordercolor="#333333" height="14" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Наименование рубрики:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" height="14" valign="top"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" maxlength="100" size="45">
        </p>
      </td>
    </tr>
    <tr> 
      <td width="300" bordercolor="#333333" valign="middle" height="5"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Период хранения:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="5"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" valign="top" height="5"> 
        <p><font size="2"> 
          <input type="text" name="period" value="<%=isFirst?"15":Request.Form("period")%>" maxlength="20">
          дней 
          <input type="checkbox" name="isnews" value="1" checked>
          создавать &quot;Архив&quot;</font></p>
      </td>
    </tr>
    <tr> 
      <td width="300" bordercolor="#333333" valign="middle" height="14"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Длина страницы раздела:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" valign="top" height="14"> 
        <p><font size="2"> 
          <input type="text" name="pglen" value="<%=isFirst?"50":Request.Form("pglen")%>" maxlength="20">
          публикаций</font></p>
      </td>
    </tr>
    <tr> 
      <td width="300" bordercolor="#333333" valign="middle" height="2"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Редактор рубрики:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="2"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" valign="top" height="2"> 
        <p><font size="2"> 
          <select name="redactor">
            <%
			Records.Source="Select ID, NAME from SMI_USR where USR_TYPE_ID<3 and SMI_ID="+smi_id+" order by NAME"
			Records.Open()
			while(!Records.EOF){%>
            <option value="<%=Records("ID").Value%>" 
				<%=!isFirst&&(Records("ID").Value==Request.Form("redactor"))?"selected":""%>> 
            <%=Records("NAME").Value%> </option>
            <%	Records.MoveNext()
			}
			Records.Close()
			%>
          </select>
          </font></p>
      </td>
    </tr>
  </table>
  <p align="center"><font color="#FF0000">Параметры выделенные красным цветом 
    обязательны к заполнению!</font></p>
  <p align="center"> 
    <input type="submit" name="Submit" value="Сохранить">
  </p>
  <hr width="780">
</form>
<%} 
else 
{%>
<center>
  <h1><font color="#3333FF">Рубрика добавлена!</font></h1>
</center>
<%}%>
<p>&nbsp;</p>
<p align="center"><a href="index.asp">Вернуться к публикациям</a></p>
<p align="center">&nbsp;</p>
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
