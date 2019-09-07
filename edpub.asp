<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var pid=parseInt(Request("pid"))
if (isNaN(pid)) {pid=0}
if (pid==0){Response.Redirect("index.asp")}

// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var name=""
var autor=""
var pdate=""
var url=""
var digest=""
var coment=""
var ishtml=0
var news=""
var hid=0
var sql=""
var stt=1000
var creator=0
var sminame=""
var hiname=""
var pstat=0
var usok=false
var id_usr=0
var isFirst=true
var ErrorMsg=""
var hiadr=""
var ShowForm=true
var filename=""
var ts=""
var str=""
var ErrorMsg=""


sql="Select t1.*, t2.name as hname, t2.url as hurl from PUBLICATION t1, HEADING t2 where t1.id="+pid+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Source=sql
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("index.asp")
}
name=String(Records("NAME").Value)
hiname=String(Records("HNAME").Value)
digest=String(Records("DIGEST").Value)
pdate=Records("PUBLIC_DATE").Value
pstat=Records("STATE").Value
url=TextFormData(Records("URL").Value,"")
coment=TextFormData(Records("COMENT").Value,"")
autor=TextFormData(Records("AUTOR").Value,"")
creater=Records("CREATER_ID").Value
ishtml=Records("ISHTML").Value
hid=Records("HEADING_ID").Value
hiadr=TextFormData(Records("HURL").Value,"pubheading.asp")
hiadr+="?hid="+hid
Records.Close()

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="addpub.asp?hid="+hid
		Response.Redirect("logpub.asp")
	}
	stt=Session("tip_mem_pub")
	id_usr=TextFormData(Session("id_mem_pub"),"0")
	if (stt<7 && creater==id_usr && pstat==0) {usok=true}
	if (stt<5) {usok=true}
	if (stt==3 && pstat>0) {usok=false}
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			stt=0
		}
		Records.Close()
	} else {
		usok=true
		stt=0
	}
}

if (!usok) {Response.Redirect("index.asp")}

var fs= new ActiveXObject("Scripting.FileSystemObject")
filename=PubFilePath+pid+".pub"
if (!fs.FileExists(filename)) { filename="" }
if (filename!="") {
	ts= fs.OpenTextFile(filename)
	str=ts.ReadAll()
	ts.Close()
	while (str.indexOf("<")>=0) {str=str.replace("<","&lt;")}
	news=str
}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

isFirst=String(Request.Form("Submit"))=="undefined"

if ( !isFirst ) {
	name=TextFormData(Request.Form("name"),"")
	autor=TextFormData(Request.Form("autor"),"")
	url=TextFormData(Request.Form("url"),"")
	pdate=TextFormData(Request.Form("pdate"),"")
	digest=TextFormData(Request.Form("digest"),"")
	news=TextFormData(Request.Form("news"),"")
	coment=TextFormData(Request.Form("coment"),"")
	ishtml=TextFormData(Request.Form("ishtml"),"0")
	
	// убрать коментарий, если надо запретить HTML  ******  while (news.indexOf("<")>=0) {news=news.rplace("<","&lt;")}
	while (name.indexOf("'")>=0) {name=name.replace("'","\"")}
	while (autor.indexOf("'")>=0) {autor=autor.replace("'","\"")}
	while (digest.indexOf("'")>=0) {digest=digest.replace("'","\"")}
	while (coment.indexOf("'")>=0) {coment=coment.replace("'","\"")}
	while (pdate.indexOf(",")>=0) {pdate=pdate.replace(",",".")}
	while (pdate.indexOf("/")>=0) {pdate=pdate.replace("/",".")}
	while (pdate.indexOf("-")>=0) {pdate=pdate.replace("-",".")}
	
	if (name.length<3) {ErrorMsg+="Слишком короткая тема.<br>"}
	if (name.length>100) {ErrorMsg+="Слишком длинная тема.<br>"}
	if (autor.length>100) {ErrorMsg+="Слишком длинное имя автора.<br>"}
	if (url.length>80) {ErrorMsg+="Слишком длинная URL.<br>"}
	if (coment.length>200) {ErrorMsg+="Слишком длинный коментарий.<br>"}
	if (digest.length>250) {ErrorMsg+="Слишком длинный дайджест.<br>"}
	if (digest.length<5) {ErrorMsg+="Слишком короткий дайджест.<br>"}
	if (news.length>120000) {ErrorMsg+="Слишком длинный текст публикации.<br>"}
	if (news.length<15) {ErrorMsg+="Слишком короткий текст публикации.<br>"}
	if (!/(\d{2}).(\d{2}).(\d{4})$/.test(pdate)) {ErrorMsg+="Неверный формат даты публикации.<br>"}
	
	if (ErrorMsg=="") {
		sql=edpub
		sql=sql.replace("%ID",pid)
		sql=sql.replace("%NAM",name)
		sql=sql.replace("%DIG",digest)
		sql=sql.replace("%PDAT",pdate)
		sql=sql.replace("%COM",coment)
		sql=sql.replace("%AUT",autor)
		sql=sql.replace("%URL",url)
		sql=sql.replace("%ISH",ishtml)
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
			ts=fs.OpenTextFile(filename,2,true)
			ts.Write(news)
			ts.Close()
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

// Убираем служебные символы
while (name.indexOf("\"")>=0) {name=name.replace("\"","&quot;")}
while (autor.indexOf("\"")>=0) {autor=autor.replace("\"","&quot;")}
while (coment.indexOf("\"")>=0) {coment=coment.replace("\"","&quot;")}
while (digest.indexOf("\"")>=0) {digest=digest.replace("\"","&quot;")}
while (name.indexOf("<")>=0) {name=name.replace("<","&lt;")}
while (autor.indexOf("<")>=0) {autor=autor.replace("<","&lt;")}
while (coment.indexOf("<")>=0) {coment=coment.replace("<","&lt;")}
while (digest.indexOf("<")>=0) {digest=digest.replace("<","&lt;")}

%>

<html>
<head>
<title>Редактирование публикации</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<LINK REL="stylesheet" HREF="style.css" TYPE="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#CCCCCC"> 
      <p><a href="index.asp">На главную страницу</a> <%=sminame%>  || <a href="<%=hiadr%>"><%=hiname%></a> 
      </p>
    </td>
  </tr>
</table>
<p align="left">&nbsp;&nbsp;<a href="area.asp">Кабинет</a></p>
<h1 align="center">Редактирование публикации</h1>
<p align="center"><b><%=name%></b></p>
<%if(ErrorMsg!=""){%>
<center>
  <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%>
<%if (ShowForm) {%>
<form name="form1" method="post" action="edpub.asp">
  <p align="center">
<input type="hidden" name="pid" value="<%=pid%>">
  </p>
  <table width="870" border="1" bordercolor="#FFFFFF" align="center">
    <tr> 
      <td width="250" bgcolor="#CCCCCC" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>Параметры:</b></p>
        </div>
      </td>
      <td width="6">&nbsp;</td>
      <td bgcolor="#CCCCCC" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>Значения:</b></p>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Заголовок публикации:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?name:Request.Form("name")%>" maxlength="100" size="45">
          до 250 символов</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2">Автор:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="autor" value="<%=isFirst?autor:Request.Form("autor")%>" maxlength="100" size="45">
          до 100 символов</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2">Внешний URL публикации:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="url" value="<%=isFirst?url:Request.Form("url")%>" maxlength="80" size="45">
          до 80 символов</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Дата публикации:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="pdate" value="<%=isFirst?pdate:Request.Form("pdate")%>" maxlength="10" size="45">
          можно указать будущую дату</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Дайджест:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="digest" value="<%=isFirst?digest:Request.Form("digest")%>" maxlength="190" size="45">
          до 250 символов</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" height="14" valign="top"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Текст публикации:</font><font size="2">&nbsp;&nbsp;</font></p>
          <p align="left">Для удобства заполнения текст можно заранее подготовить 
            в текстовом редакторе (Notepad, WORD и т.п.).</p>
          <p align="left">&nbsp;</p>
          <p align="left"><font size="1">Для наиболее правильного отображения 
            текста на сайте, необходимо удалить лишние возвраты абзацев (Enter).</font></p>
          <p align="left"><font size="1">Правильно отформатированный абзац выглядит 
            в стандартном блокноте как сплошная строка.</font></p>
        </div>
      </td>
      <td width="6" height="14" valign="top"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" height="14" valign="top" bgcolor="#CCCCCC"> 
        <textarea name="news" cols="50" rows="15"><%=news%></textarea>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" height="14" valign="middle"> 
        <div align="right"> 
          <p><font size="2">Коментарий:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" height="14" valign="top"> 
        <p> 
          <input type="text" name="coment" value="<%=isFirst?coment:Request.Form("coment")%>" maxlength="190" size="45">
        </p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" height="14" valign="middle"> 
        <div align="right"> 
          <p><font size="2">HTML текст</font><font size="2">:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" height="14" valign="top"> 
        <p> 
          <input type="checkbox" name="ishtml" value="1" <%if (ishtml==1){%>checked<%}%>>
          пометить, если текст публикации в формате HTML !!</p>
      </td>
    </tr>
  </table>
  <p align="center"><font color="#FF0000">Параметры выделенные красным цветом 
    обязательны к заполнению!</font></p>
  <p align="center"> 
    <input type="submit" name="Submit" value="Сохранить">
  </p>
</form>
<%} else {%>
<p align="center">-------------------------------</p>
<p align="center"><b>Изменения сохранены!</b></p>
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
<%}%>
</body>
</html>
