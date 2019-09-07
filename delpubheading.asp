<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\Path.inc" -->
<!-- #include file="inc\err.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var id=parseInt(Request("hid"))
if (isNaN(id)) {Response.Redirect("index.asp")}

var name=""
var hid=0
var retname="Вернутся в кабинет"
var sql=""
var ShowForm=true
var isok=true
var path=""
var hdd=0
var hadr=""
var nm=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var filename=""
var pic1=""
var pic2=""
var pic3=""
var pid=0

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="delpub.asp?pid="+id
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<3) {usok=true}
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

if (!usok) {Response.Redirect("pubarea.asp")}

sql="Select * from heading where id="+id+" and smi_id="+smi_id
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("pubarea.asp")
}
hid=Records("HI_ID").Value
name=String(Records("NAME").Value)
Records.Close()

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> | "+path
	hdd=Records("HI_ID").Value
	Records.Close()
}
if (path=="") {path="<a href=\"index.asp\">СМИ</a>"}

if (String(Request.Form("next1")) != "undefined") {
	if (Request.Form("agr")==1) {
		sql=allpubheading
		sql=sql.replace("%SMI",smi_id)
		sql=sql.replace("%HID",id)
		sql=sql.replace("%HID1",id)
		Records.Source=sql
		Records.Open()
		while (!Records.EOF) {
			pid=Records("ID").Value
			filename=PubFilePath+pid+"pub"
			if (!fs.FileExists(filename)) { filename="" }
			pic1=PubFilePath+pid+".gif"
			if (!fs.FileExists(pic1)) {pic1="" }
			if (pic1=="") {
				pic1=PubFilePath+pid+".jpg"
				if (!fs.FileExists(pic1)) { pic1="" }
			}
			pic2=PubFilePath+"l"+pid+".gif"
			if (!fs.FileExists(pic2)) {pic2="" }
			if (pic2=="") {
				pic2=PubFilePath+"l"+pid+".jpg"
				if (!fs.FileExists(pic2)) { pic2="" }
			}
			pic3=PubFilePath+"r"+pid+".gif"
			if (!fs.FileExists(pic3)) {pic3="" }
			if (pic3=="") {
				pic3=PubFilePath+"r"+pid+".jpg"
				if (!fs.FileExists(pic3)) { pic3="" }
			}
			if (filename!="") {fs.DeleteFile(filename)}
			if (pic1!="") {fs.DeleteFile(pic1)}
			if (pic2!="") {fs.DeleteFile(pic2)}
			if (pic3!="") {fs.DeleteFile(pic3)}
			Records.MoveNext()
		}
		Records.Close()
		sql="Delete from heading where id="+id
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
			var ShowForm=false
		}
	} else {Response.Redirect("pubarea.asp")}
}

%>

<Html>
<Head>
<Title>Удаление рубрики <%=name%></Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" bordercolor="#000000">
  <tr> 
    <td bordercolor="#CCCCCC"> 
      <p><a href="pubarea.asp"><b>Кабинет редактора</b></a> | <a href="index.asp">Редактировать 
        издание</a> | <a href="bloknews.asp">Информационные блоки</a> </p>
    </td>
  </tr>
</table>
<h1 align="center"><font size="3">Удаление рубрики публикаций</font></h1>
<p>Путь&gt; <font color="#FF0000"><%=path%></font></p>
<% if (ShowForm) {%>
<p align="center"><b>Вы действительно хотите удалить рубрику: <font color="#FF0000"><%=name%></font> ?</b></p>
<p>&nbsp;</p>
<p align="center"><b><font size="3" color="#FF0000">&nbsp;&nbsp;Внимание!</font></b> 
  Удаление рубрики!</p>
<p align="center">Удаление рубрики необратимо и ее, а так же публикации в нее 
  входящие, невозможно будет востановить!</p>
<p align="center">Если Вы решили удалить рубрику, то поставьте флажок в соответствующем 
  поле и нажмите кнопку &quot;Продолжить&quot;.</p>
<form name="form1" method="post" action="delpubheading.asp">
  <input type="hidden" name="hid" value="<% =id %>">
  <p align="center"> 
    <input type="checkbox" name="agr" value="1">
    Да, я хочу удалить эту рубрику: (<b><font color="#FF0000" size="3"><%=name%></font></b>) !</p>
  <p align="center">&nbsp;</p>
  <p align="center"> 
    <input type="submit" name="next1" value="Продолжить">
  </p>
</form>
<%} else {%>
<p>&nbsp;</p>
<p align="center"><font color="#FF0000"><b>Рубрика удалена!</b></font></p>
<%}%>
<p align="center"><a href="index.asp">Вернуться к публикациям</a></p>
<p align="center"><a href="pubarea.asp">Кабинет редактора</a></p>
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
</Body>
</Html>
