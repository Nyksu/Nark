<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var usok=false
var id_usr=0
var creater=""
var imgname=""
var pid=0
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var hid=0
var hdd=0
var path=""
var tpm=1000
var name_mem=""
var bnm=""
var bpos=""
var bid=0

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="area.asp"
		Response.Redirect("logpub.asp")
	}
	tpm=Session("tip_mem_pub")
	//if (tpm<5) {usok=true}
	id_usr=TextFormData(Session("id_mem_pub"),"0")
	name_mem=Session("name_mem_pub")
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			tpm=0
			name_mem=Session("name_mem")
		}
		Records.Close()
	} else {
		tpm=0
		name_mem=Session("name_mem")
	}
}

if (tpm>=7) {Response.Redirect("index.asp")}

%>


<html>
<head>
<title>Администрирование сайта.</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF"> 
<tr> <td bgcolor="#CCCCCC" bordercolor="#333333"> <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На 
главную страницу</a> | <A HREF="area.asp">В кабинет администратора</A></font></p></td></tr> 
</table><div align="center"> <h1>&nbsp;</h1></div><h1 align="center"><b>Остановленные 
информационные материалы</b></h1><%
var recs=CreateRecordSet()
Records.Source="Select t1.* from publication t1, heading t2 where t1.state=0 and t1.heading_id=t2.id and t2.smi_id="+smi_id+" order by date_create"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	cdat=Records("DATE_CREATE").Value
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgname=PubImgPath+pid+".gif"
    if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
	if (imgname=="") {
		imgname=PubImgPath+pid+".jpg"
		if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
	}
	creater="Администратор"
	recs.Source="Select * from smi_usr where id="+String(Records("CREATER_ID").Value)
	recs.Open()
	if (!recs.EOF) {
		creater=String(recs("NAME").Value)
	}
	recs.Close()
	path=""
	hid=String(Records("HEADING_ID").Value)
	hdd=hid
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
	}
%> <table width="100%" border="1" bordercolor="#FFFFFF" align="center"> <tr bordercolor="#333333" bgcolor="#CCCCCC"> 
<td colspan="2"> <p><b><font color="#FFFF00"> <font color="#FFFFFF">|| <a href="pubresume.asp?pid=<%=pid%>&st=1"><FONT COLOR="#006600">Опубликовать</FONT></a> 
|| <a href="delpub.asp?pid=<%=pid%>">Удалить</a> || <a href="removepub.asp?pid=<%=pid%>&hid=<%=hid%>">ПЕРЕМЕСТИТЬ</a> 
|| <a href="edpub.asp?pid=<%=pid%>">ИЗМЕНИТЬ</a> || <a href="bloknews.asp?pid=<%=pid%>">Разместить 
в блок</a></font></font></b></p></td></tr> <tr bgcolor="#FFFFFF" bordercolor="#CCCCCC"> 
<td colspan="2"> <h1><%=pname%></h1><p>Рубрика &gt; <%=path%></p></td></tr> <tr bordercolor="#CCCCCC"> 
<td> <p><font color="#000000">Дата создания: <b><%=cdat%></b> || Дата выхода<b> 
: <%=pdat%> || </b>Разместил:<b> <%=creater%></b> <b>|| </b>Автор:<b> <%=autor%></b></font></p></td><td> 
<div align="center"> <p><b>Фото Миниатюра</b></p></div></td></tr> <tr valign="top"> 
<td bordercolor="#CCCCCC" width="81%"> <p><font size="2" >Дайджест: <%=digest%></font> 
<a href="<%=url%>?pid=<%=pid%>">Просмотр текста</a></p><p>&nbsp; </p><p><FONT COLOR="#FF0000">Размещено 
в Блоках:&nbsp;<font size="1" color="#0000FF"> <%
	  recs.Source="Select t2.id, t1.posit, t2.subj from news_pos t1, block_news t2 where t1.block_news_id=t2.id and t1.publication_id="+pid
	  recs.Open()
	  while (!recs.EOF ) {
		bnm=TextFormData(recs("SUBJ").Value,"")
		bpos=String(recs("POSIT").Value)
		bid=String(recs("ID").Value)
	  %> &nbsp;<a href="block.asp?bk=<%=bid%>"><%=bnm%></a>&nbsp;(<%=bpos%>)&nbsp;| 
<%
	  recs.MoveNext()
	  }
	  recs.Close()
	  %> </font></FONT></p></td><td bordercolor="#CCCCCC" width="20%"> <div align="center"> 
<p> <%if (imgname != "") {%> <img src="<%=imgname%>" > <%}else{%> No image <%}%> 
</p><p>| <A HREF="addpubimg.asp?pid=<%=pid%>">Добавить/заменить иллюстрацию (фотографию)</A> 
| </p></div></td></tr> </table>
<%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
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
  </body>
</html>
