<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var pid=parseInt(Request("pid"))
if (isNaN(pid)) {Response.Redirect("index.asp")}

var hid=0
var sminame=""
var tit=""
var hdd=0
var nm=""
var hiname=""
var period=0
var path=""
var hadr=""
var pname=""
var pdat=""
var autor=""
var news=""
var imgname=""
var imgLname=""
var imgRname=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var sos=""
var sql=""
var ishtml=0
var usok=false
var id_usr=0
var creater=""
var crr=0
var bnm=""
var bpos=""
var bid=0

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="newsshow.asp?pid="+pid
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<7) {usok=true}
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

Records.Source="Select t1.* from publication t1, heading t2 where t1.id="+pid+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("index.asp")
}
hid=Records("heading_id").Value
pname=String(Records("NAME").Value)
pdat=Records("PUBLIC_DATE").Value
autor=TextFormData(Records("AUTOR").Value,"")
ishtml=TextFormData(Records("ISHTML").Value,"0")
crr=parseInt(Records("CREATER_ID").Value)
Records.Close()
if (isNaN(pid)) {crr=0}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

Records.Source="Select * from smi_usr where id="+crr
Records.Open()
	if (!Records.EOF) {
		creater=String(Records("NAME").Value)
	} else {creater="Администратор"}
Records.Close()

tit=sminame

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
	}
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> | "+path
	hdd=Records("HI_ID").Value
	Records.Close()
}

var ddt = new Date()
var dt=""
var str=""
var sumdat=Server.CreateObject("datesum.DateSummer")

str=String(ddt.getMonth()+1)
if (str.length==1) {str="0"+str}
dt="."+str+"."+ddt.getYear()
str=String(ddt.getDate())
if (str.length==1) {str="0"+str}
dt=str+dt
dt=sumdat.SummToDate(dt,"-"+period)
if (sumdat.DateComparing(dt,pdat) > 0) {sos="В архиве"} else {sos="В публикации"}

path="<a href=\"index.asp\">"+sminame+"</a> | "+path
tit+=" | "+hiname

filnam=PubFilePath+pid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml==0) {
	while (!ts.AtEndOfStream){
		news+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}

imgname=PubImgPath+pid+".gif"
if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
if (imgname=="") {
	imgname=PubImgPath+pid+".jpg"
	if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
}

imgLname=PubImgPath+"l"+pid+".gif"
if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
if (imgLname=="") {
	imgLname=PubImgPath+"l"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
}

imgRname=PubImgPath+"r"+pid+".gif"
if (!fs.FileExists(PubFilePath+"r"+pid+".gif")) { imgRname="" }
if (imgRname=="") {
	imgRname=PubImgPath+"r"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"r"+pid+".jpg")) { imgRname="" }
}

%>

<html>
<head>
<title><%=tit%> | <%=pname%></title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p>&nbsp;<%=path%> <a href="area.asp">Вернуться в кабинет</a></p>
    </td>
  </tr>
</table>
<p align="center">Рубрика: <font color="#135086"><b><font size="3"><%=tit%></font></b></font></p>
<table width="780" border="1" bordercolor="#FFFFFF" align="center">
  <tr bgcolor="#CCCCCC" bordercolor="#333333"> 
    <td colspan="2"> 
      <p><b>|| <%=pname%> || <%=sos%> || Разместил: </b><%=creater%></p>
    </td>
  </tr>
  <tr> 
    <td bordercolor="#CCCCCC" width="590" height="65"> 
      <p><b>Дата публикации</b>: <%=pdat%><br>
        <b>Автор</b>: <%=autor%><br></p>
      <p><%=news%></p> </td>
    <td bordercolor="#CCCCCC" width="186" height="65" valign="top"> 
      <p align="center"><b>Миниатюра для дайджеста</b></p>
      <div align="center"> 
        <%if (imgname != "") {%>
        <img src="<%=imgname%>" > 
        <%}else{%>
        No image 
        <%}%>
      </div>
    </td>
  </tr>
</table>
<table width="780" border="1" bordercolor="#FFFFFF" align="center">
  <tr bordercolor="#333333" bgcolor="#CCCCCC"> 
    <td width="50%" valign="top"> 
      <p align="center"><b>Публикация размещена в блоках:</b></p>
    </td>
  </tr>
  <tr bordercolor="#CCCCCC"> 
    <td width="50%" valign="top" height="15"> 
      <p align="center"><font size="1" color="#0000FF"> 
        <%
	  Records.Source="Select t2.id, t1.posit, t2.subj from news_pos t1, block_news t2 where t1.block_news_id=t2.id and t1.publication_id="+pid
	  Records.Open()
	  while (!Records.EOF ) {
		bnm=TextFormData(Records("SUBJ").Value,"")
		bpos=String(Records("POSIT").Value)
		bid=String(Records("ID").Value)
	  %>
        &nbsp;<a href="block.asp?bk=<%=bid%>"><%=bnm%></a>&nbsp;(<%=bpos%>)<br>
        <%
	 Records.MoveNext()
	  }
	  Records.Close()
	  %>
        </font></p>
    </td>
  </tr>
</table>

<table width="780" border="1" bordercolor="#FFFFFF" align="center">
  <tr bordercolor="#333333" bgcolor="#CCCCCC"> 
    <td width="50%" valign="top"> 
      <p align="center"><b>Основная фотография</b></p>
    </td>
  </tr>
  <tr bordercolor="#CCCCCC"> 
    <td width="50%" valign="top"> 
      <div align="center"> 
        <%if (imgLname != "") {%>
        <img src="<%=imgLname%>" > 
        <%}else{%>
        No image 
        <%}%>
      </div>
    </td>
  </tr>
</table>
<table width="780" border="1" bordercolor="#FFFFFF" align="center">
  <tr bordercolor="#333333" bgcolor="#CCCCCC"> 
    <td width="50%" valign="top"> 
      <p align="center"><b>Дополнительная фотография</b></p>
    </td>
  </tr>
  <tr bordercolor="#CCCCCC"> 
    <td width="50%" valign="top"> 
      <div align="center"> 
        <%if (imgRname != "") {%>
        <img src="<%=imgRname%>" > 
        <%}else{%>
        No image 
        <%}%>
      </div>
    </td>
  </tr>
</table>
<p>&nbsp;</p><table width="100%" border="1" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p align="center"><a href="addpubimg.asp?pid=<%=pid%>">Добавить 
        фотографию</a> | <a href="area.asp">Вернуться в кабинет</a> | <a href="index.asp">Вернуться 
        к публикациям</a></p>
    </td>
  </tr>
</table>
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
