<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\path.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var hid=0
var hname=""
var url=""
var nid=0
var name=""
var ndat=""
var nadr=""
var per=0
var kvopub=0
var pname=""
var pdat=""
var autor=""
var digest=""
var imgLname=""
var imgname=""
var path=""
var hdd=0
var hadr=""
var nm=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var isnews=1
var ishtml=0
var blokname=""
var tpm=1000
var usok=false
var sminame=""

var tps=parseInt(Request("tps"))
var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
if (isNaN(tps)) {tps=1}
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}

if (String(Session("id_mem"))=="undefined") {
	if (Session("tip_mem_pub")<3) {usok=true}
	tpm=Session("tip_mem_pub")
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			tpm=0
		}
		Records.Close()
	} else {
		usok=true
		tpm=0
	}
}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

%>
<html>
<head>
<title><%=sminame%> - Тюмень</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-style: normal; font-weight: 400}
.link {  font-family: Arial, Helvetica, sans-serif; color: #009933; text-decoration: none}
h1 {  color: #009933; font-family: Arial, Helvetica, sans-serif; font-size: 18px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-style: normal; font-weight: 400 ; color: #009933; line-height: 18px}
.web { font-family: Arial, Helvetica, sans-serif; color: #0000FF; text-decoration: none ; font-size: 12px}
.nav { font-family: Arial, Helvetica, sans-serif; font-size: 13px; font-weight: bold; color: #006600}
-->
</style>
</head>

<body bgcolor="#F5FCF5" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/bg.gif">
<%if (usok) {%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td bgcolor="#CCCCCC" height="25"> 
      <p>&nbsp;<a href="area.asp" class="web">В кабинет редактора</a> | <a href="addnewsheading.asp?hid=<%=hid%>" class="web">+Добавить 
        раздел сайта</a></p>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#333333" height="1"></td>
  </tr>
</table>
<%}%>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td background="images/fon_green.gif" width="175" align="center" bgcolor="#006600"><img src="images/office.jpg" width="165" border="0"></td>
    <td width="7" bgcolor="#990000"><img src="images/trans_g_r.gif" width="7" height="180"></td>
    <td background="images/fon_red.gif" valign="middle" align="center" bgcolor="#990000"><img src="images/tit.gif" width="366" height="84" alt="Управление Госнаркоконтроля по Тюменской области"></td>
    <td width="7" bgcolor="#990000"><img src="images/trans_r_g.gif" width="7" height="180"></td>
    <td background="images/fon_green.gif" width="175" align="center" bgcolor="#006600"><img src="images/orel.gif" width="157" height="150"></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="780" align="center">
  <tr> 
    <td width="1" bgcolor="#006600" valign="bottom"> 
 </td>
    <td bgcolor="#FFFFFF" valign="top"> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="50"></td>
          <td bgcolor="#FFFFFF" valign="top"> 
            <table border="0" cellspacing="0" cellpadding="0" width="539" align="center">
              <tr> 
                <td width="539" height="18"></td>
              </tr>
              <tr> 
                <td width="539"> 
                  <%
// В переменной bk содержится код блока новостей
var bk=100
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
                  <h1><b><%=blokname%></b></h1>
                </td>
              </tr>
              <tr> 
                <td bgcolor="#006600" height="1" width="539"></td>
              </tr>
            </table>
            <%
// В переменной bk содержится код блока новостей
var bk=100
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgLname=PubImgPath+"l"+pid+".gif"
    if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
	if (imgLname=="") {
		imgLname=PubImgPath+"l"+pid+".jpg"
		if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
	}
	path=""
	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
var news=""
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
}

%>
            <table border="0" cellspacing="9" cellpadding="0" align="left">
              <tr> 
                <td> 
                  <%if (imgLname != "") {%>
                  <a href="newshow.asp?pid=<%=pid%>" class="nav"><img src="<%=imgLname%>" border="1" alt="<%=pname%>" class="photo" width="200"></a> 
                  <%}else{%>
                  &nbsp; 
                  <%}%>
                </td>
              </tr>
            </table>
            <table border="0" cellspacing="9" cellpadding="0">
              <tr> 
                <td valign="top" height="150"> 
                  <h2><font color="#000000"><b><%=pname%></b></font></h2>
                  <p><%=digest%><br>
                    <a href="newshow.asp?pid=<%=pid%>"><font size="2">Подробно</font></a><font size="2"> 
                    | <a href="print.asp?pid=<%=pid%>" target="_blank">Для печати</a></font></p>
                </td>
              </tr>
            </table>
            <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
            <table border="0" cellspacing="0" cellpadding="0" width="539" align="center">
              <tr> 
                <td bgcolor="#F5FCF5" background="images/line1.gif" height="1"></td>
              </tr>
            </table>
            <%if (usok) {%>
            <b><a href="addnewsheading.asp?hid=<%=hid%>" class="link">&nbsp;<br>
            +<img src="images/dot.gif" width="16" height="16" align="absmiddle" border="0" alt="Добавить новый информационный раздел сайта"> 
            Добавить раздел сайта</a></b> 
            <%}%>
          </td>
        </tr>
        <tr> 
          <td width="50"></td>
          <td align="right"> 
            <%
// маркек признака новостей
isnews=1
// если необходимо вывести рубрики не новостей то установить в ноль
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by sequence, name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date<='TODAY' order by public_date desc, id desc"
	}
	recs.Open()
	if (!recs.EOF) {
		nid=String(recs("ID").Value)
		name=String(recs("NAME").Value)
		digest=String(recs("DIGEST").Value)
		nadr=TextFormData(recs("URL").Value,"newshow.asp")
		nadr+="?pid="+nid
		ndat=recs("PUBLIC_DATE").Value
	} else {
		nid=0
		name=""
		digest=""
		nadr=""
		ndat=""
}
	recs.Close()
	kvopub=0
	recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	while (String(kvopub).length<4) {kvopub="0"+String(kvopub)}
	Records.MoveNext()
%>
            <table border="0" cellspacing="2" cellpadding="0" width="539">
              <tr> 
                <td width="539" height="18"></td>
              </tr>
              <tr> 
                <td width="539"><a href="<%=url%>" class="link"><img src="images/dot.gif" width="16" height="16" border="0" align="absmiddle" alt="В раздел <%=hname%>"> 
                  <b><%=hname%></b></a></td>
              </tr>
              <tr> 
                <td bgcolor="#006600" height="1" width="539"></td>
              </tr>
            </table>
            <table border="0" cellspacing="0" cellpadding="0" width="518">
              <tr> </tr>
              <tr> 
                <td height="10" width="539"> 
                  <p><b><%=name%></b></p>
                </td>
              </tr>
              <tr> 
                <td height="10" width="539"> 
                  <p>[<%=ndat%>] <%=digest%> <a href="<%=nadr%>"><font size="2">Подробно</font></a></p>
                </td>
              </tr>
            </table>
            <%
} Records.Close()
delete recs
%>
            <br>
          </td>
        </tr>
      </table>
    </td>
    <td width="1" valign="top" bgcolor="#EAF4EB" background="images/line1-h.gif"></td>
    <td width="178" valign="top" bgcolor="#FCFEFC"> 
      <table border="0" cellspacing="5" cellpadding="0" width="100%" align="center">
        <tr> 
          <td> 
            <p align="left"><b><a href="org.asp" class="nav">Наши реквизиты</a><br>
              <%
inews=0
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+inews+" order by sequence, name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%><a  href="<%=url%>" class="nav"><%=hname%></a><br>
<%
} 
Records.Close()
%>
      <a href="message.asp" class="nav">Адрес доверия</a></b><br>
<a href="mailto:nark@tmn.ru">nark@tmn.ru</a></p>
            <hr noshade size="1" width="90%" align="left">
            <div align="left"> 
              <p align="center"> 
                <%
// В переменной bk содержится код блока новостей
var bk=101
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
                <%=blokname%></p>
              <p>
                <%
// В переменной bk содержится код блока новостей
var bk=101
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgname=PubImgPath+pid+".gif"
    if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
	if (imgname=="") {
		imgname=PubImgPath+pid+".jpg"
		if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
	}
	path=""
	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
var news=""
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
}

%>
                <img src="images/link.gif" align="absmiddle"><a href="<%=url%>" target="blank" class="web"><%=pname%></a><br>
                <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
            </p></div>
          </td>
        </tr>
      </table>
      <div align="center"><img src="http://counter.rambler.ru/top100.cnt?527741" alt="" width=1 height=1 border=0> 
        <a href="http://top100.rambler.ru/top100/"> <img src="http://top100-images.rambler.ru/top100/banner-88x31-rambler-green2.gif" alt="Rambler's Top100" width=88 height=31 border=0></a> 
      </div>
</td>
    <td valign="top" bgcolor="#006600" width="1"></td>
  </tr>
</table>
<div align="center"></div>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td background="images/bott.gif" align="right" width="778" height="19" bgcolor="#006600"><a href="area.asp"><img src="images/bott.gif" width="20" height="19" border="0"></a></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td align="center" width="778" height="19" bgcolor="#FFFFFF"> 
      <p> 
        <%
inews=1
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+inews+" order by sequence, name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%><a  href="<%=url%>" class="web"> 
      <%=hname%><%if (hdd != hid){%></a> | <%}%>
        <%
} 
Records.Close()
%>
  </a>   |  <%
inews=0
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+inews+" order by sequence, name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%><a  href="<%=url%>" class="web"><%=hname%></a> |
<%
} 
Records.Close()
%>
      </p>
    </td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td align="center" height="1" bgcolor="#006600"></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center"></td>
    <td align="center" height="6"> </td>
    <td width="1" align="center"></td>
  </tr>
</table>
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td align="center" width="780" height="1" bgcolor="#006600"></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center"></td>
    <td align="center" width="778" height="6"> </td>
    <td width="1" align="center"></td>
  </tr>
</table>
<div align="center"></div>
<div style="FONT-FAMILY: 'MS Sans Serif', Geneva, sans-serif; FONT-SIZE: 8px;>
<div align="center"> 
  <div align="center">copyright © Группа общественных связей Управления Госнаркоконтроля 
    РФ по Тюменской области<br>
    Перепечатка материалов возможна только со ссылкой на авторов сайта </div>
</div>
</body>
</html>
