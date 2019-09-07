<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->

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
var ishtml=0
var isnews=1
var lgok=false
var usok=false
var bnm=""
var bpos=""
var bid=0

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
	if (Session("tip_mem_pub")<4) {lgok=true}
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			lgok=true
		}
		Records.Close()
	} else {
		usok=true
		lgok=true
	}
}


Records.Source="Select t1.* from publication t1, heading t2 where t1.state=1 and t1.id="+pid+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
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
Records.Close()

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

tit=sminame

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	isnews=Records("ISNEWS").Value
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
	}
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> > "+path
	hdd=Records("HI_ID").Value
	Records.Close()
}

var ddt = new Date()
var dt=""
var str=""
var sumdat=Server.CreateObject("datesum.DateSummer")
if ( isnews ) {
str=String(ddt.getMonth()+1)
if (str.length==1) {str="0"+str}
dt="."+str+"."+ddt.getYear()
str=String(ddt.getDate())
if (str.length==1) {str="0"+str}
dt=str+dt
dt=sumdat.SummToDate(dt,"-"+period)
if (sumdat.DateComparing(dt,pdat) > 0) {sos="В архиве"} else {sos="В публикации"}
}

path="<a href=\"/\">"+"На главную"+"</a> > "+path
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
<title><%=tit%> > <%=pname%> г. Тюмень</title>
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
      <p>&nbsp;<a href="/" class="web">На главную</a> | <a href="area.asp" class="web">В 
        кабинет редактора</a> | </p>
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
    <td background="images/fon_green.gif" width="175" align="center" bgcolor="#006600"><a href="/"><img src="images/office.jpg" width="165" border="0"></a></td>
    <td width="7" bgcolor="#990000"><img src="images/trans_g_r.gif" width="7" height="180"></td>
    <td background="images/fon_red.gif" valign="middle" align="center" bgcolor="#990000"><img src="images/tit.gif" width="366" height="84" alt="Управление Госнаркоконтроля по Тюменской области"></td>
    <td width="7" bgcolor="#990000"><img src="images/trans_r_g.gif" width="7" height="180"></td>
    <td background="images/fon_green.gif" width="175" align="center" bgcolor="#006600"><a href="/"><img src="images/orel.gif" width="157" height="150" border="0" alt="На главную страницу"></a></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" align="center">
  <tr> 
    <td width="1" valign="top" height="271" bgcolor="#006600"></td>
    <td width="50" valign="top">&nbsp;</td>
    <td valign="top" align="left"> 
      <p><br>
        <%=path%> < <a href="javascript:history.back(1)">Вернуться назад</a> </p>
            <table width="527" border="0" bordercolor="#FFFFFF" cellspacing="0" class="base_text">
        <tr valign="top" bordercolor="#FFFFFF"> 
          <td width="527"> 
            <h1><%=pname%></h1><% if (lgok) {%>
	  <form>
              <div align="right">
                <input type="hidden" name="Choose WEB-site">
                <select name="list" size="1">
                  <option selected value="#">Работа с текущей публикацией</option>
				  <option value="#">---------------</option>
                  <option VALUE="addpubimg.asp?pid=<%=pid%>">Добавить / заменить 
                  фото</option>
                  <option VALUE="edpub.asp?pid=<%=pid%>">Редактировать</option>
                  <option VALUE="bloknews.asp?pid=<%=pid%>">Разместить в блок</option>
                  <option VALUE="pubresume.asp?pid=<%=pid%>&st=0">Временно приостановить</option>
                  <option VALUE="delpub.asp?pid=<%=pid%>">Удалить</option>
                </select>
                <input type="BUTTON" value="&gt;&gt;&gt;" border="0" align="top"
onclick="top.location.href=this.form.list.options[this.form.list.selectedIndex].value">
              </div>
            </form><%}%>

          </td>
        </tr>
        <tr valign="top" bordercolor="#FFFFFF"> 
          <td height="55" width="527"> 
            <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&copy; <%=pdat%>&nbsp; <%=autor%><br>
              <a href="print.asp?pid=<%=pid%>" target="_blank"><img src="images/print.gif" width="18" height="16" align="absmiddle" border="0"> 
              Версия для печати</a> | <a href="message.asp">Задать вопрос</a></p>
            <%if (imgLname != "") {%>
            <img src="<%=imgLname%>" align="left" border="1" alt="<%=pname%>" > 
            <%}else{%>
            <%}%>
            <%=news%> </td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellspacing="0">
        <tr bordercolor="#CCCCCC"> 
          <td valign="top" height="40"> 
            <div align="center"> 
              <%if (imgRname != "") {%>
              <img src="<%=imgRname%>" border="1" > 
              <%}else{%>
              &nbsp; 
              <%}%>
            </div>
          </td>
        </tr>
      </table>
      <table border="0" cellspacing="2" cellpadding="2" width="100%">
        <tr> 
          <td bgcolor="#FFFFFF" nowrap> 
            </td>
        </tr>
      </table>
      <%if (usok) {%>
      <table width="100%" border="2" align="center" bordercolor="#FFFFFF">
        <tr> 
          <td height="2" bordercolor="#CCCCCC"> 
            <div align="center"> 
              <p><b><font color="#006633">Публикация размещена в блоках:</font></b></p>
            </div>
          </td>
        </tr>
        <tr> 
          <td height="21" bordercolor="#CCCCCC"> 
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
      <%}%>
      <div align="center"><font face="Arial, Helvetica, sans-serif"><font size="1">
        <% if (lgok) {%>
        | <a href="addpubimg.asp?pid=<%=pid%>"><font color="#006600">Добавить 
        фото</font></a></font> 
        <%}%>
        <%if (usok) {%>
        <b><font face="Arial, Helvetica, sans-serif" size="1">| <a href="pubresume.asp?pid=<%=pid%>&st=0"><font color="#006600">Остановить</font></a> 
        | <a href="delpub.asp?pid=<%=pid%>"><font color="#006600">Удалить</font></a> 
        | <a href="bloknews.asp?pid=<%=pid%>"><font color="#006600">Разместить 
        в блок</font></a> | <a href="edpub.asp?pid=<%=pid%>"><font color="#006600">Редактировать</font></a></font></b> 
        </font> 
        <%}%>
      </div>
    </td>
    <td bgcolor="#FFFFFF" width="1" valign="top" align="center" background="images/line1-h.gif"></td>
    <td bgcolor="#FCFEFC" width="178" valign="top" align="center"> 
      <table border="0" cellspacing="5" cellpadding="0" width="100%">
        <tr> 
          <td valign="top"> 
            <p><b> 
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
%>
              <a  href="<%=url%>" class="nav"> 
              <%=hname%> 
              </a><br><%
} 
Records.Close()
%><a href="org.asp" class="nav">Наши реквизиты</a><br>
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
%>
              <a  href="<%=url%>" class="nav"> 
              <%=hname%> 
              </a><br>
               <%
} 
Records.Close()
%><a href="message.asp" class="nav">Адрес доверия</a></b><br>
<a href="mailto:nark@tmn.ru">nark@tmn.ru</a>
              </b></p>
          </td>
        </tr>
      </table>
      <hr noshade size="1" width="120">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" height="12">
        <tr> 
          <td bgcolor="#FFFFFF" height="3"></td>
        </tr>
      </table>
      <table border="0" cellspacing="0" cellpadding="0" bordercolor="#666699">
        <tr> 
          <td align="center"> 
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
            <p><%=blokname%></p>
          </td>
          <td> </td>
        </tr>
      </table>
      <table border="0" cellspacing="5" cellpadding="0" width="100%">
        <tr> 
          <td valign="middle" height="80"> 
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
            </p>
          </td>
        </tr>
      </table>
      <hr noshade size="1" width="120"><div align="center"><img src="http://counter.rambler.ru/top100.cnt?527741" alt="" width=1 height=1 border=0> 
        <a href="http://top100.rambler.ru/top100/"><img src="http://top100-images.rambler.ru/top100/banner-88x31-rambler-green2.gif" alt="Rambler's Top100" width=88 height=31 border=0></a> 
      </div>
    </td>
    <td width="1" height="271" bgcolor="#006600"></td>
  </tr>
</table>
<font color="#FFFFFF"></font> 
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
%> | <a  href="<%=url%>" class="web"> 
      <%=hname%><%if (hdd != hid){%></a><%}%>
        <%
} 
Records.Close()
%> |
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
