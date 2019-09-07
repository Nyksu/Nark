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

var hid=parseInt(Request("hid"))
if (isNaN(hid)) {Response.Redirect("index.asp")}

var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}

if (hid==0) {Response.Redirect("index.asp")}

var hname=""
var hiname=""
var url=""
var pid=0
var pname=""
var pdat=""
var autor=""
var digest=""
var sminame=""
var period=0
var pos=0
var lpag=0
var pp=0
var hdd=0
var path=""
var nm=""
var hadr=""
var imgname=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var nid=0
var name=""
var ndat=""
var nadr=""
var per=0
var kvopub=0
var isnews=1
var inews=1
var ishtml=0
var tpm=1000
var usok=false
var blokname=""

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

tit=sminame

Records.Source="Select * from heading where id="+hid+" and smi_id="+smi_id
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("index.asp")
}
isnews=Records("ISNEWS").Value
Records.Close()


hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		path=nm+"   "+path
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
		lpag=Records("PAGE_LENGTH").Value
		isnews=Records("ISNEWS").Value
	}
	else {
		path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> > "+path
	}
	hdd=Records("HI_ID").Value
	Records.Close()
}

path="<a href=\"/\">"+"На главную"+"</a> > "+path
tit+=" | "+hiname

if (pg>0) {pp=pg*lpag}
%>

<html>
<head>
<title><%=tit%> - Тюмень</title>
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
    <td bgcolor="#CCCCCC"> 
      <p>&nbsp;<a href="/" class="web">На главную</a> | <a href="area.asp" class="web">В 
        кабинет редактора</a> | <a href="addpub.asp?hid=<%=hid%>" class="web">+Добавить 
        публикацию в раздел &quot;<%=hiname%>&quot;</a></p>
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
    <td valign="top"> 
      <table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" width="96%">
        <tr> 
          <td valign="middle" height="45" align="center"> 
            <p align="left">&nbsp;&nbsp;<%=path%> 
              <%if (isnews==1) {%>
              &lt; <a href="archive.asp?hid=<%=hid%>&pg=<%=pg%>">Архив</a> 
              <%}%>
              | <a href="message.asp">Ваш вопрос по теме</a> </p>
          </td>
        </tr>
        <tr> 
          <td height="56" valign="middle" align="center"> 
		  <%if (tpm<7) {%>
	  <form>
              <div align="right">
                <input type="hidden" name="Choose WEB-site">
                <select name="list" size="1">
                  <option selected value="#">Работа с текущим разделом</option>
				  <option value="#">---------------</option>
                  <option VALUE="addpub.asp?hid=<%=hid%>">Добавить публикацию</option>
                  <option VALUE="ednewsheading.asp?hid=<%=hid%>">Настройки 
              раздела</option>
                  <option VALUE="addnewsheading.asp?hid=<%=hid%>">Создать подраздел</option>
                  <option VALUE="delpubheading.asp?hid=<%=hid%>">Удалить раздел</option>
                  </select>
                <input type="BUTTON" value="&gt;&gt;&gt;" border="0" align="top"
onclick="top.location.href=this.form.list.options[this.form.list.selectedIndex].value">
              </div>
            </form>
            <%}%>
            <p> 
              <%
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id="+hid+" and smi_id="+smi_id+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date<='TODAY' order by public_date desc, id desc"
	}
	recs.Open()
	if (!recs.EOF) {
		nid=String(recs("ID").Value)
		name=String(recs("NAME").Value)
		nadr=TextFormData(recs("URL").Value,"newshow.asp")
		nadr+="?pid="+nid
		ndat=recs("PUBLIC_DATE").Value
	} else {
		nid=0
		name=""
		nadr=""
		ndat=""
	}
	recs.Close()
	kvopub=0
	if (name!="") {
		recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
              <a  href="<%=url%>"><%=hname%></a> | 
              <%} 
Records.Close()
delete recs
%>
            </p>
            <%
if (isnews==1){
Records.Source="Select * from publication where heading_id="+hid+"and state=1 and public_date>='TODAY'-"+period+" and public_date<='TODAY' order by public_date desc, id desc"
} else {
Records.Source="Select * from publication where heading_id="+hid+"and state=1 and public_date<='TODAY' order by public_date desc, id desc"
}
Records.Open()
while (!Records.EOF && pos<=lpag*(1+pg))
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	if  (pos>=pp) {
		imgname=PubImgPath+pid+".gif"
    	if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
		if (imgname=="") {
			imgname=PubImgPath+pid+".jpg"
			if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
		}
%>
            <table width="100%" border="0" bordercolor="#FFFFFF" align="center" cellspacing="3" class="base_text">
              <tr valign="top" bordercolor="#FFFFFF"> 
                <td align="center"> 
                  <%if (imgname != "") {%>
                  <a href="<%=url%>"><img src="<%=imgname%>" border="0" alt="<%=pname%>" ></a> 
                  <%}else{%>
                  <%}%>
                </td>
                <td bordercolor="#FFFFFF" valign="top"> 
                  <p><b><a href="<%=url%>"><%=pname%></a></b> 
                    <br><%if (usok) {%>
                    <a href="pubresume.asp?pid=<%=pid%>&st=0"><font color="#006600">Остановить 
                    публикацию</font></a> | <a href="delpub.asp?pid=<%=pid%>"><font color="#006600">Удалить 
                    публикацию</font></a>|<br>
                    <a href="bloknews.asp?pid=<%=pid%>"><font color="#006600">Разместить 
                    в блок</font></a> | <a href="edpub.asp?pid=<%=pid%>"><font color="#006600">Редактировать</font></a> 
                  </p>
                  <p> 
                    <%}%>
                    [<%=pdat%>] <%=digest%> [<a href="<%=url%>">далее</a>]</p>
                </td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> 
                 </td>
              </tr>
            </table>
            <%
	}
	Records.MoveNext()
	pos+=1
} 
Records.Close()
%>
            <hr NOSHADE width="300" size="1">
            <table width="100%" border="0" cellspacing="0"  align="center">
              <tr> 
                <td> 
                  <p align="CENTER"> 
                    <%if (pg>0) {%>
                    <a href="pubheading.asp?hid=<%=hid%>&pg=<%=pg-1%>">Предыдущая 
                    страница</a> 
                    <%
  } 
  if (pos>lpag*(1+pg)) {
  %>
                    <a href="pubheading.asp?hid=<%=hid%>&pg=<%=pg+1%>">Следующая 
                    страница</a> 
                    <%}%>
                  </p>
                </td>
              </tr>
            </table>
          <%if (tpm<7) {%>
            <div align="center"> 
              <p><font size="1"><a href="ednewsheading.asp?hid=<%=hid%>">Настройки 
                раздела</a> | <a href="delpubheading.asp?hid=<%=hid%>">Удалить 
                раздел</a> | <a href="addnewsheading.asp?hid=<%=hid%>">Создать 
                подраздел</a> + <a href="addpub.asp?hid=<%=hid%>">Добавить публикацию</a> 
                </font></p>
            </div><%}%>
              </td>
        </tr>
      </table>
    </td>
    <td bgcolor="#FFFFFF" width="1" valign="top" align="center" background="images/line1-h.gif"></td>
    <td bgcolor="#FCFEFC" width="178" valign="top" align="center"> 
      <table border="0" cellspacing="5" cellpadding="0" width="100%">
        <tr> 
          <td valign="top"> 
<p><b><%
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
              <%if (hdd != hid){%>
             <a  href="<%=url%>" class="nav"> 
              <%}else{%>
              <font color="#009933">
              <%}%>
              <%=hname%> 
              <%if (hdd != hid){%>
              </a><br>
              <%} else {%>
              </font><br>
              <%}%>
              <%
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
              <%if (hdd != hid){%>
              <a  href="<%=url%>" class="nav"> 
              <%}else{%>
              <font color="#009933">
              <%}%>
              <%=hname%> 
              <%if (hdd != hid){%>
              </a><br>
              <%} else {%>
              </font><br>
              <%}%>
              <%
} 
Records.Close()
%><a href="message.asp" class="nav">Адрес доверия</a></b><br>
<a href="mailto:nark@tmn.ru">nark@tmn.ru</a></b></p>
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
            <p><%=blokname%></p></td>
          <td> </td>
        </tr>
      </table>
      <table border="0" cellspacing="5" cellpadding="0" width="100%">
        <tr> 
          <td valign="middle" height="80"> 
            <p><%
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
%></p>
          </td>
        </tr>
      </table>
      <hr noshade size="1" width="120"><div align="center"><img src="http://counter.rambler.ru/top100.cnt?527741" alt="" width=1 height=1 border=0> 
        <a href="http://top100.rambler.ru/top100/"> <img src="http://top100-images.rambler.ru/top100/banner-88x31-rambler-green2.gif" alt="Rambler's Top100" width=88 height=31 border=0></a> 
      </div>
    </td>
    <td width="1" height="271" bgcolor="#006600"></td>
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
%>
        <%if (hdd != hid){%>
        <a  href="<%=url%>" class="web"> 
        <%}else{%><font class="web"><font color="#000000">
        <%}%>
        <%=hname%><%if (hdd != hid){%></a> | 
        <%} else {%></font></font>
        | 
        <%}%>
        <%
} 
Records.Close()
%>
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
        <%if (hdd != hid){%>
        <a  href="<%=url%>" class="web"><%}else{%><font class="web"><font color="#000000"><%}%>
        <%=hname%><%if (hdd != hid){%></a> | 
        <%} else {%></font></font>
        | 
        <%}%>
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
