<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\path.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=4
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
var blokname=""
var ishtml=0
var tpm=1000
var usok=false

sql="Select COUNT_MSG from GET_COUNT_MSG_ST(2,0)"
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	msgcount=Records("COUNT_MSG").Value
}
Records.Close()

if (String(Session("id_mem"))=="undefined") {
	if (Session("tip_mem_pub")<7) {stt=1}
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			stt=2
		}
		Records.Close()
	} else {
		usok=true
		stt=2
	}
}



%>

<html>
<head>
<title>Авто Тюмень - Тюменский Автомобильный Портал - Авто 72RUS -  Автомобили в Тюмени - Покупка - Продажа - Авто Дилеры</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="auto.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00"> 
<tr bgcolor="#666666"> <td valign="middle" height="25"> <div align="left"> <h1><b> 
<font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2">АВТО ТЮМЕНЬ 
- Тюменский Автомобильный Портал 72RUS.RU <!--begin of Rambler's Top100 code --> 
<a href="http://top100.rambler.ru/top100/"> <img src="http://counter.rambler.ru/top100.cnt?388244" alt="" width=1 height=1 border=0></a> 
<!--end of Top100 code--> </font></b></h1></div></td></tr> </table><table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" bgcolor="#333333"> 
<tr> <td width="150" valign="middle" height="3"> </td><td height="3"> </td><td width="468" height="3"> 
</td></tr> </table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" height="60">
  <tr> 
    <td height="120" background="head/fon.gif" valign="bottom" width="237" bgcolor="#000000"><img  src="head/board.gif" width="237" height="120" usemap="#Map2" border="0"></td>
    <td height="120" background="head/lin.gif" valign="bottom" align="right" bgcolor="#000000"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#000000" valign="middle" align="center" height="59" background="head/tyumen.jpg" width="280"><font color="#FFCC00"><b>Тюмень</b></font></td>
          <td bgcolor="#000000" height="59" align="right"> 
            <table width="468" border="0" cellspacing="0" height="60" cellpadding="0">
              <tr> 
                <%
// В переменной bk содержится код блока новостей
var bk=65
// Не забывать его менять!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	ishtml=TextFormData(Records("ISHTML").Value,"0")

news=""
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
%>
                <td align="CENTER"><%=news%></td>
                <%
Records.MoveNext()
} 
Records.Close()
%>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#000000" colspan="2"><img  src="head/city.gif" width="100%" height="12"></td>
        </tr>
        <tr align="left"> 
          <td colspan="2"> 
            <h1><img  src="head/audioCD.gif" width="112" height="43" usemap="#Map" border="0" alt="Магнитолы и сигнализации" align="absmiddle"> 
              <map name="Map"> 
                <area shape="rect" coords="72,18,106,37" href="auto_subj.asp?subj=141" alt="Атомагнитолы и Сигнализации" title="Атомагнитолы и Сигнализации">
              </map>
            </h1>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<map name="Map2"> 
  <area shape="circle" coords="101,125,99" href="index.asp" alt="В начало" title="В начало" target="_self">
</map>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" bgcolor="#333333">
  <tr> 
    <td width="150" valign="middle" height="4"> </td>
    <td height="4"> </td>
    <td width="468" height="4"> </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
  <tr> 
    <td bgcolor="#CCCCCC" valign="middle" height="20" align="left"> <a href="auto_subj.asp?subj=17"><img src="Sighns/z070403.gif" width="40" height="21" border="0" alt="Автотранспорт легковой импортный"></a><a href="auto_subj.asp?subj=14"><img src="Sighns/z071000.gif" width="40" height="21" border="0" alt="Автотранспорт легковой отечественный"></a><a href="auto_subj.asp?subj=50"><img src="Sighns/z070401.gif" width="40" height="21" border="0" alt="Грузовой транспорт"></a><a href="auto_subj.asp?subj=51"><img src="Sighns/z070404.gif" width="40" height="21" border="0" alt="Автобусы"></a><a href="auto_subj.asp?subj=133"><img src="Sighns/z070402.gif" width="40" height="21" border="0" alt="Прицепы и полуприцепы"></a><a href="auto_subj.asp?subj=19"><img src="Sighns/z070405.gif" width="40" height="21" border="0" alt="Специализированная техника"></a><a href="auto_subj.asp?subj=170"><img src="Sighns/z070406.gif" width="40" height="21" border="0" alt="Мотоциклы и мопеды"></a><a href="auto_subj.asp?subj=23"><img src="Sighns/z070502.gif" width="40" height="21" border="0" alt="Запчасти и авто товары"></a><a href="contact.asp"><img src="Sighns/post.gif" width="40" height="21" border="0" alt="Обратная связь"></a></td>
    <td width="430" height="20" align="right" bgcolor="#CCCCCC"> 
      <table width="430" border="1" cellspacing="2" cellpadding="0" bordercolor="#CCCCCC">
        <tr bordercolor="#000000" bgcolor="#FFFFFF"> 
          <td width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>АВТО 
              ТЮМЕНЬ</b></font></div>
          </td>
          <td width="82" height="20"> 
            <div align="center"><a href="auto_subj.asp"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>Объявления</b></font></a></div>
          </td>
          <td width="82" height="20"> 
            <div align="center"><a href="auto_penalty.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>Штрафы</b></font></a></div>
          </td>
          <td width="82" height="20"> 
            <div align="center"><a href="auto_regions.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>Регионы</b></font></a></div>
          </td>
          <td width="82" height="20"> 
            <div align="center"><a href="auto_dov.asp"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>Доверенность</b></font></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" bgcolor="#333333">
  <tr> 
    <td width="150" valign="middle" height="4"> </td>
    <td height="4"> </td>
    <td width="468" height="4"> </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" height="60">
  <tr> 
    <td width="237" align="left" valign="top" bgcolor="#CCCCCC"> 
      <div align="center"> 
        <table width="96%" border="1" cellspacing="2" cellpadding="0" bgcolor="#0033FF" bordercolor="#FFFFFF">
          <tr> 
            <td height="15"> 
              <p><font color="#FFFFFF"><b>Разделы сайта</b></font></p>
            </td>
          </tr>
        </table>
        <table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor="#990000"> 
            <td height="4"></td>
            <td width="50" height="4" align="right"></td>
          </tr>
          <tr> 
            <td height="22" bgcolor="#E8ECEC"> 
              <p><b><img src="Img/arrow1.gif" width="12" height="11" align="absmiddle"> 
                <a href="auto_subj.asp?subj=2">Авто Объявления</a> </b></p>
            </td>
            <td width="50" height="22" bgcolor="#000000" align="right"> 
              <%while (String(msgcount).length<4) {msgcount="0"+String(msgcount)}%>
              <p><font size="1"><b> <font size="2" color="#FFFFFF"><%=msgcount%></font></b></font></p>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="39"> 
              <p> 
            </td>
            <td width="50" height="39" valign="top"> 
              <p align="center"><font size="1">Добавлено</font></p>
            </td>
          </tr>
        </table>
        <%
// маркек признака новостей
isnews=1
// если необходимо вывести рубрики не новостей то установить в ноль
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date>='TODAY'-"+per+" order by public_date desc, id desc"
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
	recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	while (String(kvopub).length<4) {kvopub="0"+String(kvopub)}
	Records.MoveNext()
%>
        <table width="95%" border="0" cellspacing="0">
          <tr bgcolor="#990000"> 
            <td height="4"></td>
            <td width="50" height="4" align="right"></td>
          </tr>
          <tr> 
            <td height="23" bgcolor="#E8ECEC"> 
              <p><b><img src="Img/arrow1.gif" width="12" height="11" align="absmiddle"> 
                <a href="<%=url%>"><%=hname%></a></b></p>
            </td>
            <td width="50" height="23" align="right" bgcolor="#000000"> 
              <p><font size="2" color="#FFFFFF"><b><%=kvopub%></b></font></p>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF" valign="top"> 
            <td height="39"> 
              <p><font size="1" color="#0000FF"><a href="<%=nadr%>"><%=name%></a></font></p>
            </td>
            <td width="50" height="39" align="right"> 
              <p><font size="1">Добавлено<br>
                <b><%=ndat%></b></font></p>
            </td>
          </tr>
        </table>
        <%
} Records.Close()
delete recs
%>
        <%
// маркек признака новостей
isnews=0
// если необходимо вывести рубрики не новостей то установить в ноль
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
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
	recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	while (String(kvopub).length<4) {kvopub="0"+String(kvopub)}
	Records.MoveNext()
%>
        <TABLE WIDTH="95%" BORDER="0" CELLSPACING="0">
          <TR BGCOLOR="#990000"> 
            <TD HEIGHT="4"></TD>
            <TD WIDTH="50" HEIGHT="4" ALIGN="right"></TD>
          </TR>
          <TR> 
            <TD HEIGHT="23" BGCOLOR="#E8ECEC"> 
              <P><B><IMG SRC="Img/arrow1.gif" WIDTH="12" HEIGHT="11" ALIGN="absmiddle"> 
                <A HREF="<%=url%>"><%=hname%></A></B></P>
            </TD>
            <TD WIDTH="50" HEIGHT="23" ALIGN="right" BGCOLOR="#000000"> 
              <P><FONT SIZE="2" COLOR="#FFFFFF"><B><%=kvopub%></B></FONT></P>
            </TD>
          </TR>
          <TR BGCOLOR="#FFFFFF" VALIGN="top"> 
            <TD HEIGHT="39"> 
              <P><FONT SIZE="1" COLOR="#0000FF"><A HREF="<%=nadr%>"><%=name%></A></FONT></P>
            </TD>
            <TD WIDTH="50" HEIGHT="39" ALIGN="right"> 
              <P><FONT SIZE="1">Добавлено<BR>
                <B><%=ndat%></B></FONT></P>
            </TD>
          </TR>
        </TABLE>
        <%
} Records.Close()
delete recs
%>
        <table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor="#990000"> 
            <td height="4"></td>
            <td width="62" height="4" align="right"></td>
          </tr>
          <tr> 
            <td height="22" bgcolor="#E8ECEC"> 
              <p><b><img src="Img/arrow1.gif" width="12" height="11" align="absmiddle"> 
                <a href="http://www.auction.72rus.ru/">Сибирский Аукцион</a></b></p>
            </td>
            <td width="62" height="22" bgcolor="#000000" align="right"> 
              <p> 
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="39"> 
              <p><a href="http://www.auction.72rus.ru/auction.asp?subj_id=5"><font size="1"><b>Раздел 
                &quot;Автомобили и запчасти&quot;</b></font></a><b><font size="1"> 
                - проведение индивидуальных торгов. </font></b> 
            </td>
            <td width="62" height="39" valign="top"> 
              <p align="center">&nbsp;</p>
            </td>
          </tr>
        </table>
        <table width="100%" border="1" cellspacing="2" cellpadding="0" bgcolor="#0033FF" bordercolor="#FFFFFF">
          <tr> 
            <td> 
              <%
// В переменной bk содержится код блока новостей
var bk=26
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
              <h1><font color="#FFFFFF" size="2"><%=blokname%></font></h1>
            </td>
          </tr>
        </table>
        <%
// В переменной bk содержится код блока новостей
var bk=26
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
	hadr=TextFormData(recs("URL").Value,"vvr_list.asp.asp")
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
        <font face="Arial, Helvetica, sans-serif"><b> </b></font> 
        <table width="100%" border="1" cellspacing="3" cellpadding="0" bgcolor="#FFFFFF" bordercolor="#666666">
          <tr> 
            <td bgcolor="#FFFFFF" bordercolor="#333333"><%=news%> </td>
          </tr>
        </table>
        <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
        <p align="left">&nbsp; </p>
        <%if (tpm<3) {%>
        <P><FONT SIZE="1"><A HREF="addnewsheading.asp?hid=0">Добавить рубрику 
          новостей</A></FONT></P>
        <%}%>
      </div>
    </td>
    <td valign="top" align="center" bgcolor="#666666"> 
      <table width="98%" border="1" cellspacing="2" cellpadding="0" bgcolor="#0033FF" bordercolor="#FFFFFF">
        <tr> 
          <td> 
            <p><font color="#FFFFFF"><b>Авто Новости - События - Новинки</b></font></p>
          </td>
        </tr>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFFFFF" valign="top" align="left"> 
          <td height="147" colspan="2" bgcolor="#E6F2EC"> 
            <%
// В переменной bk содержится код блока новостей
var bk=18
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2, block_news t3 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+"and t2.block_news_id=t3.id and t3.smi_id="+smi_id+" order by t2.posit"
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
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a>  "+path
	hdd=recs("HI_ID").Value
	recs.Close()
	}

%>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#990000"> 
                <td height="4" width="497"></td>
                <td width="144" height="4" align="right"></td>
              </tr>
              <tr> 
                <td height="22" bgcolor="#E8ECEC" align="left" width="497"> 
                  <p><b><font size="2"> </font><img src="Img/arrow1.gif" width="12" height="11" align="absmiddle"> 
                    <font class=globalnav><b><%=path%></b></font></b></p>
                </td>
                <td width="144" height="22" bgcolor="#000033" align="right"> 
                  <p><font size="1"><b> <font size="2" color="#FFFFFF">AUTO.72RUS.RU</font></b></font></p>
                </td>
              </tr>
            </table>
            <table width="100%" border="1" bordercolor="#FFFFFF" cellspacing="0">
              <tr bgcolor="#666666"> 
                <td bgcolor="#FFFFFF" height="23"> 
                  <h1><font face="Arial, Helvetica, sans-serif" color="#CC3333"><b> 
                    </b></font><font face="Arial, Helvetica, sans-serif"><b><%=pdat%> &nbsp;<%=pname%> </b></font></h1>
                </td>
              </tr>
              <tr bordercolor="#CCCCCC" bgcolor="#FFFFFF"> 
                <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
                  <p align="center"><font face="Arial, Helvetica, sans-serif"> 
                    <%if (imgLname != "") {%>
                    <a href="newshow.asp?pid=<%=pid%>"><img src="<%=imgLname%>" border="1" ></a> 
                    <%}else{%>
                    &nbsp; 
                    <%}%>
                    </font></p>
                  <p align="left"><font size="2" face="Arial, Helvetica, sans-serif"><%=digest%></font> <font face="Arial, Helvetica, sans-serif" color="#990000" size="1"><a href="newshow.asp?pid=<%=pid%>">Подробнее...</a> 
                    </font> </p>
                </td>
              </tr>
              <tr bordercolor="#CCCCCC" bgcolor="#FFFFFF"> </tr>
            </table>
            <font face="Arial, Helvetica, sans-serif" color="#990000" size="1"> 
            <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
            </font></td>
        </tr>
        <tr bgcolor="#CCCCCC"> </tr>
      </table>
      <table width="98%" border="1" cellspacing="2" cellpadding="0" bgcolor="#0033FF" bordercolor="#FFFFFF">
        <tr> 
          <td> 
            <p><font color="#FFFFFF"><b>Авто Новости - События - Новинки</b></font></p>
          </td>
        </tr>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFFFFF" valign="top" align="left"> 
          <td height="39" colspan="2"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#990000"> 
                <td height="4" width="497"></td>
                <td width="144" height="4" align="right"></td>
              </tr>
              <tr> 
                <td height="22" bgcolor="#E8ECEC" align="left" width="497"> 
                  <p><b><font size="2"> </font><img src="Img/arrow1.gif" width="12" height="11" align="absmiddle"> 
                    <font class=globalnav></font></b></p>
                </td>
                <td width="144" height="22" bgcolor="#000033" align="right"> 
                  <p><font size="1"><b> <font size="2" color="#FFFFFF">AUTO.72RUS.RU</font></b></font></p>
                </td>
              </tr>
            </table>
            <table width="100%" border="1" bordercolor="#FFFFFF" cellspacing="0">
              <tr bordercolor="#CCCCCC" bgcolor="#FFFFFF"> 
                <%
// В переменной bk содержится код блока новостей
var bk=19
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgname=+pid+".gif"
    if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
	if (imgname=="") {
		imgname=+pid+".jpg"
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
	}
%>
                <td bgcolor="#E6F2EC" bordercolor="#FFFFFF" valign="top" width="50%"><font size="1">- 
                  <a href="newshow.asp?pid=<%=pid%>"><font face="Arial, Helvetica, sans-serif"><%=pname%></font></a><font face="Arial, Helvetica, sans-serif"> 
                  (<%=pdat%></font></font><font size="1" face="Arial, Helvetica, sans-serif">)<br>
                  <%=digest%>...</font></td>
                <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
              </tr>
            </table>
            <h1> 
              <%
// В переменной bk содержится код блока новостей
var bk=20
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2, block_news t3 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+"and t2.block_news_id=t3.id and t3.smi_id="+smi_id+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
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
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
	}
%>
            </h1>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
              <tr> 
                <td bgcolor="#990000" height="4"></td>
              </tr>
              <tr> 
                <td bgcolor="#E8ECEC" height="21"> 
                  <p><img src="Img/arrow1.gif" width="12" height="11" align="absmiddle"> 
                    <%=path%></p>
                </td>
              </tr>
            </table>
            <table width="100%" border="1" bordercolor="#FFFFFF" cellspacing="0">
              <tr bordercolor="#CCCCCC" bgcolor="#FFFFFF"> 
                <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
                  <h1><%=pdat%> &nbsp;<%=pname%></h1>
                  <p><font face="Arial, Helvetica, sans-serif"> 
                    <%if (imgname != "") {%>
                    <img src="<%=imgname%>" align="left" > 
                    <%}else{%>
                    &nbsp; 
                    <%}%>
                    </font><font size="2" face="Arial, Helvetica, sans-serif"><%=digest%></font> <font face="Arial, Helvetica, sans-serif" color="#990000" size="1"><a href="newshow.asp?pid=<%=pid%>">Подробнее...</a> 
                    </font> </p>
                  <p> 
                  <p> 
                </td>
              </tr>
            </table>
            <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
        <tr bgcolor="#CCCCCC"> 
          <td height="5" colspan="2" align="center"> </td>
        </tr>
      </table>
    </td>
    <td align="left" width="150" valign="top" bgcolor="#666666"> 
      <table width="100%" border="1" cellspacing="2" cellpadding="0" bgcolor="#0033FF" bordercolor="#FFFFFF">
        <tr> 
          <td> 
            <p><font color="#FFFFFF"><b>Информация</b></font></p>
          </td>
        </tr>
      </table>
      <b><font size="2"> </font></b> 
      <table width="150" border="1" cellspacing="3" cellpadding="0" bordercolor="#FFFFFF" bgcolor="#0000FF">
        <tr bordercolor="#FFFFFF"> 
          <td colspan="2" bgcolor="#0033FF" valign="top" height="37"> 
            <p align="center"><font ><img src="gas.gif" width="32" height="29"></font></p>
          </td>
          <td bgcolor="#0033FF" height="37" width="90"> 
            <div align="center"> 
              <p><font color="#FFFFFF"><b>ЦЕНЫ<br>
                НА БЕНЗИН<br>
                В ТЮМЕНИ</b></font></p>
            </div>
          </td>
        </tr>
      </table>
      <table width="150" border="1" cellspacing="3" cellpadding="0" bgcolor="#FFFFFF" bordercolor="#666666">
        <tr> 
          <td bgcolor="#FFFFFF" bordercolor="#333333"> 
            <%
			Records.Source="Select valname, rate from valuta where valuta_type_id=5 and enterprise_id=1"
			Records.Open()
			while (!Records.EOF) {
				bnm=Records("VALNAME").Value
				bcost=Records("RATE").Value
			%>
            <p align="left"><b><img src="Img/arrow1.gif" width="12" height="11" align="absmiddle"> 
              <%=bnm%> - <%=bcost%> руб.</b></p>
            <p align="right"> 
              <%
				Records.MoveNext()
			}
			Records.Close()
			%>
              <b><font size="1"><a href="http://www.russianoil.ru/" target="_blank">Русская 
              Нефть</a></font></b> <br>
            </p>
          </td>
        </tr>
      </table>
      <table width="100%" border="1" cellspacing="2" cellpadding="0" bgcolor="#0033FF" bordercolor="#FFFFFF">
        <tr> 
          <td> 
            <%
// В переменной bk содержится код блока новостей
var bk=21
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
            <h1><font color="#FFFFFF" size="2"><%=blokname%></font></h1>
          </td>
        </tr>
      </table>
      <font face="Arial, Helvetica, sans-serif"><b> 
      <%
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2, block_news t3 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+"and t2.block_news_id=t3.id and t3.smi_id="+smi_id+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
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
%>
      </b></font> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
        <tr> 
          <td bgcolor="#990000" height="4"></td>
        </tr>
        <tr> 
          <td bgcolor="#E8ECEC" height="23"> 
            <p><img src="Img/arrow1.gif" width="12" height="11" align="absmiddle"><font face="Arial, Helvetica, sans-serif"><b> 
              </b></font> <a href="newshow.asp?pid=<%=pid%>"> 
              <%=pname%> </a></p>
          </td>
        </tr>
      </table>
      <table width="100%" border="1" bordercolor="#FFFFFF" cellspacing="0">
        <tr bordercolor="#CCCCCC" bgcolor="#FFFFFF"> 
          <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <div align="center"> 
              <%if (imgname != "") {%>
              <img src="<%=imgname%>" align="absmiddle" > 
              <%}else{%>
              &nbsp; 
              <%}%>
            </div>
          </td>
        </tr>
      </table>
      <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
      <table width="100%" border="1" cellspacing="2" cellpadding="0" bgcolor="#0033FF" bordercolor="#FFFFFF">
        <tr> 
          <td height="15"> 
            <p align="center"><font color="#FFFFFF"><b>Реклама</b></font></p>
          </td>
        </tr>
      </table>
      <table width="120" border="1" cellspacing="3" cellpadding="0" bgcolor="#FFFFFF" bordercolor="#666666" align="center">
        <tr> 
          <td bgcolor="#FFFFFF" height="37" bordercolor="#333333"> 
            <table width="120" border="0" cellspacing="0" height="60" align="center">
              <tr> 
                <%
// В переменной bk содержится код блока новостей
var bk=69
// Не забывать его менять!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	ishtml=TextFormData(Records("ISHTML").Value,"0")

news=""
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
%>
                <td align="CENTER"><%=news%></td>
                <%
Records.MoveNext()
} 
Records.Close()
%>
              </tr>
            </table>
            
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<div align="center"> 
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" bgcolor="#333333">
    <tr> 
      <td width="150" valign="middle" height="3"> </td>
      <td height="3"> </td>
      <td width="468" height="3"> </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
    <tr> 
      <td bgcolor="#CCCCCC" valign="middle" height="20"> 
        <h1><b> <font size="2"><a href="http://www.72rus.ru/">72RUS</a> - АВТО 
          ТЮМЕНЬ</font></b></h1>
      </td>
      <td width="468" height="20" align="right" bgcolor="#CCCCCC"> 
        <table width="468" border="1" cellspacing="2" cellpadding="0" bordercolor="#CCCCCC">
          <tr bordercolor="#000000" bgcolor="#FFFFFF"> 
            <td width="82" height="20"> 
              <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="index.asp">АВТО 
                ТЮМЕНЬ</a></b></font></div>
            </td>
            <td width="82" height="20"> 
              <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="auto_subj.asp">Объявления</a></b></font></div>
            </td>
            <td width="82" height="20"> 
              <div align="center"><a href="auto_penalty.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>Штрафы</b></font></a></div>
            </td>
            <td width="82" height="20"> 
              <div align="center"><a href="auto_regions.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>Регионы</b></font></a></div>
            </td>
            <td width="82" height="20"> 
              <div align="center"><a href="auto_dov.asp"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>Доверенность</b></font></a></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" bgcolor="#333333">
    <tr> 
      <td width="150" valign="middle" height="4"> </td>
      <td height="4"> </td>
      <td width="468" height="4"> </td>
    </tr>
  </table>
  <table width="100%" cellspacing="0" cellpadding="0" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td valign="bottom" height="60" width="150"> 
        <P>&nbsp;</P>
      </td>
      <td align="right" height="60"> 
        <table width="468" border="0" cellspacing="0" height="60">
          <tr> 
            <%
// В переменной bk содержится код блока новостей
var bk=66
// Не забывать его менять!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	ishtml=TextFormData(Records("ISHTML").Value,"0")

news=""
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
%>
            <td align="CENTER"><%=news%></td>
            <%
Records.MoveNext()
} 
Records.Close()
%>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#0033FF" height="13"> 
        <p align="left"><a href="pubarea.asp">Редактор</a> <a href="area.asp">Администратор<b> 
          <font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"> 
          </font></b></a><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
          портал 72RUS >> AUTO Тюмень - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
          Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
          Русинтел</font></a> <font color="#FFFFFF">&copy; 2002</font></font></b></p>
      </td>
    </tr>
  </table>
</div>
<p align="center"> © 2002 <a href="http://www.rusintel.ru">Rusintel Company</a> 
<h2 align="center"> 
  <!-- HotLog -->
  <script language="javascript">
hotlog_js="1.0";
hotlog_r=""+Math.random()+"&s=58972&im=5&r="+escape(document.referrer)+"&pg="+
escape(window.location.href);
document.cookie="hotlog=1; path=/"; hotlog_r+="&c="+(document.cookie?"Y":"N");
</script>
  <script language="javascript1.1">
hotlog_js="1.1";hotlog_r+="&j="+(navigator.javaEnabled()?"Y":"N")</script>
  <script language="javascript1.2">
hotlog_js="1.2";
hotlog_r+="&wh="+screen.width+'x'+screen.height+"&px="+
(((navigator.appName.substring(0,3)=="Mic"))?
screen.colorDepth:screen.pixelDepth)</script>
  <script language="javascript1.3">hotlog_js="1.3"</script>
  <script language="javascript">hotlog_r+="&js="+hotlog_js;
document.write("<a href='http://click.hotlog.ru/?58972' target='_top'><img "+
" src='http://hit4.hotlog.ru/cgi-bin/hotlog/count?"+
hotlog_r+"&' border=0 width=88 height=31 alt=HotLog></a>")</script>
  <noscript><a href=http://click.hotlog.ru/?58972 target=_top><img
src="http://hit4.hotlog.ru/cgi-bin/hotlog/count?s=58972&im=5" border=0 
width="88" height="31" alt="HotLog"></a></noscript> 
  <!-- /HotLog -->
  <a href="http://www.ladaonline.ru/"><img src="http://www.ladaonline.ru/pics/ban/ladaonline88_2.gif" border="0" width="88" height="31" alt="Информационный проект LADAONLINE"></a> 
  <a href="http://www.72rus.ru/"> <img src="http://www.72rus.ru/72rus.gif" width="88" height="31" alt="72RUS - Тюменский Регион" border="0"></a> 
  <!--Begin of HMAO RATINGS-->
  <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"><img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="88" height="31"></a> 
  <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
  <!--End of HMAO RATINGS-->
  <!--begin of Top100 logo-->
  <a href="http://top100.rambler.ru/top100/"> <img src="http://top100-images.rambler.ru/top100/w2.gif" alt="Rambler's Top100" width=88 height=31 border=0></a> 
  <!--end of Top100 logo -->
</h2>
</body>
</html>
