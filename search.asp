<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var tps=parseInt(Request("tps"))
var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
if (isNaN(tps)) {tps=1}
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}
var lp=10 // Длина строницы
var sstr=""
var ssql=""
var qsql=""
var ii=0
var name=""
var id=0
var coment=""
var kvo=0
var tkvo=0
var dat=""
var url=""
var ll=0

if (sch!="") {
	// Какой-то запрос
	sstr=sch
	while (sch.indexOf(".")>=0) {sch=sch.replace(".","")}
	while (sch.indexOf(",")>=0) {sch=sch.replace(","," ")}
	while (sch.indexOf(" - ")>=0) {sch=sch.replace(" - "," ")}
	while (sch.indexOf(" -")>=0) {sch=sch.replace(" -"," ")}
	while (sch.indexOf("- ")>=0) {sch=sch.replace("- "," ")}
	while (sch.indexOf(";")>=0) {sch=sch.replace(";"," ")}
	while (sch.indexOf(":")>=0) {sch=sch.replace(":"," ")}
	while (sch.indexOf("\"")>=0) {sch=sch.replace("\"","")}
	while (sch.indexOf("'")>=0) {sch=sch.replace("'","")}
	while (sch.indexOf("(")>=0) {sch=sch.replace("(","")}
	while (sch.indexOf(")")>=0) {sch=sch.replace(")","")}
	while (sch.indexOf("<")>=0) {sch=sch.replace("<"," ")}
	while (sch.indexOf(">")>=0) {sch=sch.replace(">"," ")}
	while (sch.indexOf("?")>=0) {sch=sch.replace("?"," ")}
	while (sch.indexOf("=")>=0) {sch=sch.replace("="," ")}
	while (sch.indexOf(" % ")>=0) {sch=sch.replace(" % "," ")}
	while (sch.indexOf(" _ ")>=0) {sch=sch.replace(" _ "," ")}
	while (sch.indexOf("+")>=0) {sch=sch.replace("+"," ")}
	while (sch.indexOf("  ")>=0) {sch=sch.replace("  "," ")}
	ll=sch.length
	if (ll<3) {sch=""}
	if (sch.indexOf(" ")==0) {sch=sch.substring(1,ll)}
	ll=sch.length
	if (ll<3) {sch=""}
	if (sch.indexOf(" ")==ll-1) {sch=sch.substring(0,ll-1)}
	if (wrds==0) {sch=""}
	//
	//
	//
	if (sch!="") {
		ssql=sch
		while (ssql.indexOf(" ")>=0) {ssql=ssql.replace(" ","+")}
		if (sens==1) {
			if (wrds==1) { // С учетом регистра букв ВСЕ слова
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+","%' and t1.name like  '%")}
				ssql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and t1.name like '%"+ssql+"%' "
				qsql=sch
				while (qsql.indexOf(" ")>=0) {qsql=qsql.replace(" ","+")}
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+","%' and t1.digest like  '%")}
				qsql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and t1.digest like '%"+qsql+"%'"
			}
			if (wrds==2) { // С учетом регистра букв Хотябы одно слово
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+","%' or t1.name like  '%")}
				ssql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and (t1.name like '%"+ssql+"%') "
				qsql=sch
				while (qsql.indexOf(" ")>=0) {qsql=qsql.replace(" ","+")}
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+","%' or t1.digest like  '%")}
				qsql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and (t1.digest like '%"+qsql+"%')"
			}
			if (wrds==3) { // С учетом регистра букв Фраза целиком
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+"," ")}
				ssql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and t1.name like '%"+ssql+"%' "
				qsql=sch
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+"," ")}
				qsql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and t1.digest like '%"+qsql+"%'"
			}
		} else {
			if (wrds==1) { // Без учета регистра букв ВСЕ слова
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+","%') and UpCase(t1.name) like  UpCase('%")}
				ssql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and UpCase(t1.name) like UpCase('%"+ssql+"%') "
				qsql=sch
				while (qsql.indexOf(" ")>=0) {qsql=qsql.replace(" ","+")}
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+","%') and UpCase(t1.digest) like  UpCase('%")}
				qsql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and UpCase(t1.digest) like UpCase('%"+qsql+"%')"
			}
			if (wrds==2) { // Без учета регистра букв Хотябы одно слово
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+","%') or UpCase(t1.name) like  UpCase('%")}
				ssql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and (UpCase(t1.name) like UpCase('%"+ssql+"%')) "
				qsql=sch
				while (qsql.indexOf(" ")>=0) {qsql=qsql.replace(" ","+")}
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+","%') or UpCase(t1.digest) like  UpCase('%")}
				qsql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and (UpCase(t1.digest) like UpCase('%"+qsql+"%'))"
			}
			if (wrds==3) { // Без учета регистра букв Фраза целиком
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+"," ")}
				ssql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and UpCase(t1.name) like UpCase('%"+ssql+"%') "
				qsql=sch
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+"," ")}
				qsql="Select t1.* from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and UpCase(t1.digest) like UpCase('%"+qsql+"%')"
			}
		}
		ssql=ssql+" UNION "+qsql+" order by 2"
	}
	
}

%>

<html>
<head>
<title>Результаты поиска: "<%=sch%>"  <<  Белый Парус</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="1" topmargin="3">
<table border="0" cellspacing="0" cellpadding="0" bordercolor="#666699" width="780">
  <tr> 
    <td background="images/pl_bg.gif" align="right" bgcolor="#666699"> 
      <p  class="secondarystories"><a href="gbook.asp"><font color="#FFFFFF">Гостевая 
        книга</font></a> * <a href="message.asp"><font color="#FFFFFF">Ваш вопрос</font></a> 
        * <a href="javascript:window.external.AddFavorite(parent.location,document.title)"> 
        <font color="#FFFFFF">Добавить в избранное</font></a>:</p>
    </td>
  </tr>
</table>
<table width="780" border="1" cellspacing="0" cellpadding="0" bordercolor="#666699">
  <tr bgcolor="#E2E2F1" bordercolor="#E2E2F1"> 
    <td height="16" width="400" bgcolor="#E2E2F1"> 
      <h1>&nbsp;&nbsp;&nbsp;<a href="/">Закон ON line</a> </h1>
    </td>
    <form name="form1" method="post" action="search.asp">
      <td height="16" align="right"> 
        <p><b><font color="#666699">Поиск</font></b><font color="#FFFFFF"><b> 
          &nbsp;</b></font> 
          <input type="text" name="sch2" size="40" value="" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 114px" >
          <select name="select" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 100px">
            <option value="1" <%=wrds==1?"selected":""%>>Все слова</option>
            <option value="2" <%=wrds==2?"selected":""%>>Одно из слов</option>
            <option value="3" <%=wrds==3?"selected":""%>>Фраза целиком</option>
          </select>
          <input type="submit" name="Findit2" style="FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 50px; HEIGHT: 20px" value="Найти">
          &nbsp;&nbsp; </p>
      </td>
    </form>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" height="12">
  <tr> 
    <td width="1"></td>
    <td valign="top" width="778"></td>
    <td bgcolor="#123D87" width="1"></td>
  </tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="1"></td>
    <td width="12" bgcolor="#FFFFFF"></td>
    <td valign="top" bgcolor="#FFFFFF" width="754"> 
      <table border="0" cellspacing="0" cellpadding="0" bordercolor="#666699" width="754">
        <tr> 
          <td background="images/pl_bg.gif" width="250" align="center"><span class="secondarystories"><a href="area.asp"><font color="#FFFFFF">Кабинет</font></a> 
            &nbsp; <a href="/"><font color="#FFFFFF">На главную</font></a></span></td>
          <td width="20"><img src="images/pl_ang.gif" width="20" height="16"></td>
          <td background="images/pl_bg_next.gif" width="484"></td>
        </tr>
      </table>
    </td>
    <td valign="bottom" width="12" bgcolor="#FFFFFF"></td>
    <td width="1" bgcolor="#123D87"></td>
  </tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0" height="286">
  <tr> 
    <td width="1"></td>
    <td width="12">&nbsp;</td>
    <td width="1" valign="top" bgcolor="#666699"></td>
    <td valign="top" bgcolor="#FFFFFF" align="center"> 
      <p>Результаты поиска: " <%=sch%> " </p>
      <form name="form1" method="post" action="search.asp">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#E2E2F1"> 
              <p> Новый поиск 
                <input type="text" name="sch" size="40" value="<%=sch%>" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 220px" >
                <input type="submit" name="Findit" value="Найти" style="FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 50px; HEIGHT: 20px">
                <select name="wrds" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 100px">
                  <option value="1" <%=wrds==1?"selected":""%>>Все слова</option>
                  <option value="2" <%=wrds==2?"selected":""%>>Одно из слов</option>
                  <option value="3" <%=wrds==3?"selected":""%>>Фраза целиком</option>
                </select>
                <b><font size="2"> 
                <input type="checkbox" name="sensation" value="1" <%=sens==1?"checked":""%>>
                </font></b><font size="2">Учитывать регистр</font></p>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#123D87" height="1"></td>
          </tr>
        </table>
      </form>
      <p><%=""%></p>
      <%
if (ssql!="") {
//Records.CursorType=3
Records.Source=ssql
Records.Open()
if (!Records.EOF) {
kvo=Records.RecordCount
if ((pg+1)*lp > kvo) {tkvo=kvo} else {tkvo=(pg+1)*lp}
%>
      <p align="left"><font color="#003399">Найдено публикаций: <%=kvo%></font></p>
      <%
	ii=0
	while ((!Records.EOF) && (ii<tkvo)) {
		ii+=1
		id=String(Records("ID").Value)
		name=String(Records("NAME").Value)
		coment=TextFormData(Records("DIGEST").Value,"")
		dat=Records("PUBLIC_DATE").Value
		url=TextFormData(Records("URL").Value,"newshow.asp")
		url+="?pid="+id
		if (ii>=(pg*lp+1)) {
%>
      <table width="100%" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
        <tr bgcolor="#F2F2F2"> 
          <td width="6%"> 
            <div align="center"> 
              <p><b><font color="#000099"><%=ii%>.</font></b></p>
            </div>
          </td>
          <td bgcolor="#F2F2F2"> 
            <div align="left"> 
              <p><b></b> <b><a href="<%=url%>"><%=name%></a></b> <font color="#999999">[ 
                <%=dat%> ]</font></p>
            </div>
          </td>
        </tr>
        <tr bgcolor="#F9F9F9"> 
          <td colspan="2"> 
            <p><font size="2"><%=coment%></font></p>
          </td>
        </tr>
        <tr> 
          <td colspan="2" height="6"></td>
        </tr>
      </table>
      <%
		}
		Records.MoveNext()
	}
}
Records.Close()
%>
      <hr noshade size="1">
      <p>Страницы: 
        <%
for ( ii=1; ii<(kvo/lp + 1) ; ii++) {
%>
        <% if (ii==(pg+1)) { %>
        <%=ii%> | 
        <%} else {%>
        <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=<%=tps%>&pg=<%=ii-1%>"><%=ii%></a> 
        | 
        <%}%>
        <%
}
%>
      </p>
      <%
}
%>
    </td>
    <td width="1" valign="top" bgcolor="#666699"></td>
    <td valign="top" width="12">&nbsp;</td>
    <td width="1" bgcolor="#123D87"></td>
  </tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" height="1">
  <tr> 
    <td width="1" bgcolor="#666699"></td>
    <td width="12"></td>
    <td width="1" valign="top" bgcolor="#E2E2F1"></td>
    <td bgcolor="#666699" valign="top" width="752"> </td>
    <td width="1" valign="top" bgcolor="#666699"></td>
    <td valign="top" width="12"></td>
    <td width="1" bgcolor="#666699"></td>
  </tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" height="12">
  <tr> 
    <td width="1"></td>
    <td width="12"><img src="images/dot_tr.gif" width="12" height="6"></td>
    <td valign="top"> 
      <table border="0" cellspacing="0" cellpadding="0" bordercolor="#666699" width="752">
        <tr> 
          <td width="662"></td>
          <td width="90">&nbsp;</td>
        </tr>
      </table>
    </td>
    <td valign="top" width="12"></td>
    <td width="1" bgcolor="#666699"></td>
  </tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td width="1"></td>
    <td width="12">
    </td>
    <td width="468"> 
      <table width="468" border="0" cellspacing="0" cellpadding="0" height="60">
        <tr> 
          <td bgcolor="#F9F9F9" bordercolor="#666699" valign="middle" align="center"> 
            <script language="javascript" src="http://www.72rus.ru/banshow.asp?rid=12&u=http://zakon.72rus.ru/"></script></td>
        </tr>
      </table>
    </td>
    <td valign="top" width="12"></td>
    <td valign="top" width="274"> 
      <table border="1" cellspacing="0" cellpadding="0" bordercolor="#666699" height="60" width="100%">
        <tr> 
          <td bgcolor="#F9F9F9" bordercolor="#F9F9F9" align="center"> 
            <script language="javascript">
hotlog_js="1.0";
hotlog_r=""+Math.random()+"&s=46088&im=105&r="+escape(document.referrer)+"&pg="+
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
document.write("<a href='http://click.hotlog.ru/?46088' target='_top'><img "+
" src='http://hit3.hotlog.ru/cgi-bin/hotlog/count?"+
hotlog_r+"&' border=0 width=88 height=31 alt=HotLog></a>")</script>
            <noscript><a href=http://click.hotlog.ru/?46088 target=_top><img
src="http://hit3.hotlog.ru/cgi-bin/hotlog/count?s=46088&im=105" border=0 
width="88" height="31" alt="HotLog"></a></noscript> 
            <!-- /HotLog -->
            <script language="JavaScript">document.write('<a href="http://www.rax.ru/click" target=_blank><img src="http://counter.yadro.ru/hit?t16.1;r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" border=0 width=88 height=31 alt="rax.ru: показано число хитов за 24 часа, посетителей за 24 часа и за сегодн\я"></a>')</script>
          </td>
        </tr>
      </table>
    </td>
    <td valign="top" width="12"></td>
    <td width="1" bgcolor="#666699"></td>
  </tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" height="12">
  <tr> 
    <td width="1"></td>
    <td width="12"><img src="images/dot_tr.gif" width="12" height="6"></td>
    <td width="468" valign="top"></td>
    <td valign="top" width="12"></td>
    <td valign="top" width="274"></td>
    <td valign="top" width="12"></td>
    <td width="1" bgcolor="#666699"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" height="12" bgcolor="#FFFFFF">
  <tr> 
    <td width="1"></td>
    <td> 
      <p align="center"><a href="/">На главную</a> <img src="images/dot.gif" width="12" height="9" align="absmiddle"> 
        <a href="message.asp">Обратная связь</a> <img src="images/dot.gif" width="12" height="9" align="absmiddle"> 
        <a href="gbook.asp">Гостевая книга</a></p>
    </td>
    <td bgcolor="#666699" width="1"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" height="12" bgcolor="#E1F4FF">
  <tr> 
    <td width="1"></td>
    <td bgcolor="#FFFFFF" align="center"> 
      <p><font size="1"><font size="1"><font color="#FFFFFF"> </font></font> <a href="area.asp">&copy;</a> 
        2003 Разработка <a href="http://www.rusintel.ru"> Русинтел</a></font></p>
    </td>
    <td bgcolor="#666699" width="1"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" bordercolor="#666699" width="780">
  <tr> 
    <td width="690"></td>
    <td width="90"><img src="images/ugolok.gif" width="90" height="9"></td>
  </tr>
</table>
</body>
</html>
