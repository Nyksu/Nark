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

path="<a href=\"index.asp\">"+sminame+"</a> > "+path
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
<title><%=tit%> > <%=pname%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body bgcolor="#FCEDBE" text="#000000" topmargin="0" leftmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr bgcolor="#135086"> 
    <td valign="middle" width="500"> 
      <p><b><font color="#DBEAF2">&quot;<%=sminame%>&quot;</font></b></p>
    </td>
    <td bgcolor="#135086" align="right"> <font></font> 
      <p><font color="#DBEAF2"><b>Независимая оценка, консалтинг </b></font></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr bgcolor="#135086"> 
    <td valign="middle" width="500" bgcolor="#DBEAF2"> 
      <p>&nbsp;<%=path%> </p>
    </td>
    <td bgcolor="#DBEAF2" align="right"> 
      <p><b><font color="#135086"><%=nm%></font></b></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="1" bgcolor="#135086">
  <tr> 
    <td></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="12">
  <tr> 
    <td></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
  <tr bgcolor="#FCEDBE"> 
    <td valign="top" align="right" width="12">&nbsp;</td>
    <td valign="top" align="right" width="202"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="D92B2B">
        <tr align="center"> 
          <td colspan="5"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#DBEAF2">
              <tr> 
                <td align="center" class="dir_title" valign="middle" bgcolor="#135086" width="23">&nbsp;</td>
                <td align="left" class="dir_title" valign="middle" bgcolor="#135086"> 
                  <p><b><font color="#DBEAF2">Разделы</font></b></p>
                </td>
                <td width="21" bgcolor="#135086">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table width="100%" border="1" cellspacing="0" cellpadding="5" class="base_text" bordercolor="#135086">
        <tr> 
          <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <li> <a class="globalnav" href="index.asp"> 
              <p><b>В начало</b></p>
              </a> 
              <%
isnews=1
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
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
            <li><a  class=globalnav href="<%=url%>"> 
              <p><b><%=hname%></b></p>
              </a> 
              <%
} 
Records.Close()
%>
              <%
isnews=0
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
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
            <li><a  class=globalnav href="<%=url%>"> 
              <p><b><%=hname%></b></p>
              </a> 
              <%
} 
Records.Close()
%>
            <li> <a class=globalnav href="message.asp"> 
              <p><b>E-mail</b> 
              </a> 
            <li> <a class=globalnav href="org.asp"> 
              <p><b>Реквизиты</b> 
              </a> 
              <%if (usok) {%>
            <li> 
              <p><b><font face="Arial, Helvetica, sans-serif"><a href="addnewsheading.asp?hid=<%=hid%>"><font color="#006600">Добавить 
                раздел сайта</font></a></font><font face="Arial, Helvetica, sans-serif" size="1"> 
                <%}%>
                </font></b> </p>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="12">
        <tr> 
          <td></td>
        </tr>
      </table>
      <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr bgcolor="#135086"> 
          <td width="14">&nbsp;</td>
          <td nowrap valign="middle" align="left"> 
            <p><b> 
              <%
// В переменной bk содержится код блока новостей
var bk=71
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
              <font color="#DBEAF2"><%=blokname%> </font></b></p>
          </td>
          <td width="14">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="1" cellspacing="0" cellpadding="5" class="base_text" bordercolor="#135086">
        <tr> 
          <td bgcolor="#FFFFFF" bordercolor="#FFFFFF" valign="top"> 
            <%
var pidd=0
var  pname2=""
// Не забывать его менять!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	pidd=String(Records("ID").Value)
	pname2=String(Records("NAME").Value)
%>
            <li> 
              <p> &nbsp;<a href="newshow.asp?pid=<%=pidd%>"><%=pname2%></a></p>
              <%
	Records.MoveNext()
} 
Records.Close()
// Блок  71 закончился
%>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="12">
        <tr> 
          <td></td>
        </tr>
      </table>
    </td>
    <td valign="top" align="left" width="12">&nbsp;</td>
    <td valign="top" align="left" bgcolor="#FCEDBE"> 
      <table border="0" cellspacing="0" cellpadding="0" width="220">
        <tr bgcolor="#135086"> 
          <td width="14">&nbsp;</td>
          <td nowrap valign="middle" bgcolor="#135086"> 
            <p><font color="#DBEAF2"><%=tit%></font></p>
          </td>
          <td width="14">&nbsp;</td>
        </tr>
      </table>
      <table width="95%" border="1" cellspacing="0" cellpadding="5" class="base_text" bordercolor="#135086" height="300">
        <tr> 
          <td valign="top" bgcolor="#FFFFFF" bordercolor="#FFFFFF" align="left"> 
            <table width="100%" border="0" bordercolor="#FFFFFF" align="center" cellspacing="0" class="base_text">
              <tr valign="top" bordercolor="#FFFFFF"> 
                <td> 
                  <h1><font color="#232073"><%=pname%></font></h1>
                </td>
              </tr>
              <tr valign="top" bordercolor="#FFFFFF"> 
                <td height="55"> 
                  <p> 
                    <%if (imgLname != "") {%>
                    <img src="<%=imgLname%>" align="left" border="1" > 
                    <%}else{%>
                    &nbsp; 
                    <%}%>
                    <%=news%> </p>
                </td>
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
            <table border="0" cellspacing="2" cellpadding="2" width="300">
              <tr> 
                <td bgcolor="#FFFFFF" nowrap> 
                  <p><b><font color="#999999">&copy; <%=pdat%>&nbsp; <%=autor%></font></b></p>
                </td>
              </tr>
            </table>
            <hr noshade size="1" width="100%">
            <%if (usok) {%>
            <table width="100%" border="2" align="center" bordercolor="#FFFFFF">
              <tr> 
                <td height="2" bordercolor="#CCCCCC"> 
                  <div align="center"> 
                    <p><b>Публикация размещена в блоках:</b></p>
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
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="12">
        <tr> 
          <td></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="151" height="1" bgcolor="#000000"></td>
    <td height="1" width="1" bgcolor="#000000"></td>
    <td height="1" bgcolor="#000000"></td>
  <tr bgcolor="#D1B789"> 
    <td colspan="3"> 
      <div align="center"> 
        <p>&nbsp;</p>
      </div>
    </td>
  </tr>
  <tr> 
    <td width="151" height="1" bgcolor="#000000"></td>
    <td height="1" width="1" bgcolor="#000000"></td>
    <td height="1" bgcolor="#000000"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor="#000000"> 
    <td width="300"> 
      <p>&nbsp;</p>
    </td>
    <td align="RIGHT"> 
      <p><b><font color="#FFFFFF"><%=sminame%></font></b></p>
    </td>
  </tr>
</table>
<p align="CENTER"><font face="Arial, Helvetica, sans-serif" size="1"> 
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
	Records.MoveNext()
%>
  <a href="<%=url%>"><%=hname%></a> | 
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
	
	Records.MoveNext()
%>
  <a href="<%=url%>"><%=hname%></a> | 
  <%
} Records.Close()
delete recs
%>
  </font><font face="Arial, Helvetica, sans-serif"> </font></p>
<p align="CENTER"><font size="1">| <a href="pricemap.asp"><b>Каталог</b></a><b> 
  | <a href="area.asp">Администратор</a> | <a href="message.asp">Обратная связь</a> 
  | <a href="org.asp">Реквизиты</a> |</b></font></p>
<p align="center"><font size="1">&copy; 2003 программирование <a href="http://www.rusintel.ru">Русинтел</a></font></p>
<hr size="1" noshade align="center" width="468">


</body>
</html>
