<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\path.inc" -->

<%
// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=17
// +++  smi_id - ��� ��� � ������� SMI !!

var bk=parseInt(Request("bk"))
if (isNaN(bk)) {Response.Redirect("bloknews.asp")}

var sminame=""
var usok=false
var id_usr=0
var sql=""
var bid=0
var name=""
var coment=""
var sname=""
var imgname=""
var pid=0
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var hid=0
var hdd=0
var path=""
var pos=0
var autor=""

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="block.asp?bk="+bk
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<5) {usok=true}
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

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

Records.Source="Select * from block_news where smi_id="+smi_id+" and id="+bk
Records.Open()
if (Records.EOF ) {
	Records.Close()
	Response.Redirect("bloknews.asp")
}
name=String(Records("SUBJ").Value)
coment=String(Records("NAME").Value)
Records.Close()

%>

<html>
<head>
<title>���� ��������</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">�� ������� ��������</a> 
        | <a href="area.asp">������� ��������������</a> | <a href="bloknews.asp">�������������� �����</a> </font></p>
    </td>
  </tr>
</table>
<h1 align="center">����: <%=sminame%></h1>
<h1 align="center">�������������� ����: <font color="#0000CC"><%=name%></font></h1>
<p align="center"><%=coment%></p>
<table width="780" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF">
  <tr> 
    <td bordercolor="#333333" bgcolor="#EFE2CF"> 
      <p><b>���������� / �������� ������� ����� &quot;<%=name%>&quot;</b></p>
    </td>
  </tr>
</table>
<table width="780" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF">
  <tr>
    <td bordercolor="#333333">
      <p>������� � �����: || 
        <%
Records.Source="Select * from news_pos where block_news_id="+bk+" order by posit"
Records.Open()
while (!Records.EOF )
{
	pos=String(Records("POSIT").Value)
%>
        <font color="#FF0000"><%=pos%></font> || 
        <%
	Records.MoveNext()
} 
Records.Close()%>
      </p>
      <form name="delpos" method="post" action="delposbloq.asp">
        <p>������� ������� �����: 
          <input type="hidden" name="bk" value="<%=bk%>">
          <input type="hidden" name="pos" value="<%=pos%>" >
          <font color="#FF0000"><%=pos%></font>&nbsp;&nbsp;&nbsp; 
          <input type="submit" name="del_btn" value="�������">
          (!��������������� �������!)</p>
      </form>
      <form name="addpos" method="post" action="addposbloq.asp">
        <p><b><font color="#CC3333">�������� ������� � ����:</font></b> 
          <input type="hidden" name="bk" value="<%=bk%>">
          <input type="text" name="pos" value="<%=parseInt(pos)+1%>" maxlength="2" size="10">
          <input type="submit" name="add_btn" value="��������">
          (!��������������� �������!) </p>
      </form>
      <form name="addpos" method="post" action="clearposbloq.asp">
        <p>�������� ����: 
          <input type="hidden" name="bk" value="<%=bk%>">
          <input type="submit" name="clear_btn" value="��������">
        </p>
      </form>
</td>
  </tr>
</table>
<p>&nbsp;</p><table width="780" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF">
  <tr>
    <td bordercolor="#333333" bgcolor="#EFE2CF"> 
      <p><b>���������� � ����� &quot;<%=name%>&quot;</b></p>
    </td>
  </tr>
</table>
<%
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	pos=String(Records("POSIT").Value)
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
<table width="780" border="1" bordercolor="#FFFFFF" align="center">
  <tr bordercolor="#333333" bgcolor="#EFE2CF"> 
    <td colspan="2"> 
      <p><b><font color="#FFFF00"> <font color="#660000">�������:</font></font><%=pos%> 
        ||</b> <a href="delbkpub.asp?pid=<%=pid%>&bk=<%=bk%>&tp=1">������� �� 
        �����</a> || <a href="delbkpub.asp?pid=<%=pid%>&bk=<%=bk%>">����������� 
        � ������ ����</a> || <a href="delbkpub.asp?pid=<%=pid%>&ps=<%=pos%>&bk=<%=bk%>">�������� 
        �������</a> ||</p>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF" bordercolor="#CCCCCC"> 
    <td colspan="2"> 
      <p><b>���������: <%=pname%> || </b>���� ������<b> : <%=pdat%></b></p>
    </td>
  </tr>
  <tr bordercolor="#CCCCCC"> 
    <td> 
      <p>���� ��������: <font color="#009900"><b><%=cdat%></b></font> || �������� 
        (�����):<b> <%=autor%></b></p>
    </td>
    <td>
      <p align="center"><b>���� ���������</b></p>
    </td>
  </tr>
  <tr valign="top"> 
    <td bordercolor="#CCCCCC" width="600"> 
      <p><font size="2" > <b>��������:</b> <%=digest%> <a href="<%=url%>">������ 
        �����</a></font></p>
      <p> <br><b>���������� ��������� � �������</b> <b>&gt;</b><font color="#FF0000"><b> 
        <%=path%></b></font></p>
    </td>
    <td bordercolor="#CCCCCC"> 
      <div align="center"> 
        <p> 
          <%if (imgname != "") {%>
          <img src="<%=imgname%>" > 
          <%}else{%>
          No image 
          <%}%>
        </p>
        <p><a href="addpubimg.asp?pid=<%=pid%>">�������� 
          / �������� ����</a> </p>
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
<p>&nbsp; </p>
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
<p align="CENTER"><font size="1"><b>| <a href="index.asp"><font face="Arial, Helvetica, sans-serif" size="1"> 
  <b>� ������</b></font></a> | <a href="message.asp">�������� �����</a> | <a href="org.asp">���������</a> |</b></font></p>
<p align="center"><font size="1">&copy; 2003 ���������������� <a href="http://www.rusintel.ru">��������</a></font></p>
<hr size="1" noshade align="center" width="468">


  </body>
</html>
