<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
var hid=parseInt(Request("hid"))
if (isNaN(hid)) {hid=0}
if (hid==0){Response.Redirect("index.asp")}

var usok=false
var sql=""

// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=17
// +++  smi_id - ��� ��� � ������� SMI !!

var tit=""
var ts=""
var ErrorMsg=""
var id=0
var ShowForm=true
var sminame=""
var isFirst=false
var ddt = new Date()
var pdate=""
var str=""
var hiadr=""
var stt=0
var autor=""
var url=""
var digest=""
var news=""
var coment=""
var filename=""
var isnews=1
var id_usr=0
var ishtml=0

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="addpub.asp?hid="+hid
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<7) {usok=true}
	if (Session("tip_mem_pub")<3) {stt=1}
	id_usr=TextFormData(Session("id_mem_pub"),"0")
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			stt=1
		}
		Records.Close()
	} else {
		usok=true
		stt=1
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
hiname=String(Records("NAME").Value)
isnews=Records("ISNEWS").Value
hiadr=TextFormData(Records("URL").Value,"pubheading.asp")
hiadr+="?hid="+hid
Records.Close()
tit+=" | "+hiname

if (!usok) {Response.Redirect("index.asp")}

str=String(ddt.getMonth()+1)
if (str.length==1) {str="0"+str}
pdate="."+str+"."+ddt.getYear()
str=String(ddt.getDate())
if (str.length==1) {str="0"+str}
pdate=str+pdate

isFirst=String(Request.Form("Submit"))=="undefined"
if (!isFirst) {
	name=TextFormData(Request.Form("name"),"")
	autor=TextFormData(Request.Form("autor"),"")
	url=TextFormData(Request.Form("url"),"")
	pdate=TextFormData(Request.Form("pdate"),"")
	digest=TextFormData(Request.Form("digest"),"")
	news=TextFormData(Request.Form("news"),"")
	coment=TextFormData(Request.Form("coment"),"")
	ishtml=TextFormData(Request.Form("ishtml"),"0")
	
	// ������ ����������, ���� ���� ��������� HTML  ******  while (news.indexOf("<")>=0) {news=news.rplace("<","&lt;")}
	while (name.indexOf("'")>=0) {name=name.replace("'","\"")}
	while (autor.indexOf("'")>=0) {autor=autor.replace("'","\"")}
	while (digest.indexOf("'")>=0) {digest=digest.replace("'","\"")}
	while (coment.indexOf("'")>=0) {coment=coment.replace("'","\"")}
	while (pdate.indexOf(",")>=0) {pdate=pdate.replace(",",".")}
	while (pdate.indexOf("/")>=0) {pdate=pdate.replace("/",".")}
	while (pdate.indexOf("-")>=0) {pdate=pdate.replace("-",".")}
	
	if (name.length<3) {ErrorMsg+="������� �������� ����.<br>"}
	if (name.length>100) {ErrorMsg+="������� ������� ����.<br>"}
	if (autor.length>100) {ErrorMsg+="������� ������� ��� ������.<br>"}
	if (url.length>80) {ErrorMsg+="������� ������� URL.<br>"}
	if (coment.length>200) {ErrorMsg+="������� ������� ����������.<br>"}
	if (digest.length>250) {ErrorMsg+="������� ������� ��������.<br>"}
	if (digest.length<5) {ErrorMsg+="������� �������� ��������.<br>"}
	if (news.length>120000) {ErrorMsg+="������� ������� ����� ����������.<br>"}
	if (news.length<15) {ErrorMsg+="������� �������� ����� ����������.<br>"}
	if (!/(\d{2}).(\d{2}).(\d{4})$/.test(pdate)) {ErrorMsg+="�������� ������ ���� ����������.<br>"}
	
	if (ErrorMsg=="") {
		id=NextID("PUBID")
		filename=PubFilePath+id+".pub"
		var fs= new ActiveXObject("Scripting.FileSystemObject")
		sql=inspub
		sql=sql.replace("%ID",id)
		sql=sql.replace("%NAM",name)
		sql=sql.replace("%DIG",digest)
		sql=sql.replace("%PDAT",pdate)
		sql=sql.replace("%ST",stt)
		sql=sql.replace("%HID",hid)
		sql=sql.replace("%COM",coment)
		sql=sql.replace("%AUT",autor)
		sql=sql.replace("%URL",url)
		sql=sql.replace("%USR",id_usr)
		sql=sql.replace("%ISH",ishtml)
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
			ts=fs.OpenTextFile(filename,2,true)
			ts.Write(news)
			ts.Close()
		}
		catch(e){
			Connect.RollbackTrans()
			ErrorMsg+=ListAdoErrors()
		}
		if (ErrorMsg==""){
		  Connect.CommitTrans()
		  ShowForm=false
		}
	}
}

%>

<html>
<head>
<title>���������� ����������. <%=tit%></title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#CCCCCC" bordercolor="#CCCCCC">
      <p><a href="index.asp">�� ������� ��������</a> | <%=sminame%> 
        <%if(hiname!=""){%>
        | <a href="<%=hiadr%>"><%=hiname%></a> 
        <%}%>
      </p>
    </td>
  </tr>
</table>
<h1 align="center">���������� ���������� � ������ <a href="<%=hiadr%>"><%=hiname%></a> </h1>
<p> 
  <%if(ErrorMsg!=""){%>
</p>
<center>
  <p> <font color="#FF3300" size="2"><b>������!</b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%>
<%if(ShowForm){%>
<form name="form1" method="post" action="addpub.asp">
  <p align="center">��� ���������� �������, ����������, ��������� ���� �����: 
    <input type="hidden" name="hid" value="<%=hid%>">
  </p>
  <table width="870" border="1" bordercolor="#FFFFFF" align="center">
    <tr> 
      <td width="250" bgcolor="#CCCCCC" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>���������:</b></p>
        </div>
      </td>
      <td width="6">&nbsp;</td>
      <td bgcolor="#CCCCCC" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>��������:</b></p>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">��������� ����������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" maxlength="100" size="45">
          �� 250 ��������</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2">�����:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="autor" value="<%=isFirst?"":Request.Form("autor")%>" maxlength="100" size="45">
          �� 100 ��������</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2">URL ����������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="url" value="<%=isFirst?"":Request.Form("url")%>" maxlength="80" size="45">
          �� 80 �������� (������ ����: &quot;http://www...&quot;)</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">���� ����������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="pdate" value="<%=isFirst?pdate:Request.Form("pdate")%>" maxlength="10" size="45">
          ����� ������� ������� ����</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">��������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" valign="top"> 
        <p> 
          <input type="text" name="digest" value="<%=isFirst?"":Request.Form("digest")%>" maxlength="190" size="45">
          �� 250 ��������</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" height="14" valign="top"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">����� ����������:</font><font size="2">&nbsp;&nbsp;</font></p>
          <p align="left">��� �������� ���������� ����� ����� ������� ����������� 
            � ��������� ��������� (Notepad, WORD � �.�.).</p>
          <p align="left">&nbsp;</p>
          <p align="left"><font size="1">��� �������� ����������� ����������� 
            ������ �� �����, ���������� ������� ������ �������� ������� (Enter).</font></p>
          <p align="left"><font size="1">��������� ����������������� ����� �������� 
            � ����������� �������� ��� �������� ������.</font></p>
        </div>
      </td>
      <td width="6" height="14" valign="top"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" height="14" valign="top" bgcolor="#CCCCCC"> 
        <p> 
          <textarea name="news" cols="40" rows="15"><%=isFirst?"":Request.Form("news")%></textarea>
        </p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" height="14" valign="middle"> 
        <div align="right"> 
          <p><font size="2">����������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" height="14" valign="top"> 
        <p> 
          <input type="text" name="coment" value="<%=isFirst?"":Request.Form("coment")%>" maxlength="190" size="45">
          �� ����� ����������� �����</p>
      </td>
    </tr>
    <tr> 
      <td width="250" bordercolor="#333333" height="14" valign="middle"> 
        <div align="right"> 
          <p><font size="2">HTML �����</font><font size="2">:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#333333" height="14" valign="top"> 
        <p> 
          <input type="checkbox" name="ishtml" value="1">
          ��������, ���� <font size="2" color="#FF0000">&quot;����� ����������</font>&quot; 
          ���������� � ������� HTML !!</p>
      </td>
    </tr>
  </table>
  <p align="center"><font color="#FF0000">��������� ���������� ������� ������ 
    ����������� � ����������!</font></p>
  <p align="center"> 
    <input type="submit" name="Submit" value="���������">
  </p>
</form>
<%} 
else 
{%>
<center>
  <h1><font color="#3333FF">���������� ���������!</font></h1>
</center>
<%}%>
<hr width="780">
<p align="center"><a href="index.asp">��������� �� ������� ��������</a></p>
<p align="center"><a href="index.asp">��������� � �����������</a></p>
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
<p align="center"><font size="1"><b>| </b><a href="index.asp"><font size="1">� 
  ������</font></a><b> | </b><a href="message.asp">�������� �����</a><b> | </b><a href="org.asp">���������</a><b> |</b></font></p>
<p align="center"><font size="1">&copy; 2003 ���������������� <a href="http://www.rusintel.ru">��������</a></font></p>
<p align="center">&nbsp;</p>
<p align="center">&nbsp; </p>
</body>
</html>
