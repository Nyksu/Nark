<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\sql.inc" -->

<%
var pid=Request("pid")

// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=17
// +++  smi_id - ��� ��� � ������� SMI !!

var hid=0
var hdd=0
var sql=""
var name=""
var articul=""
var srname=""
var sminame=""
var pname=""
var nm=""
var path=""
var imgname=""
var ext=""
var size=0
var ErrorMsg=""
var posit=0
var pref=""
var usok=false

if (isNaN(parseInt(pid))) {
	if (Request.ServerVariables("REQUEST_METHOD")=="POST") {
	   updown = Server.CreateObject("ANUPLOAD.OBJ")
		pid=updown.Form("pid")
		posit=updown.Form("posit")
		if (posit==1) {pref=""}
		if (posit==2) {pref="l"}
		if (posit==3) {pref="r"}
		if (!isNaN(parseInt(pid))) {
		try {
			updown.Delete(PubFilePath+pref+pid+".gif")
			updown.Delete(PubFilePath+pref+pid+".jpg")
			updown.SavePath = PubFilePath
			size=parseInt(updown.GetSize("file"))
			ext=updown.GetExtension("file").toUpperCase()
			if(ext!="JPG" && ext!="GIF"){throw "����������� ������ JPG ��� GIF �����."}
			if (size>20480){throw "�� ����� 20kB."}
			if (size==0){throw "��� �����."}
			updown.SaveAs("file",PubFilePath+pref+pid+"."+ext)
			Response.Redirect("newshow.asp?pid="+pid)
			}
    	catch(e){ErrorMsg+=String(e.message)=="undefined"?e:e.message}
		
		} else {Response.Redirect("index.asp")}
	}
	else {Response.Redirect("index.asp")}
}

Records.Source="Select t1.* from publication t1, heading t2 where t1.id="+pid+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("index.asp")
}
hid=Records("heading_id").Value
pname=String(Records("NAME").Value)
Records.Close()

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="addpubimg.asp?pid="+pid
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<4) {usok=true}
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

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		path=nm+" | "+path
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
	}
	else {
		path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> | "+path
	}
	hdd=Records("HI_ID").Value
	Records.Close()
}

path="<a href=\"index.asp\">"+sminame+"</a> | "+path

%>


<html>
<head>
<title>�������� ���������� � ����������: <%=pname%></title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#DBEAF2"> 
    <td colspan="3" height="17" bgcolor="#CCCCCC" bordercolor="#CCCCCC"> 
      <p align="left"><%=path%>
        ����������: <a href="newshow.asp?pid=<%=pid%>"> <%=pname%></a></p>
    </td>
  </tr>
</table>
<%if(ErrorMsg!=""){%>
<center>
  <p> <font color="#FF3300" size="2"><b>������!</b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%>
<form name="form1" method="post" action="addpubimg.asp" enctype="multipart/form-data">
  <h1> 
    <input type="hidden" name="pid" value="<% =pid %>">
  </h1>
  <h1 align="center"> <b>�������� ���������� � ����������: </b></h1>
  <p align="center"><font size="4" color="#CC3300"><b><%=pname%></b></font></p>
  <div align="center"> 
    <p>����������� ����� GIF ��� JPG �� 20 ��������</p>
  </div>
  <div align="center"> 
    <table width="500" border="2" bordercolor="#FFFFFF">
      <tr> 
        <td bordercolor="#333333" bgcolor="#CCCCCC"> 
          <div align="center"> 
            <p><b>�������� ���������� �� ����� ���������� ��� �������� �� ������:</b></p>
          </div>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#333333"> 
          <p align="left">1. ������� ���������� ����������: 
            <select name="posit">
              <option value="1">��������� ��� ���������</option>
              <option value="2">�������� ����� �� ������</option>
              <option value="3">�������������� ����� �� ������</option>
            </select>
          </p>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#333333"> 
          <div align="left"> 
            <p>2. 
              <input type="file" name="file" size="40">
            </p>
          </div>
        </td>
      </tr>
    </table>
  </div>
  <p align="center"> 
    <input type="submit" name="Submit" value="���������">
  </p>
</form>
<p align="center"><font size="1">��� �������� ���������� �������� ������� � ������� 
  &quot;���������&quot; ��� ������ ����� </font></p>
<table width="500" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF">
  <tr>
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p align="center"><b>������� ���������� � ������������</b></p>
    </td>
  </tr>
  <tr> 
    <td> 
      <ul>
        <li><font face="Arial, Helvetica, sans-serif" size="2"><b>��������� ��� 
          ���������</b> - ��������� �� ��������� �������� ������ � � �������������� 
          ������ ����� �� ��������� ����������. ������������� ������: ������ �� 
          150 ��������</font></li>
        <li><font face="Arial, Helvetica, sans-serif" size="2"><b>�������� ���������� 
          ����� �� ������</b> - ��������� �� �������� ����������. ������������� 
          ������: ������ �� 300 ��������</font></li>
        <li><font face="Arial, Helvetica, sans-serif" size="2"><b>�������������� 
          ���������� ����� �� ������</b> - ��������� � ������ ����� ����������. 
          ����� ������� ��� �������� �������������� ����������� ���������� (�������, 
          ����� � �.�.). ��� ����������� ����������� ������� �������� ������������� 
          ������������ �������������-��������������� ����. ������������ ������ 
          �� 400 ����.</font></li>
      </ul>
    </td>
  </tr>
</table>
<hr width="500">
<p align="center">&nbsp;</p>
<p align="center"><font size="2"><a href="newshow.asp?pid=<%=pid%>">��������� 
  � ����������</a></font></p>
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
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
</body>
</html>
