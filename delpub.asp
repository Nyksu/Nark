<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=17
// +++  smi_id - ��� ��� � ������� SMI !!

var id=parseInt(Request("pid"))
if (isNaN(id)) {Response.Redirect("index.asp")}

var name=""
var hid=0
var retname="�������� � �������"
var sql=""
var ShowForm=true
var isok=true
var path=""
var hdd=0
var hadr=""
var nm=""

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="delpub.asp?pid="+id
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<3) {usok=true}
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

if (!usok) {Response.Redirect("area.asp")}

sql="Select t1.* from publication t1, heading t2 where t1.id="+id+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("area.asp")
}
hid=Records("heading_id").Value
name=String(Records("NAME").Value)
Records.Close()

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> | "+path
	hdd=Records("HI_ID").Value
	Records.Close()
}

if (String(Request.Form("next1")) != "undefined") {
	if (Request.Form("agr")==1) {
		sql="Delete from publication where id="+id
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
		}
		catch(e){
			Connect.RollbackTrans()
			isok=false
		}
		if (isok){
			Connect.CommitTrans()
			var ShowForm=false
		}
	} else {Response.Redirect("area.asp")}
}

%>

<Html>
<Head>
<Title>�������� ������� <%=name%></Title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><a href="index.asp">�� ������� ��������</a> | <font color="#FF0000"><%=path%></font></p>
    </td>
  </tr>
</table>
<h1 align="center"><font size="3">�������� ����������</font></h1>

<p align="center">��������� � <font color="#FF0000"><%=path%></font></p>

<% if (ShowForm) {%>
<p align="center">&nbsp;</p>
<p align="center"><b>�� ������������� ������ ������� ����������: <font color="#FF0000"><%=name%></font> ?</b></p>
<p>&nbsp;</p>
<table width="500" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF">
  <tr>
    <td bordercolor="#333333">
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;��������!</font></b> 
        �������� ����������!</p>
      <p align="left">�������� ���������� ���������� � �� ���������� ����� �����������!</p>
      <p align="left">���� �� ������ ������� ����������, �� ��������� ������ � 
        ��������������� ���� � ������� ������ &quot;����������&quot;.</p>
</td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<form name="form1" method="post" action="delpub.asp">
  <input type="hidden" name="pid" value="<% =id %>">
  <p align="center"> 
    <input type="checkbox" name="agr" value="1">
    ��, � ���� ������� ��� ����������: (<b><font color="#FF0000" size="3"><%=name%></font></b>) 
    !</p>
  <p align="center">&nbsp;</p>
  <p align="center"> 
    <input type="submit" name="next1" value="����������">
  </p>
</form>
<%} else {%>
<p>&nbsp;</p>
<p align="center"><font color="#FF0000"><b>������ ������!</b></font></p>
<%}%>
<p align="center"><a href="index.asp">��������� � �����������</a></p>
<p align="center"><a href="area.asp">������� ��������������</a></p>
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

</Body>
</Html>
