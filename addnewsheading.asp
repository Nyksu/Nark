<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\sql.inc" -->

<%
var hid=parseInt(Request("hid"))
if (isNaN(hid)) {hid=0}

var usok=false
var sql=""
// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=17
// +++  smi_id - ��� ��� � ������� SMI !!

var sminame=""
var tit=""
var hiname=""
var ErrorMsg=""
var ShowForm=true
var name=""
var url=""
var period=0
var pglen=0
var redactor=0
var id=0
var isnews=1

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="addnewsheading.asp?hid="+hid
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")==1) {usok=true}
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

tit=sminame

if (hid>0) {
	Records.Source="Select * from heading where id="+hid
	Records.Open()
	if (Records.EOF) {
		Records.Close()
		Response.Redirect("index.asp")
	}
	hiname=String(Records("NAME").Value)
	Records.Close()
	tit+=" | "+hiname
}

isFirst=String(Request.Form("Submit"))=="undefined"

if (!isFirst) {
	name=TextFormData(Request.Form("name"),"")
	url=TextFormData(Request.Form("url"),"")
	period=parseInt(Request.Form("period"))
	pglen=parseInt(Request.Form("pglen"))
	redactor=parseInt(Request.Form("redactor"))
	if (parseInt(Request.Form("isnews"))==1) {isnews=1} else {isnews=0}
	
	if (name.length<3) {ErrorMsg+="������� �������� ������������ �������.<br>"}
	if (isNaN(period)) {ErrorMsg+="�� ���������� ������  ��������.<br>"}
	if (isNaN(pglen)) {ErrorMsg+="�� ���������� ����� ��������.<br>"}
	if (pglen<1) {ErrorMsg+="�� ���������� ����� ��������.<br>"}
	if (isNaN(redactor)) {ErrorMsg+="�� ������ �������� �������.<br>"}
	
	if (ErrorMsg=="") {
		id=NextID("PUBHEADINGID")
		sql=insheading
		sql=sql.replace("%ID",id)
		sql=sql.replace("%NAM",name)
		sql=sql.replace("%SMI",smi_id)
		sql=sql.replace("%HID",hid)
		sql=sql.replace("%URL",url)
		sql=sql.replace("%PER",period)
		sql=sql.replace("%PL",pglen)
		sql=sql.replace("%USR",redactor)
		sql=sql.replace("%ISNEWS",isnews)
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
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
<title>���������� �������: (<%=tit%>)</title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
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
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>������!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 

<%if(ShowForm){%>
<h1 align="center"><b>�� ���������� ������� � : <%=tit%></b></h1>
 
<form name="form1" method="post" action="addnewsheading.asp">
  <p align="center">��� ���������� �������, ����������, ��������� ���� �����: 
    <input type="hidden" name="hid" value="<%=hid%>">
    <font size="2">
    <input type="hidden" name="url" value="<%=isFirst?"":Request.Form("url")%>" maxlength="50">
    </font> </p>
  <table width="780" border="1" bordercolor="#FFFFFF" align="center">
    <tr> 
      <td width="300" bgcolor="#CCCCCC" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>���������:</b></p>
        </div>
      </td>
      <td width="6"> 
        <p>&nbsp;</p>
      </td>
      <td bgcolor="#CCCCCC" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>��������:</b></p>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="300" bordercolor="#333333" height="14" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">������������ �������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" height="14" valign="top"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" maxlength="100" size="45">
        </p>
      </td>
    </tr>
    <tr> 
      <td width="300" bordercolor="#333333" valign="middle" height="5"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">������ ��������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="5"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" valign="top" height="5"> 
        <p><font size="2"> 
          <input type="text" name="period" value="<%=isFirst?"15":Request.Form("period")%>" maxlength="20">
          ���� 
          <input type="checkbox" name="isnews" value="1" checked>
          ��������� &quot;�����&quot;</font></p>
      </td>
    </tr>
    <tr> 
      <td width="300" bordercolor="#333333" valign="middle" height="14"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">����� �������� �������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="14"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" valign="top" height="14"> 
        <p><font size="2"> 
          <input type="text" name="pglen" value="<%=isFirst?"50":Request.Form("pglen")%>" maxlength="20">
          ����������</font></p>
      </td>
    </tr>
    <tr> 
      <td width="300" bordercolor="#333333" valign="middle" height="2"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">�������� �������:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="6" height="2"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" valign="top" height="2"> 
        <p><font size="2"> 
          <select name="redactor">
            <%
			Records.Source="Select ID, NAME from SMI_USR where USR_TYPE_ID<3 and SMI_ID="+smi_id+" order by NAME"
			Records.Open()
			while(!Records.EOF){%>
            <option value="<%=Records("ID").Value%>" 
				<%=!isFirst&&(Records("ID").Value==Request.Form("redactor"))?"selected":""%>> 
            <%=Records("NAME").Value%> </option>
            <%	Records.MoveNext()
			}
			Records.Close()
			%>
          </select>
          </font></p>
      </td>
    </tr>
  </table>
  <p align="center"><font color="#FF0000">��������� ���������� ������� ������ 
    ����������� � ����������!</font></p>
  <p align="center"> 
    <input type="submit" name="Submit" value="���������">
  </p>
  <hr width="780">
</form>
<%} 
else 
{%>
<center>
  <h1><font color="#3333FF">������� ���������!</font></h1>
</center>
<%}%>
<p>&nbsp;</p>
<p align="center"><a href="index.asp">��������� � �����������</a></p>
<p align="center">&nbsp;</p>
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
</body>
</html>
