<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%

if ((Session("is_adm_mem")!=1)&&(Session("is_host")!=1)){
Session("backurl")="area.asp"
Response.Redirect("login.asp")
}

%>


<html>
<head>
<title>����������������� �����.</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr bgcolor="#135086"> 
    <td valign="middle" width="500"> 
      <p>&nbsp;</p>
    </td>
    <td bgcolor="#135086" align="right"> <font></font> 
      <p><font color="#DBEAF2"></font></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr bgcolor="#135086"> 
    <td valign="middle" width="500" bgcolor="#DBEAF2"> 
      <p>&nbsp; </p>
    </td>
    <td bgcolor="#DBEAF2" align="right"> <font></font> 
      <p><font color="#135086"><b>������� ��������������</b></font></p>
    </td>
  </tr>
</table>
<table width="100%" border="1"> <tr> <td bordercolor="#333333"> <div align="left"> 
<p><font color="#0000CC"><b>������������, <%=Session("name_mem")%> !</b></font></p></div></td></tr> 
</table><table width="100%" border="1"> <tr bordercolor="#333333" bgcolor="#CCCCCC"> 
<td width="50%"> <div align="center"> <p><b>����������</b></p></div></td><td width="50%"> 
<div align="center"> <p><b>�����������</b></p></div></td></tr> <tr bordercolor="#333333" valign="top"> 
    <td width="50%"> 
      <ul>
        <li> 
          <p><a href="pswusr.asp?usr=<%=Session("id_mem")%>">�������� ������</a></p>
        </li>
        <li> 
          <p><a href="index.asp">������������� ����</a></p>
        </li>
      </ul>
    </td>
    <td width="50%"> 
<ul> <li> 
          <p><a href="edcomp.asp">�������� ���������</a></p>
        </li><LI> 
<P><A HREF="bloknews.asp">�������������� �����</A></P></LI><LI><P><A HREF="pubarea.asp">������������� 
�������������� ���������</A></P></LI></ul></td></tr> </table>
<div align="center"> 
  <p>&nbsp;</p>
  <p><font color="#FFFFFF"><a href="index.asp">�� ������� ��������</a></font><a href="index.asp"></a> 
  </p>
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
  <p><font size="1"><b>| </b><a href="index.asp"><font size="1">� ������</font></a><b> 
    | </b><a href="message.asp">�������� �����</a><b> | </b><a href="org.asp">���������</a><b> 
    |</b></font></p>
  <p align="center"><font size="1">&copy; 2003 ���������������� <a href="http://www.rusintel.ru">��������</a></font></p>
  <p align="center">&nbsp;</p>
  <p>&nbsp; </p>
</div>
</body>
</html>
