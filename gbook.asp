<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
// ��� �������� �����
var guestbook=3
// �������� ��� ������ ������
var isadm=0
var sql=""
var tbook=1
var namebook=""
var kvo=0
var name=""
var id=parseInt(Request("id"))
var op=parseInt(Request("op"))
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}
var dat=""
var email=""
var city=""
var txt=""
var st=0
var ii=0
var ErrorMsg=""
var namop=""
var plen=10
var pp=0
var tkvo=0


var tps=parseInt(Request("tps"))
var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
if (isNaN(tps)) {tps=1}
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}


if (String(Session("gbook"))=="undefined") {Session("gbook")=0}
if ((Session("is_adm_mem")==1) || (Session("is_host")==1)) { isadm=1 }

if (isNaN(id)) {id=0}
if (isNaN(op)) {op=0}
if (Session("gbook")>1) {op=0}
if ((op>1) && (isadm==0)) {op=0}
if ((op>2) && (id==0)) {op=0}
if ((id>0) && (op==0)) {id=0}

switch  (op) {
	case 1 : namop="���������� ��������� � �������� �����"; break;
	case 2 : namop="��������� ���������� �������� �����"; break;
	case 3 : namop="�������� ���������"; break;
	case 4 : namop="���������� ���������"; break;
	case 5 : namop="��������� ���������"; break;
	case 6 : namop="��������� ���������"; break;
}

sql="Select * from guestbook where id="+guestbook+" and enterprise_id="+company+" and state=1"
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("index.asp")
}
tbook=Records("STATE_DEF").Value
namebook=Records("NAME").Value
plen=parseInt(Records("PAGELEN").Value)
Records.Close()

if (id>0) {
	sql="Select * from gbmsg where guestbook_id="+guestbook+" and id="+id
	Records.Source=sql
	Records.Open()
	if (Records.EOF){
		Records.Close()
		Response.Redirect("gbook.asp")
	}
	name=Records("AUTORNAME").Value
	id=Records("ID").Value
	dat=Records("POSTDATE").Value
	email=TextFormData(Records("EMAIL").Value,"")
	city=TextFormData(Records("CITY").Value,"")
	txt=TextFormData(Records("COMENT").Value,"")
	st=Records("STATE").Value
	Records.Close()
}

var isFirst=String(Request.Form("Submit"))=="undefined"
if ((!isFirst) && (op>0)){
	sql=""
	if ((op==3) && (Request.Form("kill")==666)) {
		// �������� 
		sql="Delete from gbmsg where id="+id
	} else { if (op==3) {Response.Redirect("gbook.asp")}}
	if ((op==4) && (Request.Form("lockit")==33)) {
		// ����������� 
		sql="Update gbmsg set state=0 where id="+id
	} else { if (op==4) {Response.Redirect("gbook.asp")}}
	if (op==5 && Request.Form("unlockit")==333) {
		// ��������� 
		sql="Update gbmsg set state=1 where id="+id
	} else { if (op==5) {Response.Redirect("gbook.asp")}}
	if (op==2) {
		// ��������� ��������� �������� �����
		name=TextFormData(Request.Form("name"),"")
		tbook=Request.Form("tips")==1?1:0
		plen=parseInt(Request.Form("plen"))
		if (isNaN(plen)) {ErrorMsg+="�������� ������ ����� ��������.<br>"}
		if (name.length>100) {ErrorMsg+="������� ������� ������������ �������� �����.<br>"}
		if (name.length<2) {ErrorMsg+="������� �������� ������������ �������� �����.<br>"}
		if (ErrorMsg=="") {
		 sql="Update guestbook set name='%NAME', state_def=%SD, pagelen=%PL where id="+guestbook
		 sql=sql.replace("%NAME",name)
		 sql=sql.replace("%SD",tbook)
		 sql=sql.replace("%PL",plen)
		}
	}
	if (op==1) {
		// ���������� ���������
		name=TextFormData(Request.Form("name"),"")
		email=TextFormData(Request.Form("email"),"")
		city=TextFormData(Request.Form("city"),"")
		txt=TextFormData(Request.Form("txt"),"")
		while (txt.indexOf("/n")>=0) {txt=txt.rplace("/n"," ")}
		if (name.length>100) {ErrorMsg+="������� ������� ���.<br>"}
		if (name.length<2) {ErrorMsg+="������� �������� ���.<br>"}
		if (email.length>100) {ErrorMsg+="������� ������� E-mail.<br>"}
		if ((email != "") && (email.length<7)) {ErrorMsg+="������������ E-mail.<br>"}
		if (email.length>0) {	 
			if (!/(\w+)@((\w+).)*(\w+)$/.test(email)) {ErrorMsg+="�������� ������ ���� 'E-mail'.<br>"}}
		if (city.length>100) {ErrorMsg+="������� ������� ������������ ������ (������).<br>"}
		if (city.length<3) {ErrorMsg+="������� �������� ������������ ������ (������).<br>"}
		if (txt.length>500) {ErrorMsg+="������� ������� ���������.<br>"}
		if (txt.length<3) {ErrorMsg+="������� �������� ���������.<br>"}
		if (ErrorMsg=="") {
			id=NextID()
			sql="Insert into gbmsg (id,guestbook_id,autorname,email,city,postdate,state,coment) "
			sql+="values (%ID, %GB, '%NAME', '%EMAIL', '%CITY', 'NOW', %ST, '%TXT')"
			sql=sql.replace("%NAME",name)
			sql=sql.replace("%ID",id)
			sql=sql.replace("%GB",guestbook)
			sql=sql.replace("%EMAIL",email)
			sql=sql.replace("%CITY",city)
			sql=sql.replace("%ST",tbook)
			sql=sql.replace("%TXT",txt)
			Session("lastadd")=id
		}
	}
	if (op==6) {
		// ���������� ���������
		name=TextFormData(Request.Form("name"),"")
		email=TextFormData(Request.Form("email"),"")
		city=TextFormData(Request.Form("city"),"")
		txt=TextFormData(Request.Form("txt"),"")
		while (txt.indexOf("/n")>=0) {txt=txt.rplace("/n"," ")}
		if (name.length>100) {ErrorMsg+="������� ������� ���.<br>"}
		if (name.length<2) {ErrorMsg+="������� �������� ���.<br>"}
		if (email.length>100) {ErrorMsg+="������� ������� E-mail.<br>"}
		if ((email != "") && (email.length<7)) {ErrorMsg+="������� �������� E-mail.<br>"}
		if (city.length>100) {ErrorMsg+="������� ������� ������������ ������ (������).<br>"}
		if (city.length<3) {ErrorMsg+="������� �������� ������������ ������ (������).<br>"}
		if (txt.length>500) {ErrorMsg+="������� ������� ���������.<br>"}
		if (txt.length<3) {ErrorMsg+="������� �������� ���������.<br>"}
		if (ErrorMsg=="") {
			sql="Update gbmsg set guestbook_id=%GB ,autorname='%NAME' ,email='%EMAIL' ,city='%CITY' ,coment='%TXT'  "
			sql+="where id=%ID"
			sql=sql.replace("%NAME",name)
			sql=sql.replace("%GB",guestbook)
			sql=sql.replace("%ID",id)
			sql=sql.replace("%EMAIL",email)
			sql=sql.replace("%CITY",city)
			sql=sql.replace("%TXT",txt)
		}
	}
	if ((ErrorMsg=="") && (sql!="")) {
			Connect.BeginTrans()
			try{
				Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="������ ��������.<br>"
			}
			if (ErrorMsg==""){
				Connect.CommitTrans()
				if (isadm==0 && op==1) {
					Session("gbook")+=1					
				}
		  		Response.Redirect("gbook.asp")
			}
	} 
}

%>

<html>
<head>
<title>�������� �����</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="3" leftmargin="1">
<table border="0" cellspacing="0" cellpadding="0" bordercolor="#666699" width="780">
  <tr> 
    <td background="images/pl_bg.gif" align="right" bgcolor="#666699"> 
      <p  class="secondarystories"><a href="gbook.asp"><font color="#FFFFFF">�������� 
        �����</font></a> * <a href="message.asp"><font color="#FFFFFF">��� ������</font></a> 
        * <a href="javascript:window.external.AddFavorite(parent.location,document.title)"> 
        <font color="#FFFFFF">�������� � ���������</font></a>:</p>
    </td>
  </tr>
</table>
<table width="780" border="1" cellspacing="0" cellpadding="0" bordercolor="#666699">
  <tr bgcolor="#E2E2F1" bordercolor="#E2E2F1"> 
    <td height="16" width="400" bgcolor="#E2E2F1"> 
      <h1>&nbsp;&nbsp;&nbsp;<%=namebook%></h1>
    </td>
    <form name="form1" method="post" action="search.asp">
      <td height="16" align="right"> 
        <p><b><font color="#666699">�����</font></b><font color="#FFFFFF"><b> 
          &nbsp;</b></font> 
          <input type="text" name="sch" size="40" value="" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 114px" >
          <select name="wrds" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 100px">
            <option value="1" <%=wrds==1?"selected":""%>>��� �����</option>
            <option value="2" <%=wrds==2?"selected":""%>>���� �� ����</option>
            <option value="3" <%=wrds==3?"selected":""%>>����� �������</option>
          </select>
          <input type="submit" name="Findit" style="FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 50px; HEIGHT: 20px" value="�����">
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
          <td background="images/pl_bg.gif" width="250" align="center"><span class="secondarystories">&nbsp; <a href="/"><font color="#FFFFFF">�� �������</font></a></span></td>
          <td width="20"><img src="images/pl_ang.gif" width="20" height="16"></td>
          <td background="images/pl_bg_next.gif" width="484"></td>
        </tr>
      </table>
    </td>
    <td valign="bottom" width="12" bgcolor="#FFFFFF"></td>
    <td width="1" bgcolor="#123D87"></td>
  </tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0" height="350">
  <tr> 
    <td width="1"></td>
    <td width="12">&nbsp;</td>
    <td width="1" valign="top" bgcolor="#123D87"></td>
    <td valign="top" bgcolor="#FFFFFF"> 
      <h1 align="center"><b><font size="4" color="#0000CC"></font></b></h1>
      <%
if (isadm==1 && op!=2) {
%>
      <p><a href="gbook.asp?op=2&id=0"><font color="#006633">�������� ��������� 
        �������� �����</font></a></p>
      <%
}
%>
      <%
if (op==0) {
%>
      <p><a href="gbook.asp?op=1&id=0">�������� ����� ��������� � �������� �����</a></p>
      <%
sql="Select * from gbmsg where guestbook_id="+guestbook
if (isadm!=1) {
	if (String(Session("lastadd"))=="undefined") {sql+=" and state=1"}
	else {sql+=" and ((state=1 and id<>"+Session("lastadd")+") or (id="+Session("lastadd")+"))"}
}
sql+=" order by id desc"
Records.Source=sql
Records.Open()
if (!Records.EOF) {
kvo=Records.Recordcount
if ((pg+1)*plen > kvo) {tkvo=kvo} else {tkvo=(pg+1)*plen}
%>
      <p>��������� ���������: <font color="#000099"><b><%=kvo%></b></font> </p>
      <%
ii=0
while ((!Records.EOF) && (ii<tkvo)) {
	name=Records("AUTORNAME").Value
	id=Records("ID").Value
	dat=Records("POSTDATE").Value
	email=TextFormData(Records("EMAIL").Value,"")
	city=TextFormData(Records("CITY").Value,"")
	txt=TextFormData(Records("COMENT").Value,"")
	while (txt.indexOf("<")>=0) {txt=txt.rplace("<","&lt;")}
	st=Records("STATE").Value
	ii+=1
	if (ii>=(pg*plen+1)) {
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
        <tr> 
          <td width="40" bgcolor="#F2F2F2"> 
            <div align="center"> 
              <p><b><font size="2" color="#000099"><%=ii%>.</font></b></p>
            </div>
          </td>
          <td bgcolor="#F2F2F2"> 
            <p><b><font size="2" color="#0000FF"><%=dat%></font></b>&nbsp;&nbsp; 
              <font color="#000099"><b><%=name%></b></font> <font size="2">&nbsp;<a href="mailto:<%=email%>">e-mail</a> 
              ( <font color="#0000FF"><%=city%></font> )</font></p>
          </td>
        </tr>
        <% if (isadm==1) {%>
        <tr> 
          <td width="40"> 
            <div align="center"> 
              <p><font size="2"><b><font color="#FF0000"><%=st==0?"���":""%></font></b></font></p>
            </div>
          </td>
          <td> 
            <p><font size="2"><a href="gbook.asp?op=6&id=<%=id%>"><font color="#006633">��������</font></a> 
              | <a href="gbook.asp?op=3&id=<%=id%>"><font color="#006633">�������</font></a> 
              | <a href="gbook.asp?op=<%=st==1?"4":"5"%>&id=<%=id%>"> <%=st==1?"����������":"������������"%></a></font></p>
          </td>
        </tr>
        <%}%>
        <tr> 
          <td width="40"> 
            <p>&nbsp;</p>
          </td>
          <td bgcolor="#F9F9F9"> 
            <p><font size="2"><%=txt%></font></p>
          </td>
        </tr>
        <tr> 
          <td width="40" height="6"></td>
          <td height="6"></td>
        </tr>
      </table>
      <%
	}
	Records.MoveNext()
}
%>
      <p align="center">��������: <font size="2"> 
        <%
for ( pp=1; pp<(kvo/plen + 1) ; pp++) {
%>
        <% if (pp==(pg+1)) { %>
        <%=pp%> | 
        <%} else {%>
        <a href="gbook.asp?pg=<%=pp-1%>"><%=pp%></a> | 
        <%}%>
        <%
}
%>
        </font> 
        <%
} else {
%>
      </p>
      <p>��� ��������� � �������� �����</p>
      <%
}
Records.Close()
%>
      <%
} else {
// ��������� ���� �� ��������
%>
      <p><a href="gbook.asp">��������� � ������ ���������</a></p>
      <p align="center"><%=namop%></p>
      <table width="100%" border="0" cellspacing="1" cellpadding="1">
        <tr> 
          <td bgcolor="#F9F9F9"> 
            <% 
if (ErrorMsg!="") {
%>
            <div align="center"> 
              <table width="90%" border="1" cellspacing="1" cellpadding="1" bordercolor="#FF0000" bgcolor="#FFFFCC">
                <tr> 
                  <td> 
                    <p align="center"><b><font color="#FF0000" size="2">��������! 
                      �������� ��������� ������:</font></b></p>
                    <p align="center"><font color="#FF0000" size="2"><%=ErrorMsg%></font></p>
                  </td>
                </tr>
              </table>
            </div>
            <%
}
%>
            <form name="form1" method="post" action="gbook.asp">
              <input type="hidden" name="id" value="<%=id%>">
              <input type="hidden" name="op" value="<%=op%>">
              <%
if ((op==1) || (op==6)) {
%>
              <table width="90%" border="2" cellspacing="2" cellpadding="1" bordercolor="#F9F9F9" dwcopytype="CopyTableRow" align="center">
                <tr> 
                  <td bordercolor="#666699" bgcolor="#666699" width="200"> 
                    <div align="center"> 
                      <p><font color="#FFFFFF">���������</font></p>
                    </div>
                  </td>
                  <td width="3"> 
                    <div align="center"><b></b></div>
                  </td>
                  <td bordercolor="#666699" bgcolor="#666699"> 
                    <div align="center"> 
                      <p><font color="#FFFFFF">��������</font></p>
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td bordercolor="#E2E2F1" width="200" bgcolor="#FFFFFF"> 
                    <div align="right"> 
                      <p><font color="#FF0000">���� ��� ��� ���������:</font> 
                      </p>
                    </div>
                  </td>
                  <td width="3">&nbsp;</td>
                  <td bordercolor="#E2E2F1" bgcolor="#FFFFFF"> 
                    <p> 
                      <input type="text" name="name" value="<%=isFirst?name:Request.Form("name")%>" size="30" maxlength="100">
                      (�� 100 ��������)</p>
                  </td>
                </tr>
                <tr> 
                  <td bordercolor="#E2E2F1" width="200" bgcolor="#FFFFFF"> 
                    <div align="right"> 
                      <p><font color="#FF0000">����� (������):</font> </p>
                    </div>
                  </td>
                  <td width="3">&nbsp;</td>
                  <td bordercolor="#E2E2F1" bgcolor="#FFFFFF"> 
                    <p> 
                      <input type="text" name="city" value="<%=isFirst?city:Request.Form("city")%>" size="30" maxlength="100">
                      (�� 100 ��������)</p>
                  </td>
                </tr>
                <tr> 
                  <td bordercolor="#E2E2F1" width="200" bgcolor="#FFFFFF"> 
                    <div align="right"> 
                      <p>E-mail: </p>
                    </div>
                  </td>
                  <td width="3">&nbsp;</td>
                  <td bordercolor="#E2E2F1" bgcolor="#FFFFFF"> 
                    <p> 
                      <input type="text" name="email" value="<%=isFirst?email:Request.Form("email")%>" size="30" maxlength="100">
                      (�� 100 ��������)</p>
                  </td>
                </tr>
                <tr> 
                  <td bordercolor="#E2E2F1" width="200" valign="top" bgcolor="#FFFFFF"> 
                    <div align="right"> 
                      <p><font color="#FF0000">����� ������ ���������:</font> 
                      </p>
                    </div>
                  </td>
                  <td width="3">&nbsp;</td>
                  <td bordercolor="#E2E2F1" bgcolor="#FFFFFF"> 
                    <p> 
                      <input type="text" name="txt" value="<%=isFirst?txt:Request.Form("txt")%>" size="30" maxlength="500">
                      <br>
                      (�� 500 ��������)</p>
                  </td>
                </tr>
              </table>
              <%
}
%>
              <%
	if (op==3) {
%>
              <p><b>�������� ���������. <font color="#0000CC"><%=name%></font></b></p>
              <p><font size="2" color="#000099"><%=txt%></font> 
              <p> 
              <p> 
                <input type="checkbox" name="kill" value="666">
                ��! �������!</p>
              <%
	}
%>
              <%
	if (op==4) {
%>
              <p><b>���������� ���������. <font color="#0000CC"><%=name%></font></b></p>
              <p><font size="2" color="#000099"><%=txt%></font> 
              <p> 
              <p> 
                <input type="checkbox" name="lockit" value="33">
                ��! ������������� ���������!</p>
              <%
	}
%>
              <%
	if (op==5) {
%>
              <p><b>��������� ���������. <font color="#0000CC"><%=name%></font></b></p>
              <p><font size="2" color="#000099"><%=txt%></font> 
              <p> 
              <p> 
                <input type="checkbox" name="unlockit" value="333">
                ��! �������������� ���������!</p>
              <%
	}
%>
              <%
	if (op==2) {
%>
              <p><b>��������� �������� �����:</b></p>
              <p> 
                <input type="text" name="name" size="30" maxlength="100" value="<%=isFirst?namebook:Request.Form("name")%>">
                ��� �������� ����� (�� 100 ��������)</p>
              <p> 
                <input type="checkbox" name="tips" value="1" <%=tbook==0?"":"checked"%>>
                ��������� ����������� ��������� ����� ��� ��������.</p>
              <p> 
                <input type="text" name="plen" size="10" maxlength="3" value="<%=isFirst?plen:Request.Form("plen")%>">
                ��������� �� ����� ��������</p>
              <%
	}
%>
              <p align="center"> 
                <input type="submit" name="Submit" value="����������">
              </p>
            </form>
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <%
}
%>
    </td>
    <td width="1" valign="top" bgcolor="#123D87"></td>
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
    <td width="12"></td>
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
            <script language="JavaScript">document.write('<a href="http://www.rax.ru/click" target=_blank><img src="http://counter.yadro.ru/hit?t16.1;r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" border=0 width=88 height=31 alt="rax.ru: �������� ����� ����� �� 24 ����, ����������� �� 24 ���� � �� ������\�"></a>')</script>
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
      <p align="center"><a href="/">�� �������</a> <img src="images/dot.gif" width="12" height="9" align="absmiddle"> 
        <a href="message.asp">�������� �����</a> <img src="images/dot.gif" width="12" height="9" align="absmiddle"> 
        <a href="gbook.asp">�������� �����</a></p>
    </td>
    <td bgcolor="#666699" width="1"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" height="12" bgcolor="#E1F4FF">
  <tr> 
    <td width="1"></td>
    <td bgcolor="#FFFFFF" align="center"> 
      <p><font size="1"><font size="1"><font color="#FFFFFF"> </font></font> <a href="area.asp">&copy;</a> 
        2003 ���������� <a href="http://www.rusintel.ru"> ��������</a></font></p>
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
