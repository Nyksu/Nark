<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var admok=false

if (Session("is_adm_mem")==1) {admok=true}
if (Session("is_host")==1) {admok=true}

var name=""
var domen=""
var boss=""
var phone=""
var email=""
var city=""
var uadr=""
var padr=""
var fadr=""
var ind=""
var inn=""
var rsch=""
var ksch=""
var okonh=""
var okpo=""
var bik=""
var bank=""
var divis=0
var admid=0
var admname=""

var tps=parseInt(Request("tps"))
var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
if (isNaN(tps)) {tps=1}
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}

Records.Source="Select t1.*, t2.name as citnam from ENTERPRISE t1, CITY t2 where t1.ID="+company+" and t1.CITY_ID=t2.ID order by NAME"
Records.Open()
if (!Records.EOF) {
	name=String(Records("NAME").Value)
	domen=String(Records("DOMEN").Value)
	if (Records("BOSSFAM").Value!=null) {boss=Records("BOSSFAM").Value}
	if (Records("CITNAM").Value!=null) {city=Records("CITNAM").Value}
	if (Records("ADDRESS").Value!=null) {uadr=Records("ADDRESS").Value}
	if (Records("ADDRESS_F").Value!=null) {fadr=Records("ADDRESS_F").Value}
	if (Records("ADDRESS_P").Value!=null) {padr=Records("ADDRESS_P").Value}
	if (Records("POSTINDEX").Value!=null) {ind=Records("POSTINDEX").Value}
	if (Records("INN").Value!=null) {inn=Records("INN").Value}
	if (Records("KS").Value!=null) {ksch=Records("KS").Value}
	if (Records("RS").Value!=null) {rsch=Records("RS").Value}
	if (Records("OKONH").Value!=null) {okonh=Records("OKONH").Value}
	if (Records("OKPO").Value!=null) {okpo=Records("OKPO").Value}
	if (Records("BIK").Value!=null) {bik=Records("BIK").Value}
	if (Records("BANK").Value!=null) {bank=Records("BANK").Value}
	if (Records("BOSSIO").Value!=null) {boss=boss+" "+Records("BOSSIO").Value}
	phone=String(Records("PHONE").Value)
	if (Records("EMAIL").Value!=null) {email=Records("EMAIL").Value}
	admid=Records("USERS_ID").Value
} else {Response.Redirect("index.asp")}
Records.Close()
%>

<html>
<head>
<title>Реквизиты <%=name%> - город <%=city%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-style: normal; font-weight: 400}
.link {  font-family: Arial, Helvetica, sans-serif; color: #009933; text-decoration: none}
h1 {  color: #009933; font-family: Arial, Helvetica, sans-serif; font-size: 18px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-style: normal; font-weight: 400 ; color: #009933; line-height: 18px}
.web { font-family: Arial, Helvetica, sans-serif; color: #0000FF; text-decoration: none ; font-size: 12px}
.nav { font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; color: #006600}
-->
</style>
</head>
<body bgcolor="#F5FCF5" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/bg.gif">
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td background="images/fon_green.gif" width="175" align="center" bgcolor="#006600"><a href="/"><img src="images/office.jpg" width="165" border="0"></a></td>
    <td width="7" bgcolor="#990000"><img src="images/trans_g_r.gif" width="7" height="180"></td>
    <td background="images/fon_red.gif" valign="middle" align="center" bgcolor="#990000"><img src="images/tit.gif" width="366" height="84" alt="Управление Госнаркоконтроля по Тюменской области"></td>
    <td width="7" bgcolor="#990000"><img src="images/trans_r_g.gif" width="7" height="180"></td>
    <td background="images/fon_green.gif" width="175" align="center" bgcolor="#006600"><a href="/"><img src="images/orel.gif" width="157" height="150" border="0" alt="На главную страницу"></a></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellpadding="0" align="center" cellspacing="0">
  <tr> 
    <td width="1" bgcolor="#006600"></td>
    <td width="12" bgcolor="#FFFFFF"><img src="http://counter.rambler.ru/top100.cnt?527741" alt="" width=1 height=1 border=0></td>
    <td valign="top" bgcolor="#FFFFFF" align="center"> 
      <p align="left"><a href="/"><br>
        На главную</a> &gt; Реквизиты<br>
        <a href="javascript:history.back(1)">Вернуться назад</a> </p>
      <h1><b><%=name%></b></h1>
      <div align="center"> 
        <table width="95%" border="1" bordercolor="#FFFFFF">
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>Руководитель:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%" bgcolor="#FFFFFF"> 
              <p><%=boss%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>Телефон:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%" bgcolor="#FFFFFF"> 
              <p><%=phone%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>Домен:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%" bgcolor="#FFFFFF"> 
              <p><%=domen%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>E-mail:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%" bgcolor="#FFFFFF"> 
              <p><%=email%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b> Юр.адрес:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%" bgcolor="#FFFFFF"> 
              <p><%=uadr%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>Почтовый адрес:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%" bgcolor="#FFFFFF"> 
              <p><%=ind%>, город <%=city%>, <%=padr%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>Адрес офиса:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%" bgcolor="#FFFFFF"> 
              <p><%=fadr%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>ИНН:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td bgcolor="#FFFFFF" width="70%"> 
              <p><%=inn%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#339966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>ОКПО / ОКОНХ:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td bgcolor="#FFFFFF" width="70%"> 
              <p><%=okpo%> / <%=okonh%></p>
            </td>
          </tr>
        </table>
        <h1>Банковские реквизиты: </h1>
        <table width="95%" border="1" bordercolor="#FFFFFF">
          <tr valign="top" bordercolor="#009966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>Банк:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%"> 
              <p><%=bank%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#009966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>БИК:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%"> 
              <p><%=bik%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#009966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>Р/сч.:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%"> 
              <p><%=rsch%></p>
            </td>
          </tr>
          <tr valign="top" bordercolor="#009966"> 
            <td width="30%"> 
              <div align="right"> 
                <p><b>К/сч.:&nbsp;&nbsp;</b></p>
              </div>
            </td>
            <td width="70%"> 
              <p><%=ksch%></p>
            </td>
          </tr>
        </table>
      </div>
      <div align="center"> 
        <p> 
          <% if (admok) {%>
          <font size="2"><b><a href="edcomp.asp">Изменить реквизиты</a></b></font> 
          <%}%>
        </p>
      </div>
    </td>
    <td width="1" valign="top" bgcolor="#006600"></td>
    <td width="1" bgcolor="#123D87"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td background="images/bott.gif" align="center" width="778" height="19" bgcolor="#006600"></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td align="center" height="1" bgcolor="#006600"></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center"></td>
    <td align="center" height="6"> </td>
    <td width="1" align="center"></td>
  </tr>
</table>
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center" bgcolor="#006600"></td>
    <td align="center" width="780" height="1" bgcolor="#006600"></td>
    <td width="1" align="center" bgcolor="#006600"></td>
  </tr>
</table>
<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="1" align="center"></td>
    <td align="center" width="778" height="6"> </td>
    <td width="1" align="center"></td>
  </tr>
</table>

<div style="FONT-FAMILY: 'MS Sans Serif', Geneva, sans-serif; FONT-SIZE: 8px;>
<div align="center"> 
  <div align="center">copyright © Группа общественных связей Управления Госнаркоконтроля 
    РФ по Тюменской области<br>
    Перепечатка материалов возможна только со ссылкой на авторов сайта </div>
</div>
</body>
</html>
