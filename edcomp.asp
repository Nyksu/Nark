<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\creaters.inc" -->

<%
if ((Session("is_adm_mem")!=1)&&(Session("is_host")!=1)){
Session("backurl")="edcomp.asp"
Response.Redirect("login.asp")
}

var name=""
var domen=""
var bossio=""
var bossf=""
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
var ErrorMsg=""
var ShowForm=true
var kov="\""
var ids=0
var sql=""
var re = new RegExp("(\\w+)@((\\w+).)*(\\w+)$")


Records.Source="Select t1.*, t2.name as citnam from ENTERPRISE t1, CITY t2 where t1.ID="+company+" and t1.CITY_ID=t2.ID order by NAME"
Records.Open()
if (!Records.EOF) {
	name=String(Records("NAME").Value)
	while (name.indexOf(kov)>=0) {name=name.replace(kov,"&quot;")}
	domen=String(Records("DOMEN").Value)
	if (Records("BOSSFAM").Value!=null) {bossf=Records("BOSSFAM").Value}
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
	while (bank.indexOf(kov)>=0) {bank=bank.replace(kov,"&quot;")}
	if (Records("BOSSIO").Value!=null) {bossio=Records("BOSSIO").Value}
	phone=String(Records("PHONE").Value)
	if (Records("EMAIL").Value!=null) {email=Records("EMAIL").Value}
} else {Response.Redirect("index.asp")}
Records.Close()

isFirst=String(Request.Form("Submit"))=="undefined"

if (!isFirst) {
	city=Request.Form("city")
	name=Request.Form("name")
	uadr=Request.Form("uadr")
	padr=Request.Form("padr")
	fadr=Request.Form("fadr")
	ind=Request.Form("ind")
	email=Request.Form("email")
	phone=Request.Form("phone")
	domen=Request.Form("domen")
	inn=Request.Form("inn")
	okpo=Request.Form("okpo")
	okonh=Request.Form("okonh")
	rsch=Request.Form("rsch")
	bank=Request.Form("bank")
	ksch=Request.Form("ksch")
	bik=Request.Form("bik")
	bossio=Request.Form("bossio")
	bossf=Request.Form("bossf")
	//while (name.indexOf("'")>=0) {name=name.replace("'",kov)}
	//while (bank.indexOf("'")>=0) {bank=bank.replace("'",kov)}
	
	if (name.length < 6) {ErrorMsg+="Слишком короткое наименование компании!<br>"}
	if (city.length <3) {ErrorMsg+="Слишком короткое наименование города!<br>"}
	if ((email!="")&&(!re.test(email))) {ErrorMsg+="Некорректный E-mail.<br>"}
	
	if (ErrorMsg=="") {
		Connect.BeginTrans()
		try{
			Records.Source="Select id from city where name='"+city+"'"
			Records.Open()
			if (!Records.EOF) {
				ids=Records("ID").Value
				Records.Close()
			}
			else {
				Records.Close()
				ids=NextID("cityid")
				Records.Source="Insert into city (id,name) Values("+ids+",'"+city+"')"
				Connect.Execute(sql)
			}
			sql=updatecompany
			sql=sql.replace("%ID",company)
			sql=sql.replace("%CITY",ids)
			sql=sql.replace("%NAME",name)
			sql=sql.replace("%ADR",uadr)
			sql=sql.replace("%ADRP",padr)
			sql=sql.replace("%ADRF",fadr)
			sql=sql.replace("%IND",ind)
			sql=sql.replace("%EML",email)
			sql=sql.replace("%PHON",phone)
			sql=sql.replace("%DOM",domen)
			sql=sql.replace("%INN",inn)
			sql=sql.replace("%OKPO",okpo)
			sql=sql.replace("%OKONH",okonh)
			sql=sql.replace("%RS",rsch)
			sql=sql.replace("%BANK",bank)
			sql=sql.replace("%KS",ksch)
			sql=sql.replace("%BIK",bik)
			sql=sql.replace("%BF",bossf)
			sql=sql.replace("%BIO",bossio)
			Records.Source=sql
			Connect.Execute(sql)
		}
		catch(e){
			Connect.RollbackTrans()
			ErrorMsg+="Ошибка апдейта!<br>"
			ErrorMsg+=ListAdoErrors()
			while (name.indexOf(kov)>=0) {name=name.replace(kov,"&quot;")}
			while (bank.indexOf(kov)>=0) {bank=bank.replace(kov,"&quot;")}
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
<title>Изменение параметров компании.</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На главную страницу</a> 
        | <a href="area.asp">КАБИНЕТ Администратора</a> </font></p>
    </td>
  </tr>
</table>
<h1 align="center"><b>Изменение реквизитов компании.</b></h1>
<p>&nbsp;</p>
<p>
  <%if(ErrorMsg!=""){%>
</p>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%>
<%if(ShowForm){%>
<form name="form1" method="post" action="edcomp.asp">
<div align="center">
    <table width="95%" border="1" bordercolor="#FFFFFF">
      <tr bordercolor="#9999FF"> 
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p><font color="#135086" size="3"><b>Организация</b>:</font>&nbsp;&nbsp;</p>
          </div>
        </td>
        <td width="35%" bordercolor="#333333"> 
          <p><b><font size="3" color="#FF0000"> 
            <input type="text" name="name" maxlength="99" size="35" value="<%=name%>">
            </font></b></p>
        </td>
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>ИНН:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="inn" maxlength="14" size="14" value="<%=inn%>">
            </font></p>
        </td>
      </tr>
      <tr bordercolor="#9999FF"> 
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>Телефон:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td width="35%" bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="phone" maxlength="20" size="20" value="<%=phone%>">
            </font></p>
        </td>
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>ОКПО / ОКОНХ:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="okpo" maxlength="10" size="10" value="<%=okpo%>">
            / 
            <input type="text" name="okonh" maxlength="8" size="8" value="<%=okonh%>">
            </font></p>
        </td>
      </tr>
      <tr bordercolor="#9999FF"> 
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>Домен:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td width="35%" bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="domen" maxlength="30" size="35" value="<%=domen%>">
            </font></p>
        </td>
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>Банк:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="bank" maxlength="99" size="35" value="<%=bank%>">
            </font></p>
        </td>
      </tr>
      <tr bordercolor="#9999FF"> 
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>E-mail:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td width="35%" bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="email" maxlength="40" size="35" value="<%=email%>">
            </font></p>
        </td>
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>БИК:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="bik" maxlength="10" size="10" value="<%=bik%>">
            </font></p>
        </td>
      </tr>
      <tr bordercolor="#9999FF"> 
        <td width="15%" bordercolor="#333333" height="2" bgcolor="#CCCCCC"> 
          <div align="right"> 
            <p> Юр.адрес:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td width="35%" height="2" bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="uadr" size="35" maxlength="99" value="<%=uadr%>">
            </font></p>
        </td>
        <td width="15%" height="2" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>Р/сч.:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td height="2" bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="rsch" maxlength="20" size="20" value="<%=rsch%>">
            </font></p>
        </td>
      </tr>
      <tr bordercolor="#9999FF"> 
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>Почтовый адрес:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td width="35%" bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="ind" maxlength="6" size="8" value="<%=ind%>">
            , </font><font size="3" color="#666666">город </font><font color="#0000CC"> 
            <input type="text" name="city" size="20" maxlength="99" value="<%=city%>">
            , 
            <input type="text" name="padr" maxlength="99" size="35" value="<%=padr%>">
            </font></p>
        </td>
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333"> 
          <div align="right"> 
            <p>К/сч.:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td bordercolor="#333333"> 
          <p><font color="#0000CC"> 
            <input type="text" name="ksch" maxlength="20" size="20" value="<%=ksch%>">
            </font></p>
        </td>
      </tr>
      <tr bordercolor="#9999FF"> 
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333" height="35"> 
          <div align="right"> 
            <p>Факт. адрес:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td width="35%" bordercolor="#333333" height="35"> 
          <p><font color="#0000CC"> 
            <input type="text" name="fadr" maxlength="99" size="35" value="<%=fadr%>">
            </font></p>
        </td>
        <td width="15%" bgcolor="#CCCCCC" bordercolor="#333333" height="35"> 
          <div align="right"> 
            <p>Руководитель:&nbsp;&nbsp;</p>
          </div>
        </td>
        <td height="35" bordercolor="#333333"> 
          <p><font color="#0000CC"> Имя Отчество 
            <input type="text" name="bossio" maxlength="40" size="30" value="<%=bossio%>">
            <br>
            Фамилия 
            <input type="text" name="bossf" maxlength="25" size="25" value="<%=bossf%>">
            </font></p>
        </td>
      </tr>
    </table>
    <p><br>
      <input type="submit" name="Submit" value="Сохранить">
    </p>
    <p>&nbsp; </p>
    <hr noshade>
  </div>
</form>
<%} else {%>
<center>
  <h2><font color="#3333FF">Параметры компании обновлены!</font></h2>
</center>
<%}%>
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
<p align="center"><font size="1"><b>| </b><a href="index.asp"><font size="1">В 
  начало</font></a><b> | </b><a href="message.asp">Обратная связь</a><b> | </b><a href="org.asp">Реквизиты</a><b> |</b></font></p>
<p align="center"><font size="1">&copy; 2003 программирование <a href="http://www.rusintel.ru">Русинтел</a></font></p>
  <p align="center">&nbsp;</p>
  </body>
</html>
