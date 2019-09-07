<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\email.inc" -->


<%
var ErrorMsg=""
var Text=""
var name=""

var tps=parseInt(Request("tps"))
var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
if (isNaN(tps)) {tps=1}
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}

Records.Source="Select * from enterprise where ID="+company
Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("index.asp")
   }
// Здесь в переменной namcomp наименование фирмы. Можно где-нить использовать если надо...
namcomp=String(Records("NAME").Value)
Records.Close()

var eml=Server.CreateObject("JMail.Message")
eml.Logging=true
eml.From=fromaddres
eml.AddRecipient(Recipient)
eml.Subject=Subject
eml.Charset=characterset
eml.ContentTransferEncoding = "base64"

var isSending=false

isFirst=String(Request.Form("Submit"))=="undefined"
ShowForm=true
if(!isFirst){
	
		//-------------input validation-----------
		fio=TextFormData(Request.Form("Name"),"")
		Company=TextFormData(Request.Form("company"),"")
		Phone=TextFormData(Request.Form("Phone"),"")
		Email=TextFormData(Request.Form("email"),"")
		Text=TextFormData(Request.Form("txt"),"")


		if(Text.length>2000){ErrorMsg+="Сообщение превышает допустимый размер.<br>"}
		if(Text.length<4){ErrorMsg+="Сообщение отсутствует.<br>"}
		if(fio ==""){ErrorMsg+="Поле 'Имя' должно быть заполнено.<br>"}
		if ((Email != "") && (!/(\w+)@((\w+).)*(\w+)$/.test(Email))) {ErrorMsg=ErrorMsg+"Неверный формат поля 'E-mail'.<br>"}
		

		if (ErrorMsg==""){
			
			try{
				var  pop3=Server.CreateObject("JMail.POP3")
				pop3.Connect(fromaddres,pswsmtp,servsmtp)
				pop3.Disconnect()
				eml.FromName=fromname+fio
				eml.AppendText("Сообщение с интернет сайта nark.72rus.ru \n")
				eml.AppendText(" Ф.И.О. : "+fio+"\n")
				eml.AppendText(" Компания : "+Company+"\n")
				eml.AppendText(" Телефон : "+Phone+"\n")
				eml.AppendText(" Email : "+Email+"\n")
				eml.AppendText("\n Сообщение : \n \n")
				eml.AppendText(TextFormData(Request.Form("txt")))
				isSending=eml.Send(servsmtp)
	 			if (isSending) {ShowForm=false}
			}
			catch(e){
				ErrorMsg+="Проблемы с почтой.<br>"
			}
		} else {
			
			
		}

	
}

%>



<Html>
<Head>
<Title>Управление госнаркоконтроля России по Тюменской области. Форма почтового сообщения.</Title>
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
</Head>

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
<table width="780" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td width="1" bgcolor="#006600"></td>
    <td valign="top" bgcolor="#F9F9F9" width="15"><img src="http://counter.rambler.ru/top100.cnt?527741" alt="" width=1 height=1 border=0></td>
    <td valign="top" bgcolor="#F9F9F9"> 
      <p align="left"><a href="/"> <br>
        На главную</a> &gt; Адрес доверия <br>
        <a href="javascript:history.back(1)">Вернуться назад</a></p>
      <p align="center"><b>Заполнив форму вы можете отправить свое сообщение или 
        вопрос <br>
        в <%=namcomp%><br>
        либо по электронной почте: <a href="mailto:nark@tmn.ru">nark@tmn.ru</a></b></p>
      <%if(ErrorMsg!=""){%>
      <center>
        <h2><font color="#FF3300">Ошибка!</font> <br>
          <%=ErrorMsg%></h2>
      </center>
      <%}%>
      <form name="Guest" method="post" action="message.asp">
        <%if(ShowForm){%>
        <table width="550" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#006600">
          <tr valign="middle" bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <td width="160" align="right" height="10"></td>
            <td height="10"></td>
          </tr>
          <tr valign="middle" bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <td width="160" align="right" height="30"> 
              <p>Ваше имя: </p>
            </td>
            <td height="30"> 
              <input type="text" name="Name" value="<%=isFirst?"":Request.Form("name")%>" size="50" maxlength="50">
            </td>
          </tr>
          <tr valign="middle" bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <td width="160" align="right" height="32"> 
              <p>Организация:</p>
            </td>
            <td height="32"> 
              <input type="text" name="company" value="<%=isFirst?"":Request.Form("company")%>" size="50" maxlength="100">
            </td>
          </tr>
          <tr valign="middle" bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <td width="160" align="right" height="32"> 
              <p>Контактный телефон:</p>
            </td>
            <td height="32"> 
              <input type="text" name="Phone" value="<%=isFirst?"":Request.Form("phone")%>" size="50" maxlength="50">
            </td>
          </tr>
          <tr valign="middle" bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <td width="160" align="right" height="32"> 
              <p>Ваш e-mail:</p>
            </td>
            <td height="32"> 
              <input type="text" name="email" value="<%=isFirst?"":Request.Form("email")%>" size="50" maxlength="50">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <td width="160" align="right" valign="top" height="144"> 
              <p>Текст сообщения:</p>
            </td>
            <td height="144"> 
              <textarea name="txt" cols="40" rows="8"><%=Text%></textarea>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
            <td width="160" align="right" valign="top" height="32">&nbsp;</td>
            <td height="32"> 
              <input type="submit" name="Submit" value="Переслать">
              <input type="reset" name="Clering" value="Очистить">
            </td>
          </tr>
        </table>
        <%}
else{%>
        <center>
          <p>Спасибо! Ваше сообщение отправлено.</p>
          <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На 
            главную страницу</a></font></p>
        </center>
        <%}%>
      </form>
    </td>
    <td width="1" bgcolor="#006600"></td>
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
</Body>
</Html>
