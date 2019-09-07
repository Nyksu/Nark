<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=17
// +++  smi_id - код СМИ в таблице SMI !!

var pid=parseInt(Request("pid"))
if (isNaN(pid)) {Response.Redirect("index.asp")}

var hid=0
var sminame=""
var tit=""
var hdd=0
var nm=""
var hiname=""
var period=0
var path=""
var hadr=""
var pname=""
var pdat=""
var autor=""
var news=""
var imgname=""
var imgLname=""
var imgRname=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var sos=""
var ishtml=0
var isnews=1
var lgok=false
var usok=false
var bnm=""
var bpos=""
var bid=0

var tps=parseInt(Request("tps"))
var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
if (isNaN(tps)) {tps=1}
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}

if (String(Session("id_mem"))=="undefined") {
	if (Session("tip_mem_pub")<3) {usok=true}
	if (Session("tip_mem_pub")<4) {lgok=true}
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			lgok=true
		}
		Records.Close()
	} else {
		usok=true
		lgok=true
	}
}


Records.Source="Select t1.* from publication t1, heading t2 where t1.state=1 and t1.id="+pid+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("index.asp")
}
hid=Records("heading_id").Value
pname=String(Records("NAME").Value)
pdat=Records("PUBLIC_DATE").Value
autor=TextFormData(Records("AUTOR").Value,"")
ishtml=TextFormData(Records("ISHTML").Value,"0")
Records.Close()

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

tit=sminame

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	isnews=Records("ISNEWS").Value
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
	}
	path="<a href=\""+hadr+"?hid="+hdd+"\"><font color=\"#FFFFFF\">"+nm+"</font></a> > "+path
	hdd=Records("HI_ID").Value
	Records.Close()
}

var ddt = new Date()
var dt=""
var str=""
var sumdat=Server.CreateObject("datesum.DateSummer")
if ( isnews ) {
str=String(ddt.getMonth()+1)
if (str.length==1) {str="0"+str}
dt="."+str+"."+ddt.getYear()
str=String(ddt.getDate())
if (str.length==1) {str="0"+str}
dt=str+dt
dt=sumdat.SummToDate(dt,"-"+period)
if (sumdat.DateComparing(dt,pdat) > 0) {sos="В архиве"} else {sos="В публикации"}
}

path="<a href=\"index.asp\"><font color=\"#FFFFFF\">"+sminame+"</font></a> > "+path
tit+=" | "+hiname

filnam=PubFilePath+pid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml==0) {
	while (!ts.AtEndOfStream){
		news+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}

imgname=PubImgPath+pid+".gif"
if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
if (imgname=="") {
	imgname=PubImgPath+pid+".jpg"
	if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
}

imgLname=PubImgPath+"l"+pid+".gif"
if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
if (imgLname=="") {
	imgLname=PubImgPath+"l"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
}

imgRname=PubImgPath+"r"+pid+".gif"
if (!fs.FileExists(PubFilePath+"r"+pid+".gif")) { imgRname="" }
if (imgRname=="") {
	imgRname=PubImgPath+"r"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"r"+pid+".jpg")) { imgRname="" }
}

%>

<html>
<head>
<title><%=tit%> > <%=pname%> г. Тюмень</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: "Times New Roman", Times, serif; font-size: 14px; font-style: normal; font-weight: 400}
.link {  font-family: Arial, Helvetica, sans-serif; text-decoration: none}
h1 {  font-family: "Times New Roman", Times, serif; font-size: 18px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-style: normal; font-weight: 400 ; color: #009933; line-height: 18px}
.web { font-family: Arial, Helvetica, sans-serif; color: #0000FF; text-decoration: none ; font-size: 12px}
.nav { font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; color: #006600}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="15" marginwidth="35">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF" height="96">
  <tr> 
    <td width="1" valign="top"></td>
    <td width="12" valign="top">&nbsp;</td>
    <td width="150" valign="middle" height="110" align="center"><a href="/"><img src="images/logo.gif" width="95" height="100" border="0"></a></td>
    <td valign="middle" align="center" height="110"> 
      <h1><b><font face="Times New Roman, Times, serif">УПРАВЛЕНИЕ ГОСНАРКОКОНТРОЛЯ 
        РОССИИ ПО ТЮМЕНСКОЙ ОБЛАСТИ</font></b></h1>
    </td>
    <td width="12">&nbsp;</td>
    <td width="1"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#123D87">
  <tr> 
    <td height="16" width="1"> </td>
    <td bordercolor="#123D87" bgcolor="#666666" height="16" width="12">&nbsp;</td>
    <td bordercolor="#666666" bgcolor="#666666" height="16" width="15" valign="top">&nbsp;</td>
    <td bordercolor="#666666" bgcolor="#666666" height="16" align="left"> 
      <p><font color="#FFFFFF"><%=tit%></font> <img src="http://counter.rambler.ru/top100.cnt?527741" alt="" width=1 height=1 border=0></p>
    </td>
    <form name="form1" method="post" action="search.asp">
    </form>
  </tr>
</table>
<hr>
<p>&nbsp;</p>

<table width="100%" border="0" cellpadding="0" align="center" cellspacing="0" bgcolor="#E1F4FF">
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"> 
      <table width="98%" border="0" bordercolor="#FFFFFF" align="center" cellspacing="0" class="base_text">
        <tr valign="top" bordercolor="#FFFFFF"> 
          <td> 
            <h1 align="center"><%=pname%></h1>
            <p align="right">&copy; <%=pdat%>&nbsp; <%=autor%></p>
            <p align="right">&nbsp;</p>
          </td>
        </tr>
        <tr valign="top" bordercolor="#FFFFFF"> 
          <td height="55"> 
            <p> 
              <%if (imgLname != "") {%>
              <img src="<%=imgLname%>" align="left" border="1" > 
              <%}else{%>
              &nbsp; 
              <%}%>
              <font face="Times New Roman, Times, serif" size="2"><%=news%></font></p>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellspacing="0">
        <tr bordercolor="#CCCCCC"> 
          <td valign="top" height="40"> 
            <div align="center"> 
              <%if (imgRname != "") {%>
              <img src="<%=imgRname%>" border="1" > 
              <%}else{%>
              &nbsp; 
              <%}%>
            </div>
          </td>
        </tr>
      </table>
      <table border="0" cellspacing="2" cellpadding="2" width="100%">
        <tr> 

        </tr>
      </table>
      
      <hr>
    </td>
  </tr>
</table>
</body>
</html>
