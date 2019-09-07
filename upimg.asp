<%@LANGUAGE="JScript"%>
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
   
<%
if ((Session("is_adm_mem")!=1)&&(Session("is_host")!=1)){
Response.Redirect("area.asp")
}

var imgname=""
var ext=""
var size=0
var ErrorMsg=""

if (Request.ServerVariables("REQUEST_METHOD")=="POST") {
	updown = Server.CreateObject("ANUPLOAD.OBJ")
	gd=updown.Form("gd")
	if (!isNaN(parseInt(gd))) { 
		updown.SavePath = GoodsPath
		size=parseInt(updown.GetSize("file"))
		ext=updown.GetExtension("file").toUpperCase()
		updown.Delete(GoodsImgPath+gd+".jpg")
		updown.Delete(GoodsImgPath+gd+".gif")
		rr=updown.SaveAs("file",GoodsPath+gd+"."+ext)
		if (rr==0){Response.Redirect("goods.asp?gd="+gd)}
		else {Response.Redirect("addgdimg.asp?gd="+gd+"&er=error")}
	} else {Response.Redirect("area.asp")} 
} else {Response.Redirect("area.asp")}
%>