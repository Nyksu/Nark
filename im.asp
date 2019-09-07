<%@LANGUAGE="JAVASCRIPT"%>
<%
Response.Contenttype="image/jpeg"
Response.Expires=0

var g=Server.CreateObject("shotgraph.image")
var xsize=60
var ysize=20

g.CreateImage(xsize,ysize,4)
g.SetColor(0,255,255,255)
g.SetColor(1,0,0,0)
g.SetBgColor(0)
g.FillRect(0,0,xsize-1,ysize-1)

g.SetBkMode("TRANSPARENT")
g.CreateFont("Arial",204,18,0,true,false,false,false,false)
g.SetTextColor(1)
g.TextOut(4,1,Session("lcode"))

Response.BinaryWrite(g.JpegImage(70,1,""))
%>
