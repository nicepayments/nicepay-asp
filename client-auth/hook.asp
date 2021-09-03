<%@Language="VBScript" CODEPAGE="65001"%>

<!--#include file="./lib/aspJSON1.19.asp"-->
<%

Set responseJSON = New aspJSON

If Request.TotalBytes > 0 Then
  Dim lngBytesCount
  lngBytesCount = Request.TotalBytes
  responseJSON.loadJSON(BytesToStr(Request.BinaryRead(lngBytesCount)))
End If

If responseJSON.data("resultCode") = "0000" Then
  '비즈니스 로직 구현
  Response.write "ok"
End If

Function BytesToStr(bytes)
  Dim Stream
  Set Stream = Server.CreateObject("Adodb.Stream")
      Stream.Type = 1
      Stream.Open
      Stream.Write bytes
      Stream.Position = 0
      Stream.Type = 2
      Stream.Charset = "iso-8859-1"
      BytesToStr = Stream.ReadText
      Stream.Close
  Set Stream = Nothing
End Function
%>