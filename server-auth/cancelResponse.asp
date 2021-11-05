<%@Language="VBScript" CODEPAGE="65001"%>
<!--#include file="./lib/uuid.asp" -->
<!--#include file="./lib/base64.asp"-->
<!--#include file="./lib/aspJSON1.19.asp"-->
<%

dim clientId, secretKey

clientId = "S2_af4543a0be4d49a98122e01ec2059a56"
secretKey = "9eb85607103646da9f9c02b128f2e5ee"

Set requestJSON = New aspJSON
Set responseJSON = New aspJSON

With requestJSON.data
  .Add "amount", Request.Form("amount")
  .Add "reason", "test"
  .Add "orderId", CreateWindowsGUID()
End With

set req = Server.CreateObject("MSXML2.ServerXMLHTTP")
req.open "POST", "https://sandbox-api.nicepay.co.kr/v1/payments/" & Request.Form("tid") & "/cancel", false
req.setRequestHeader "Authorization", "Basic " & Base64Encode(clientId & ":" & secretKey)
req.setRequestHeader "Content-Type", "application/json; charset=utf-8"
req.send requestJSON.JSONoutput()

If req.status = 200 Then
  responseJSON.loadJSON(req.responseText)
  If responseJSON.data("resultCode") = "0000" Then
    '결제 비즈니스 로직 구현
  End If
Else
	Response.write "[ERROR] " & req.status
End If
  set req = Nothing

%>

<!DOCTYPE html>
<html>

<head>
  <title>Nicepay classic ASP(VBScript)</title>
  <meta charset="UTF-8">
  <meta httpEquiv="x-ua-compatible" content="ie=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
</head>

<body>
  <h1><%= responseJSON.data("resultMsg") %></h1>
  <p>상세한 응답 body는 log를 확인해주세요</p>
</body>

<hr>
<% Response.Write responseJSON.JSONoutput()%>

</html>