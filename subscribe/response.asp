<%@Language="VBScript" CODEPAGE="65001"%>
<!--#include file="./lib/base64.asp"-->
<!--#include file="./lib/aspJSON1.19.asp"-->
<!--#include file="./lib/aes256cbc.asp"-->
<!--#include file="./lib/uuid.asp" -->
<%

dim clientId, secretKey, plaintext

clientId = "S2_af4543a0be4d49a98122e01ec2059a56"
secretKey = "9eb85607103646da9f9c02b128f2e5ee"

Set requestJSON = New aspJSON
Set responseJSON = New aspJSON

Set plaintext = CreateObject("System.Collections.ArrayList")
plaintext.Add "cardNo=" + Request.Form("cardNo")
plaintext.Add "&expYear=" + Request.Form("expYear")
plaintext.Add "&expMonth=" + Request.Form("expMonth")
plaintext.Add "&idNo=" + Request.Form("idNo")
plaintext.Add "&cardPw=" + Request.Form("cardPw")

With requestJSON.data
  .Add "encData", Encrypt(join(plaintext.ToArray(), ""), secretKey, secretKey)
  .Add "orderId", CreateWindowsGUID()
  .Add "encMode", "A2"
End With

set req = Server.CreateObject("MSXML2.ServerXMLHTTP")
req.open "POST", "https://sandbox-api.nicepay.co.kr/v1/subscribe/regist", false
req.setRequestHeader "Authorization", "Basic " & Base64Encode(clientId & ":" & secretKey)
req.setRequestHeader "Content-Type", "application/json; charset=utf-8"
req.send requestJSON.JSONoutput()

If req.status = 200 Then
  responseJSON.loadJSON(req.responseText)
  If responseJSON.data("resultCode") = "0000" Then
    '비즈니스 로직 구현
    Billing(responseJSON.data("bid"))
    'expire(responseJSON.data("bid"))

  End If
Else
	Response.write "[ERROR] " & req.status
End If
  set req = Nothing


Function Billing(bid)
  set req = Server.CreateObject("MSXML2.ServerXMLHTTP")
  Set requestJSON = New aspJSON
  Set responseJSON = New aspJSON

  With requestJSON.data
      .Add "orderId", CreateWindowsGUID()
      .Add "amount", 1004
      .Add "goodsName", "card billing test"
      .Add "cardQuota", 0
      .Add "useShopInterest", false
  End With

  req.open "POST", "https://sandbox-api.nicepay.co.kr/v1/subscribe/"+bid+"/payments", false
  req.setRequestHeader "Authorization", "Basic " & Base64Encode(clientId & ":" & secretKey)
  req.setRequestHeader "Content-Type", "application/json; charset=utf-8"
  req.send requestJSON.JSONoutput()

  If req.status = 200 Then
    responseJSON.loadJSON(req.responseText)
    If responseJSON.data("resultCode") = "0000" Then
      '결제 비즈니스 로직 구현
      Response.Write responseJSON.JSONoutput()
    End If
  Else
    Response.write "[ERROR] " & req.status
  End If
    set req = Nothing
End Function


Function expire(bid)
  set req = Server.CreateObject("MSXML2.ServerXMLHTTP")
  Set requestJSON = New aspJSON
  Set responseJSON = New aspJSON

  With requestJSON.data
      .Add "orderId", CreateWindowsGUID()
  End With

  req.open "POST", "https://sandbox-api.nicepay.co.kr/v1/subscribe/"+bid+"/expire", false
  req.setRequestHeader "Authorization", "Basic " & Base64Encode(clientId & ":" & secretKey)
  req.setRequestHeader "Content-Type", "application/json; charset=utf-8"
  req.send requestJSON.JSONoutput()

  If req.status = 200 Then
    responseJSON.loadJSON(req.responseText)
    If responseJSON.data("resultCode") = "0000" Then
      '비즈니스 로직 구현
      Response.Write responseJSON.JSONoutput()
    End If
  Else
    Response.write "[ERROR] " & req.status
  End If
    set req = Nothing
End Function

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