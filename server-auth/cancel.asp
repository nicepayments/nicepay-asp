<%@Language="VBScript" CODEPAGE="65001"%>
<!--#include file="./lib/uuid.asp" -->
<%
orderId = CreateWindowsGUID()
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
  <h1>NICEPAY TEST</h1>
  <form method="POST" action="./cancelResponse.asp">
    <label>tid</label><br>
    <input type="text" name="tid" value=""><br>

    <label>amount</label><br>
    <input type="text" name="amount" value="1004"><br><br>

    <input type="submit" value="취소요청">
  </form> 
</body>

</html>