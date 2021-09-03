<%@Language="VBScript" CODEPAGE="65001"%>

<!DOCTYPE html>
<html>

<head>
  <title>Nicepay classic ASP(VBScript)</title>
  <meta charset="UTF-8">
  <meta httpEquiv="x-ua-compatible" content="ie=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
</head>

<body>
  <h1><%= Request.Form("resultMsg") %></h1>
  <p>상세한 응답 body는 log를 확인해주세요</p>

  <br>
  <%
  For Each item In Request.Form
      Response.Write item & " : " & Request.Form(item) & "<BR />"
  Next
  %>
</body>

</html>