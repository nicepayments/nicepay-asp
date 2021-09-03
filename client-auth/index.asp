﻿<%@Language="VBScript" CODEPAGE="65001"%>
<!--#include file="./lib/uuid.asp" -->
<%
orderId = CreateWindowsGUID()
clientId = "클라이언트 키"
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
  <button onclick="clientAuth()">clientAuth 결제하기</button>

  <script src="https://pay.nicepay.co.kr/v1/js/pay/"></script> <!--nicepay payment-window client-auth-->

  <script>
    function clientAuth() {
      PAYNICE.requestPay({
        clientId: '<%=clientId%>',
        method: 'card',
        orderId: '<%=orderId%>',
        amount: 1004,
        goodsName: '나이스페이-상품',
        returnUrl: 'http://localhost:80/response.asp',
        fnError: function (result) {
          alert('고객용메시지 : ' + result.msg + '\n개발자확인용 : ' + result.errorMsg + '')
        }
      });
    }
  </script>
</body>

</html>