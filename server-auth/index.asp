<%@Language="VBScript" CODEPAGE="65001"%>
<!--#include file="./lib/uuid.asp" -->
<%
orderId = CreateWindowsGUID()
clientId = "S2_af4543a0be4d49a98122e01ec2059a56"
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
  <button onclick="serverAuth()">serverAuth 결제하기</button>

  <script src="https://pay.nicepay.co.kr/v1/js/"></script>

  <script>
    function serverAuth() {
      AUTHNICE.requestPay({
        clientId: '<%=clientId%>',
        method: 'card',
        orderId: '<%=orderId%>',
        amount: 1004,
        goodsName: '나이스페이-상품',
        returnUrl: 'http://localhost/response.asp',
        fnError: function (result) {
          alert('개발자확인용 : ' + result.errorMsg + '')
        }
      });
    }
  </script>
</body>

</html>