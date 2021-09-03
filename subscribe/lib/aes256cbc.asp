<%
'August 2021 --Version 1.0.0 by kanghokim wecv@nicepay.co.kr
'Encryption: AES-256-CBC 
'Output Text Format: hex

Function BinToHex(data)
    With Server.CreateObject("MSXML2.DomDocument").CreateElement("b64")
        .dataType = "bin.hex"
        .nodeTypedValue = data
        BinToHex = .text
    End With
End Function

Function Encrypt(plaintext, key, iv)
    Set aes = CreateObject("System.Security.Cryptography.RijndaelManaged")
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    aes.Key = utf8.GetBytes_4(LEFT(key, 32))
    aes.IV = utf8.GetBytes_4(LEFT(key, 16))
    Set aesEnc = aes.CreateEncryptor_2(aes.Key, aes.IV)
    plainBytes = utf8.GetBytes_4(plaintext)
    cipherBytes = aesEnc.TransformFinalBlock((plainBytes), 0, LenB(plainBytes))
    
    Encrypt = BinToHex(cipherBytes)
End Function

%>