<%
'https://docs.microsoft.com/en-us/troubleshoot/iis/create-guids-asp#create-a-stronger-windows-style-random-guid
'Response.Write "<P>GUID = " & CreateWindowsGUID()

    Function CreateWindowsGUID()
        CreateWindowsGUID = CreateGUID(8) & "-" & _
        CreateGUID(4) & "-" & _
        CreateGUID(4) & "-" & _
        CreateGUID(4) & "-" & _
        CreateGUID(12)
    End Function

    Function CreateGUID(tmpLength)
        Randomize Timer
        Dim tmpCounter,tmpGUID
        Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        For tmpCounter = 1 To tmpLength
        tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
        Next
        CreateGUID = tmpGUID
    End Function
%>