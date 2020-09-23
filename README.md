<div align="center">

## \_ Authenticate via HTTP 401, digest calculator \(You have no idea how long this took to work out\! \) \_


</div>

### Description

This code takes into account the NONCE, CNONCE, NONCE-count, Username, Password, HTTP Method, URi, realm......... and finally QOP type :)

THIS CODE IS FOR WEB DEVELOPERS WHO WISH TO AUTHENTICATE VIA HTTP, AND KNOW WHAT ALL OF THE ABOVE ARE... (if you dont, i suggest you stop reading now :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Barker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-barker.md)
**Level**          |Advanced
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-barker-authenticate-via-http-401-digest-calculator-you-have-no-idea-how-long-this-took__1-34533/archive/master.zip)

### API Declarations

```
'IF YOU DO NOT HAVE THE MD5 DLL, GET IT HERE:
'http://www.esquadro.com.br/md5bas.zip
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)
Public Function MD5String(KeyAndPass As String) As String
 Dim r As String * 32, t As Long
 r = Space(32)
 t = Len(KeyAndPass)
 MDStringFix KeyAndPass, t, r
 MD5String = r
End Function
```


### Source Code

```
RESPONSE RECEIVED FROM HTTP SERVER:
HTTP/1.1 401 Authorization Required..Server: Microsoft-IIS/5.0..Date: Tue, 07 May 2002 17:14:49 GMT..P3P:CP="BUS CUR CONo FIN IVDo ONL OUR PHY SAMo TELo"..Connection: close..Content-Type: text/html..WWW-Authenticate: Digest realm="hotmail.com", nonce="MTAyMDc5MTY4OTowZjY5YmE1NjEzNmM5YTE4NGZmNGQ1ZWFkNzU3ZTIxNw==", qop="auth"..X-Dav-Error: 401 Wrong email address....HMServer: H: DAV73 V: WIN2K 09.04.50.0031 i D: Apr 18 2002 12:14:38...
=============================================
TO CALCULATE THE RESPONSE VALUE:
MsgBox (GetResponse("MTAyMDc5MTY4OTowZjY5YmE1NjEzNmM5YTE4NGZmNGQ1ZWFkNzU3ZTIxNw==", "b6327c933ceeb677f8d6056c60aeabcb", "00000001", "auth", "[USERNAME]", "[PASSWORD]", "hotmail.com", "PROPFIND", "/cgi-bin/hmdata"))
=============================================
Function GetResponse(Nonce As String, CNonce As String, NonceCount As String, QOP As String, Username As String, Password As String, Realm As String, Method As String, URi As String)
 Dim Buffer As String
 Dim Buffer2 As String
 Dim Buffer3 As String
 Buffer = MD5String(Username & ":" & Realm & ":" & Password)
 Buffer2 = MD5String(Method & ":" & URi)
 Buffer3 = MD5String(Buffer & ":" & Nonce & ":" & NonceCount & ":" & CNonce & ":" & QOP & ":" & Buffer2)
 GetResponse = Buffer3
End Function
```

