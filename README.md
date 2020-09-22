<div align="center">

## Create URL from Path \(URLCreateFromPath API\)


</div>

### Description

create an url from a file path

for example

input: "E:/my photo.jpg"

output: "file:///E:/my%20photo.jpg"
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alexander Triantafyllou](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alexander-triantafyllou.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alexander-triantafyllou-create-url-from-path-urlcreatefrompath-api__1-54856/archive/master.zip)





### Source Code

```
Private Declare Sub UrlCreateFromPath Lib "shlwapi.dll" Alias "UrlCreateFromPathA" (ByVal pszPath As String, ByVal pszUrl As String, ByRef pcchUrl As Long, ByVal dwFlags As Long)
'create a url from a file path
'for example
'input: "E:/my photo.jpg"
'output: "file:///E:/my%20photo.jpg"
'Alexander Triantafyllou alextriantf@yahoo.gr
'BSc Information Technology & Telecommunications
'University of Athens , Greece
const MAX_PATH=260
Public Function url_encode(ByVal str_urlpath As String) As String
 Dim out_str As String
 Dim str_path As String
 out_str = String(MAX_PATH, 0)
 str_path = str_urlpath + String(100, 0)
 UrlCreateFromPath str_path, out_str, MAX_PATH, 0
 out_str = StripTerminator(out_str)
 url_encode = out_str
End Function
'Remove all trailing Chr$(0)'s
Function StripTerminator(sInput As String) As String
 Dim ZeroPos As Long
 ZeroPos = InStr(1, sInput, Chr$(0))
 If ZeroPos > 0 Then
  StripTerminator = Left$(sInput, ZeroPos - 1)
 Else
  StripTerminator = sInput
 End If
End Function
```

