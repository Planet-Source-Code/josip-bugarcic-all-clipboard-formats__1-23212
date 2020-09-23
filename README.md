<div align="center">

## All Clipboard Formats


</div>

### Description

It gives you possibility using all existing clipboard formats in VB. It is designed as ActiveX Control so you can use it in ASPs as well.

Standard implementation of:

- plain text

- rich text format

But also:

- Biff5

- OEM text

- DIF

- UNICODETEXT

- SYLK (The most powerfull)

- CSV
 
### More Info
 
any text as the first pameter and clipboard format as the second

puts text in clipboard in different formats. You can put any text (even two different) in two different formats and decide which one you want to take and where

Probably this will not work for large amount of text altough I didn't find any problems


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Josip Bugarcic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/josip-bugarcic.md)
**Level**          |Advanced
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/josip-bugarcic-all-clipboard-formats__1-23212/archive/master.zip)

### API Declarations

```
Private Type POINTAPI
  x As Long
  y As Long
End Type
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14
Private Const CF_HDROP = 15
Private Const CF_LOCALE = 16
Private Const CF_MAX = 17
Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"
Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
Private Const CFSTR_NETRESOURCES As String = "Net Resource"
Private Const CFSTR_FILEDESCRIPTOR As String = "FileGroupDescriptor"
Private Const CFSTR_FILECONTENTS As String = "FileContents"
Private Const CFSTR_FILENAME As String = "FileName"
Private Const CFSTR_PRINTERGROUP As String = "PrinterFriendlyName"
Private Const CFSTR_FILENAMEMAP As String = "FileNameMap"
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_MODIFY = &H80
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const APINULL = 0
```


### Source Code

```
Public Sub CopyTLP(strText As String, strSylk As String)
Dim wLenT As Integer
Dim hMemoryT As Long
Dim lpMemoryT As Long
Dim wLenS As Integer
Dim hMemoryS As Long
Dim lpMemoryS As Long
Dim retval As Variant
  If OpenClipboard(APINULL) Then
    Call EmptyClipboard
    wLenT = Len(strText) + 1
    strText = strText & vbNullChar
    hMemoryT = GlobalAlloc(GHND, wLenT + 1)
    If hMemoryT Then
      lpMemoryT = GlobalLock(hMemoryT)
      retval = lstrcpy(lpMemoryT, strText)
      Call GlobalUnlock(hMemoryT)
      retval = SetClipboardData(CF_TEXT, hMemoryT)
    End If
    wLenS = Len(strSylk) + 1
    strSylk = strSylk & vbNullChar
    hMemoryS = GlobalAlloc(GHND, wLenS + 1)
    If hMemoryS Then
      lpMemoryS = GlobalLock(hMemoryS)
      retval = lstrcpy(lpMemoryS, strSylk)
      Call GlobalUnlock(hMemoryS)
      retval = SetClipboardData(CF_SYLK, hMemoryS)
    End If
  End If
  Call CloseClipboard
End Sub
Public Sub CopyText(strText As String)
  'ExecuteCopy strText, CF_TEXT
  Clipboard.GetText vbCFText
End Sub
Public Sub CopyRTF(strText As String)
  'ExecuteCopy strText, CF_TEXT
  Clipboard.GetText vbCFRTF
End Sub
Public Sub CopyOEMText(strText As String)
  ExecuteCopy strText, CF_OEMTEXT
End Sub
Public Sub CopyDIF(strText As String)
  ExecuteCopy strText, CF_DIF
End Sub
Public Sub CopyUNICODETEXT(strText As String)
  ExecuteCopy strText, CF_UNICODETEXT
End Sub
Public Sub CopySYLK(strText As String)
  ExecuteCopy strText, CF_SYLK
End Sub
Public Sub CopyXlTable(strText As String)
Dim wCBformat As Long
wCBformat = RegisterClipboardFormat("XlTable")
If wCBformat <> 0 Then
  ExecuteCopy strText, wCBformat
End If
End Sub
Public Sub CopyBiff5(strText As String)
Dim wCBformat As Long
wCBformat = RegisterClipboardFormat("BIFF5")
If wCBformat <> 0 Then
  ExecuteCopy strText, wCBformat
End If
End Sub
Public Sub CopyCsv(strText As String)
Dim wCBformat As Long
wCBformat = RegisterClipboardFormat("Csv")
If wCBformat <> 0 Then
  ExecuteCopy strText, wCBformat
End If
End Sub
Private Sub ExecuteCopy(strText As String, clipFormat As Long)
Dim wLen As Integer
Dim hMemory As Long
Dim lpMemory As Long
Dim retval As Variant
  If OpenClipboard(APINULL) Then
    Call EmptyClipboard
    wLen = Len(strText) + 1
    strText = strText & vbNullChar
    hMemory = GlobalAlloc(GHND, wLen + 1)
    If hMemory Then
      lpMemory = GlobalLock(hMemory)
      'Call CopyMem(ByVal lpMemory, strText, wLen)
      retval = lstrcpy(lpMemory, strText)
      Call GlobalUnlock(hMemory)
       retval = SetClipboardData(clipFormat, hMemory)
    End If
  End If
  Call CloseClipboard
End Sub
Public Function Paste()
Paste = Clipboard.GetText(1)
End Function
Function CanPaste() As Boolean
  If IsClipboardFormatAvailable(CF_TEXT) Then
    CanPaste = True
  ElseIf IsClipboardFormatAvailable(CF_UNICODETEXT) Then
    CanPaste = True
  ElseIf IsClipboardFormatAvailable(CF_OEMTEXT) Then
    CanPaste = True
  ElseIf IsClipboardFormatAvailable(CF_DIF) Then
    CanPaste = True
  End If
End Function
```

