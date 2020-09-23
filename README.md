<div align="center">

## Balloon Tooltips


</div>

### Description

Adds an icon on the systray, where you can send Windows XP / 2000 style balloon messages
 
### More Info
 
You need to know how to call Windows API and also the new model of Visual Basic NET for imports


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rodrigo Sandoval](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rodrigo-sandoval.md)
**Level**          |Intermediate
**User Rating**    |4.9 (59 globes from 12 users)
**Compatibility**  |VB\.NET
**Category**       |[GUIs](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/guis__10-30.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rodrigo-sandoval-balloon-tooltips__10-68/archive/master.zip)





### Source Code

```
'//Before the declaration of the module put the following line...
Imports System.Runtime.InteropServices
'//Now, put this code on the module....
Public Result as Boolean
<StructLayout(LayoutKind.Sequential)> Public Structure NOTIFYICONDATA
 Dim cbSize As Int32
 Dim hwnd As IntPtr
 Dim uID As Int32
 Dim uFlags As Int32
 Dim uCallbackMessage As IntPtr
 Dim hIcon As IntPtr
 <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> Dim szTip As String
 Dim dwState As Int32
 Dim dwStateMask As Int32
 <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Dim szInfo As String
 Dim uVersion As Int32
 <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)> Dim szInfoTitle As String
 Dim dwInfoFlags As Int32
End Structure
Public Const NIF_MESSAGE As Int32 = &H1
Public Const NIF_ICON As Int32 = &H2
Public Const NIF_STATE As Int32 = &H8
Public Const NIF_INFO As Int32 = &H10
Public Const NIF_TIP As Int32 = &H4
Public Const NIM_ADD As Int32 = &H0
Public Const NIM_MODIFY As Int32 = &H1
Public Const NIM_DELETE As Int32 = &H2
Public Const NIM_SETVERSION As Int32 = &H4
Public Const NOTIFYICON_VERSION As Int32 = &H5
Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_INFO = &H1
Public Const NIIF_NONE = &H0
Public Const NIIF_WARNING = &H2
Public Const NIM_SETFOCUS = &H4
Public Const NIIF_GUID = &H5
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" _
Alias "Shell_NotifyIconA" (ByVal dwMessage As Int32, _
ByRef lpData As NOTIFYICONDATA) As Boolean
Public uNIF As NOTIFYICONDATA
'//Before the declaration of the Form Class add the following code
Imports System.Runtime.InteropServices
'//and now in your form put this code in the load event
'// Adds the icon
With uNIF
 .cbSize = Marshal.SizeOf(uNIF)
 .hwnd = Me.Handle
 .uID = 1
 .dwInfoFlags = NIF_ICON Or NIF_MESSAGE
 .uCallbackMessage = New IntPtr(&H500)
 .uVersion = NOTIFYICON_VERSION
 .hIcon = Me.Icon.Handle
End With
Result = Shell_NotifyIcon(NIM_ADD, uNIF)
'// Send a balloon message
With uNIF
 .uFlags = NIF_INFO
 .uVersion = 2000
 .szInfoTitle = "Test"
 .szInfo = "Testing 1,2,3 Testing"
 .dwInfoFlags = NIIF_INFO
End With
Result = Shell_NotifyIcon(NIM_MODIFY, uNIF)
'//if you want to send a balloon message with
'//the error icon just change de dwInfoFlags to
'//NIIF_ERROR, or for a warning NIIF_WARNING and
'//so on, if you want to send messages in other
'//parts of your project just put the code that
'//sends the balloon message in the event,
'//CAUTION: if you don't put the Add icon code
'//in your main form load event you will not be
'//able to receive balloon messages.
'//Please send me your comments to drkmouse@prodigy.net.mx
```

