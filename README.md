<div align="center">

## \_Like Ping


</div>

### Description

Like (Ping) IP address
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Petko Petkov](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/petko-petkov.md)
**Level**          |Intermediate
**User Rating**    |4.9 (112 globes from 23 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/petko-petkov-like-ping__1-48045/archive/master.zip)





### Source Code

<p>Private Type QOCINFO<br>
 &nbsp;&nbsp;dwSize As Long<br>
 &nbsp;&nbsp;dwFlags As Long<br>
 &nbsp;&nbsp;dwInSpeed As Long 'in bytes/second<br>
 &nbsp;&nbsp;dwOutSpeed As Long 'in bytes/second<br>
 End Type</p>
<p><br>
 Private Declare Function IsDestinationReachable Lib &quot;SENSAPI.DLL&quot;
 Alias &quot;IsDestinationReachableA&quot; (ByVal lpszDestination As String,
 ByRef lpQOCInfo As QOCINFO) As Long<br>
</p>
<p>Private Sub Form_Load()<br>
 &nbsp;&nbsp;Dim Ret As QOCINFO<br>
 &nbsp;&nbsp;Dim IP As String<br>
 &nbsp;&nbsp;Ret.dwSize = Len(Ret)<br>
 &nbsp;&nbsp;'Put desired IP<br>
 &nbsp;&nbsp;IP = &quot;217.9.238.114&quot;<br>
 &nbsp;&nbsp;If IsDestinationReachable(IP, Ret) = 0 Then<br>
 &nbsp;&nbsp;&nbsp;&nbsp;MsgBox &quot;The destination cannot be reached!&quot;<br>
 &nbsp;&nbsp;Else<br>
 &nbsp;&nbsp;&nbsp;&nbsp;MsgBox &quot;The destination can be reached!&quot; +
 vbCrLf + _<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&quot;The speed of data coming in from the destination
 is &quot; + Format$(Ret.dwInSpeed / 1048576, &quot;#.0&quot;) + &quot; Mb/s,&quot;
 + vbCrLf + _<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&quot;and the speed of data sent to the destination
 is &quot; + Format$(Ret.dwOutSpeed / 1048576, &quot;#.0&quot;) + &quot; Mb/s.&quot;<br>
 &nbsp;&nbsp;End If<br>
 End Sub </p>
<p></p>

