Attribute VB_Name = "Module1"
Sub test()

MsgBox "Hello"
MsgBox "Hello again."
MsgBox "Hello a third time."

Dim strName As String

strName = "Anthony"
MsgBox strName

MsgBox SayHello("Anthony")
MsgBox SayHello("Peter")


End Sub

Function SayHello(strInput As String) As Boolean

On Error GoTo errorhandler

SayHello = False
MsgBox strInput


Endgame:
SayHello = True
Exit Function

errorhandler:
SayHello = False
Exit Function

End Function
