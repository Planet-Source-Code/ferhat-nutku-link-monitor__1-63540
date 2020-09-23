VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function ControlPassword(pVolumeSerial As String, pUserName As String, pPassword As String) As Boolean
   Dim i As Integer
   Dim volume, userName, password As String
   On Local Error GoTo exitfunction
   
   'Reverse volume serial
   pVolumeSerial = StrReverse(pVolumeSerial)
   
   For i = 1 To Len(pVolumeSerial)
      volume = Asc(Mid(pVolumeSerial, i, 1))
      userName = Asc(Mid(pUserName, i, 1))
      password = password & (5 * volume - 2 * userName - 2) 'ALGORITHM CAN BE CHANGED
   Next
   
   If StrComp(pPassword, password, vbBinaryCompare) <> 0 Then
exitfunction:
      'Incorrect password !
      ControlPassword = False
   Else
      'Correct password !
      ControlPassword = True
   End If
   
   'Reverse volume serial to its original value
   pVolumeSerial = StrReverse(pVolumeSerial)

End Function


Public Function RemoveUnSecureChars(ByVal pinput As String, ptrim As Boolean) As String
   
   Dim res As String
    
   'Trim input or not
   If (ptrim) Then
      res = Trim(pinput)
   Else
      res = pinput
   End If
   
   
   'Replace unsecure char with empty string
   res = Replace(res, "`", "", 1, -1, 1)
   'res = Replace(res, "=", "", 1, -1, 1)
   'res = Replace(res, "&", "", 1, -1, 1)
   res = Replace(res, "%", "", 1, -1, 1)
   'res = Replace(res, "!", "", 1, -1, 1)
   'res = Replace(res, "#", "", 1, -1, 1)
   'res = Replace(res, "<", "", 1, -1, 1)
   'res = Replace(res, ">", "", 1, -1, 1)
   'res = Replace(res, "*", "", 1, -1, 1)
   'res = Replace(res, "/", "", 1, -1, 1)
   'res = Replace(res, "\", "", 1, -1, 1)
   'res = Replace(res, "And", "", 1, -1, 1)
   res = Replace(res, "'", "", 1, -1, 1)
   res = Replace(res, """", "", 1, -1, 1)
   'res = Replace(res, "Chr(34)", "", 1, -1, 1)
   'res = Replace(res, "Chr(39)", "", 1, -1, 1)
   
   RemoveUnSecureChars = res
End Function
