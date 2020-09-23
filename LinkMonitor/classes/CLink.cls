VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CLink
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LinkTextBox(ByRef pTextBox As TextBox, _
                       plinkTopic As String, _
                       plinkMode As LinkModeConstants, _
                       plinkItem As Variant)
   'Links a textbox with a program
   
   'EXAMPLE:
   'Call goCLink.LinkTextBox(txtMonitor, "iexplore|WWW_GetWindowInfo", vbLinkManual, &HFFFFFFFF)
   
   'If could not find link item sustain listen
   On Local Error GoTo listen:
   
   pTextBox.LinkTopic = plinkTopic
   pTextBox.LinkTimeout = 10
   pTextBox.LinkMode = plinkMode
   pTextBox.LinkItem = plinkItem
   pTextBox.LinkRequest
   
listen:
   On Error GoTo 0
End Sub
