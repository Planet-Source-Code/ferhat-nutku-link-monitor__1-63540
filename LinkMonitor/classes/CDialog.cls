VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''
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
Public Function OpenSaveDialog(pCommonDialog As CommonDialog, _
                  pstrDialogTitle As String, _
                  pstrFilter As String, _
                  pstrExtention As String) As String
   
   'Opens an SaveDialog box.
   'Drag & Drop a CommonDialog Control
   'OpenSaveDialog(commondialog1,"Save Window","Bitmap Dosyasý|*.BMP","bmp")

   pCommonDialog.Color = ColorConstants.vbRed
   pCommonDialog.DialogTitle = pstrDialogTitle
   pCommonDialog.Filter = pstrFilter
   pCommonDialog.DefaultExt = pstrExtention
   'pObj.Action = 2 'also true
   pCommonDialog.ShowSave 'save dialog
      
   OpenSaveDialog = pCommonDialog.FileName
End Function


Public Function OpenOpenDialog(pCommonDialog As CommonDialog, _
                  pstrDialogTitle As String, _
                  pstrFilter As String, _
                  pstrExtention As String) As String
   
   'Opens an OpenDialog box.
   'Drag & Drop a CommonDialog Control
   'OpenOpenDialog(commondialog1,"Open Window","Bitmap Dosyasý|*.BMP","bmp")
  
   pCommonDialog.Color = ColorConstants.vbBlue
   pCommonDialog.DefaultExt = pstrExtention
   pCommonDialog.DialogTitle = pstrDialogTitle
   pCommonDialog.Filter = pstrFilter
   'pCommonDialog.InitDir = pstrInitialDir
   'pCommonDialog.Action = 1 'also true
   pCommonDialog.ShowOpen 'open dialog
          
   OpenOpenDialog = pCommonDialog.FileName
End Function


Public Sub OpenPopupDialog(pForm As Form, pMenu As Object)
   
   'Opens a popup menu
   'EXAMPLE:
   'Call goDialog.OpenPopupDialog(Me, mnGridPopup)

   Call pForm.PopupMenu(pMenu, vbPopupMenuLeftAlign + vbPopupMenuRightButton)
End Sub


