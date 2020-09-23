VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'APIs
'*******************************************************************
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'*******************************************************************


'Types
'*******************************************************************
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'*******************************************************************


'Constants
'*******************************************************************
Private Const NI_ADD = &H0
Private Const NI_MODIFY = &H1
Private Const NI_DELETE = &H2
Private Const NI_MESSAGE = &H1
Private Const NI_ICON = &H2
Private Const NI_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
'*******************************************************************


'Enumerations
'*******************************************************************
Public Enum EnumMButtonConstants
   WM_LBUTTONDBLCLK = &H203
   WM_LBUTTONDOWN = &H201
   WM_LBUTTONUP = &H202
   
   WM_RBUTTONDBLCLK = &H206
   WM_RBUTTONDOWN = &H204
   WM_RBUTTONUP = &H205
End Enum
'*******************************************************************


'Events
'*******************************************************************
'Public Event LButtonDblClk()
'Public Event LButtonDown()
'Public Event LButtonUp()
'Public Event RButtonDblClk()
'Public Event RButtonDown()
'Public Event RButtonUp()
'*******************************************************************


'Variables
'*******************************************************************
Private TrayI As NOTIFYICONDATA
'*******************************************************************


'Properties
'*******************************************************************
Public Property Get Tip() As String
  
  Tip = TrayI.szTip
  
End Property

Public Property Let Tip(ByVal Value As String)

  TrayI.szTip = Value & Chr(0)
  
End Property
'*******************************************************************



'Class Initialize
'********************************************************************************
Private Sub Class_Initialize()

   TrayI.cbSize = Len(TrayI) 'Notify Icon boyutu
   TrayI.uId = 1&
   TrayI.uFlags = NI_ICON Or NI_TIP Or NI_MESSAGE 'Response flagS
   TrayI.ucallbackMessage = WM_MOUSEMOVE 'Link event. Now Form_Mouse is belongs to here.
   
    
End Sub



'Subs
'*******************************************************************
Public Sub GiveHandle(pHandleNo As Long)
   
   TrayI.hWnd = pHandleNo
   
End Sub


Public Sub AddTrayIcon(ByRef pPicture As IPictureDisp)
   'Adds the Icon to system tray
   
   'Example
   'Call goTray.AddTrayIcon(LoadPicture(MApplication.Path & "\" & MResourse.rSysTrayIcon))
   
   TrayI.hIcon = pPicture
   Call Shell_NotifyIcon(NI_ADD, TrayI)
  
End Sub


Public Sub RemoveTrayIcon()
   'Deletes the Icon from system tray
   
   'Example
   'Call goTray.RemoveTrayIcon
   
   Call Shell_NotifyIcon(NI_DELETE, TrayI)
    
End Sub


Public Sub ChangeTrayIcon(ByRef pPicture As IPictureDisp)
   'Changes the Icon in the system tray
   
   'Example
   'Call goTray.ChangeTrayIcon(LoadPicture(MApplication.Path & "\" & MResourse.rSysTrayDisIcon))
   
   TrayI.hIcon = pPicture
   Call Shell_NotifyIcon(NI_MODIFY, TrayI)
  
End Sub


