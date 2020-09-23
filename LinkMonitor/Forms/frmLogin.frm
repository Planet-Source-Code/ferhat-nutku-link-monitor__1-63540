VERSION 5.00
Begin VB.Form frmLogin 
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Authenticaton Control"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Password"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "User Name"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private strVolume As String
Private strVolumeSerial As String

Private goIni As New CIni
Private goIO As New CIO
Private goSecurity As New CSecurity

Private Sub Form_Load()
   Dim userName As String
   Dim password As String
      
   'Login control variables
   userName = goIni.INIGetSetting(App.Path & "\" & con_INI_File, "Main", "UserName")
   password = goIni.INIGetSetting(App.Path & "\" & con_INI_File, "Main", "Password")
   strVolume = left(App.Path, 3)
   strVolumeSerial = goIO.VolumeSerialNumber(strVolume)
       
   'Form
   Me.Caption = MApplication.ProductNameVersion + " - Authentication Dialog"

   'Control password
   If goSecurity.ControlPassword(strVolumeSerial, userName, password) Then
      
      'Correct Password !
      Me.Hide
      frmMain.Show
   Else
     
     'Incorrect Password !
      'Me.Show (vbModal)
      
      'Remove this for Authentication
      Me.Hide
      frmMain.Show
   End If
   
End Sub


Private Sub btnOK_Click()
   Dim userName  As String
   Dim password As String
   
   userName = Trim(txtUserName.Text)
   password = Trim(txtPassword.Text)
      
   'Username and password should have been at least 9 character long in order to comply with volume serial technique
   If Len(userName) < 9 Or Len(password) < 9 Then
      Call MsgBox("Username and password should have been at least 9 character long.", vbCritical + vbOKOnly, MApplication.ProductNameVersion + " - Login error")
      Exit Sub
   End If
  
  
   'Control password
   If goSecurity.ControlPassword(strVolumeSerial, userName, password) Then
      
      'Correct Password !
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Main", "UserName", userName)
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Main", "Password", password)
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "DB Location", "dbpath", MApplication.Path & "\" & MResourse.con_DB_Name)
      Me.Hide
      
      'Unload Me
      frmMain.Show
   Else
   
      'Incorrect Password !
      Call MsgBox("Please entered the correct password.", vbCritical + vbOKOnly, MApplication.ProductNameVersion + " - Login error")
      End
      
   End If
   
End Sub


Private Sub btnCancel_Click()
   'End program
   End
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'End program
   End
End Sub

