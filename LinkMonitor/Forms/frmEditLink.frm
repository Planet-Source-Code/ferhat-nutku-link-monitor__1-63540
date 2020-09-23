VERSION 5.00
Begin VB.Form frmEditLink 
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   Icon            =   "frmEditLink.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   6465
   StartUpPosition =   1  'CenterOwner
   Begin LinkMonitor.LabeledTextBox txtPassword 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   450
   End
   Begin LinkMonitor.LabeledTextBox txtUserName 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   450
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2040
      Width           =   900
   End
   Begin VB.CommandButton btnUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2040
      Width           =   900
   End
   Begin LinkMonitor.LabeledComboBox cmbCategory 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
   End
   Begin LinkMonitor.LabeledTextBox txtLink 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
   End
   Begin LinkMonitor.LabeledTextBox txtTitle 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
   End
End
Attribute VB_Name = "frmEditLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private goControl As New CControl
Private goAdo As New CADO
Private goSecurity As New CSecurity

'''''''''''''
'''METHODS'''
'''''''''''''

Private Sub Form_Load()

   'TextBoxes
   txtLink.LabelText = "Link"
   txtTitle.LabelText = "Title"
   txtLink.ReadOnly = True
   txtUserName.LabelText = "User Name"
   txtPassword.LabelText = "Password"
   
   'ComboBoxes
   cmbCategory.LabelText = "Category"
   
   'Form
   Me.Caption = MApplication.ProductNameVersion + " - Edit Link"
End Sub

Private Sub btnCancel_Click()
   Unload Me
End Sub

Private Sub btnUpdate_Click()
   Dim strTitle As String
   Dim strLink As String
   Dim strUserName As String
   Dim strPassword As String
   Dim strCategory As String
   Dim sSql As String
   
   strTitle = goSecurity.RemoveUnSecureChars(txtTitle.Text, True)
   strLink = goSecurity.RemoveUnSecureChars(txtLink.Text, True)
   strUserName = goSecurity.RemoveUnSecureChars(txtUserName.Text, True)
   strPassword = goSecurity.RemoveUnSecureChars(txtPassword.Text, True)
   strCategory = cmbCategory.SeletedValue
      
      
   'Update record
   sSql = "UPDATE Links SET [Title] = '" & strTitle & "', "
   sSql = sSql & "[UserName] = '" & strUserName & "', "
   sSql = sSql & "[Password] = '" & strPassword & "', "
   sSql = sSql & "[Category] = '" & strCategory & "' "
   sSql = sSql & "WHERE Address = '" & strLink & "';"
   
   Call goAdo.ExecuteCommand(frmMain.gDBName, sSql)
      
   
   'Fill or browse datagrid
   If frmMain.goBrowseSql = "" Then
      Call frmMain.FillDataGrid
   Else
      Call frmMain.BrowseDataGrid
   End If
   
   
   'Unload the form
   Unload Me
End Sub
