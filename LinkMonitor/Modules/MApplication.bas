Attribute VB_Name = "MApplication"
'''''''''''''''''''''''''''''
'All properties of App object
Public Function Comments() As String
   Comments = App.Comments
End Function

Public Function CompanyName() As String
   CompanyName = App.CompanyName
End Function

Public Function EXEName() As String
   EXEName = App.EXEName
End Function

Public Function FileDescription() As String
   FileDescription = App.FileDescription
End Function

Public Function HelpFile() As String
   HelpFile = App.HelpFile
End Function

Public Function Instance() As Long
   Instance = App.hInstance
End Function

Public Function LegalCopyright() As String
   LegalCopyright = App.LegalCopyright
End Function

Public Function LegalTrademarks() As String
   LegalTrademarks = App.LegalTrademarks
End Function

Public Function LogMode() As Long
   LogMode = App.LogMode
End Function

Public Function LogPath() As String
   LogPath = App.LogPath
End Function

Public Function Major() As Integer
   Major = App.Major
End Function

Public Function Minor() As Integer
   Minor = App.Minor
End Function

Public Function NonModalAllowed() As Boolean
   NonModalAllowed = App.NonModalAllowed
End Function

Public Function OleRequestPendingMsgText() As String
   OleRequestPendingMsgText = App.OleRequestPendingMsgText
End Function

Public Function OleRequestPendingMsgTitle() As String
   OleRequestPendingMsgTitle = App.OleRequestPendingMsgTitle
End Function

Public Function OleRequestPendingTimeout() As Long
   OleRequestPendingTimeout = App.OleRequestPendingTimeout
End Function

Public Function OleServerBusyMsgText() As String
   OleServerBusyMsgText = App.OleServerBusyMsgText
End Function

Public Function OleServerBusyMsgTitle() As String
   OleServerBusyMsgTitle = App.OleRequestPendingMsgTitle
End Function

Public Function OleServerBusyRaiseError() As Boolean
   OleServerBusyRaiseError = App.OleServerBusyRaiseError
End Function

Public Function OleServerBusyTimeout() As Long
   OleServerBusyTimeout = App.OleServerBusyTimeout
End Function

Public Function Path() As String
   Path = App.Path
End Function

Public Function PrevInstance() As Boolean
   PrevInstance = App.PrevInstance
End Function

Public Function ProductName() As String
   ProductName = App.ProductName
End Function

Public Function RetainedProject() As Boolean
   RetainedProject = App.RetainedProject
End Function

Public Function Revision() As Integer
   Revision = App.Revision
End Function

Public Function StartMode() As Integer
   StartMode = App.StartMode
End Function

Public Function TaskVisible() As Boolean
   TaskVisible = App.TaskVisible
End Function

Public Function ThreadID() As Long
   ThreadID = App.ThreadID
End Function

Public Function Title() As String
   Title = App.Title
End Function

Public Function UnattendedApp() As Boolean
   UnattendedApp = App.UnattendedApp
End Function
'All properties of App object
'''''''''''''''''''''''''''''



'Derived & other properties
Public Function ProductNameVersion() As String
   ProductNameVersion = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Function
