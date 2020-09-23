VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CIO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long


Public Sub LoadComboBox(pComboBox As ComboBox, strFilePath As String)
    'Loads a combobox items from a file
    
    'EXAMPLE:
    'Call goIO.LoadComboBox(cmbCategory, "Categories.dat")
    
    On Local Error Resume Next
    If Dir(strFilePath) <> "" Then
        Open strFilePath For Input As #1
        While Not EOF(1)
            Input #1, Item
            pComboBox.AddItem Item
        Wend
        Close #1
    End If
    frmMain.cmbCategory.ListIndex = 0
End Sub


Public Sub SaveComboBox(pComboBox As ComboBox, strFilePath As String)
    'Saves a combobox items into file.

    'EXAMPLE:
    'Call goIO.SaveComboBox(cmbCategory, "Categories.dat")
    
    On Local Error Resume Next
    Dim i As Integer
    Open strFilePath For Output As #3
    For i = 0 To pComboBox.ListCount - 1
        Item = pComboBox.List(i)
        Write #3, Item
    Next
    Close #3
End Sub

Public Sub WriteFile(pinput As String, pfilePath As String)
   'Saves text into given file path

   'EXAMPLE:
   'Call goIO.WriteFile("FileContent", "c:\file.dat")
   
   If (pfilePath <> "") Then
      Set ofileSystem = CreateObject("Scripting.FileSystemObject")
      Set otextFile = ofileSystem.CreateTextFile(pfilePath, True)
      otextFile.WriteLine (pinput)
      otextFile.Close
   End If
   
End Sub


Public Function VolumeSerialNumber(ByVal RootPath As String) As String
    'Gives serial number of the volume
    
    'EXAMPLE
    'txtSerial.Text = VolumeSerialNumber("C:\")
    
    Dim VolLabel As String
    Dim VolSize As Long
    Dim Serial As Long
    Dim MaxLen As Long
    Dim Flags As Long
    Dim Name As String
    Dim NameSize As Long
    Dim s As String

    If GetVolumeSerialNumber(RootPath, VolLabel, VolSize, Serial, MaxLen, Flags, Name, NameSize) Then
        'Create an 8 character string
        s = Format(Hex(Serial), "00000000")
        'Adds the '-' between the first 4 characters and the last 4 characters
        VolumeSerialNumber = left(s, 4) + "-" + Right(s, 4)
    Else
        'If the call to API function fails the function returns a zero serial number
        VolumeSerialNumber = "0000-0000"
    End If

End Function



