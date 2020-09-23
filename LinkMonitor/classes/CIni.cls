VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''INI CALLS'''
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

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

Public Function INISaveSetting(ByVal sFileName As String, ByVal sSection As String, ByVal sKey As String, ByVal vntValue As Variant)
' Will save an INI Setting to the specified Section and Key in the INI file
' secified by the full path name in sFileName

   #If Win32 Then
     Dim xRet          As Long
   #Else
     Dim xRet          As Integer
   #End If
   
   xRet = WritePrivateProfileString(sSection, sKey, CStr(vntValue), sFileName)

End Function


Public Function INIGetSetting(ByVal sFileName As String, ByVal sSection As String, ByVal sKey As String) As Variant
' Will return an INI entry in the specified section at the specified key in the INI file
' specified by the full path name in sFilename

   #If Win32 Then
     Dim xRet          As Long
   #Else
     Dim xRet          As Integer
   #End If

  Dim sReturnStr    As String
  Dim nStringLen    As Integer
  
  nStringLen = 255
  sReturnStr = String(nStringLen, Chr$(0))  ' Buffer String
  
  xRet = GetPrivateProfileString(sSection, sKey, "", sReturnStr, nStringLen, sFileName)

  INIGetSetting = Left(sReturnStr, xRet)
  
End Function


Public Sub INIDeleteSetting(ByVal sFileName As String, ByVal sSection As String, Optional vntKey As Variant)
' If vntKey is specified it this will delete the entry specified by vntKey, if not
' it will delete the entire section sepecified by sSection in the INI specefied by
' sFilename

   #If Win32 Then
     Dim xRet          As Long
   #Else
     Dim xRet          As Integer
   #End If



  ' If key was provided just delete that key and value, if not delete the
  ' entire section
  If IsMissing(vntKey) Then
     xRet = WritePrivateProfileString(sSection, 0&, 0&, sFileName)
  Else
     xRet = WritePrivateProfileString(sSection, CStr(vntKey), 0&, sFileName)
  End If

End Sub


Public Function INIGetAllSetting(ByVal sFileName As String, ByVal sSection As String) As Variant
' Returns an variant array of all keys(0) and values(1) same as GetAllSettings
' This is  the complicated one.    It reads all of the Key Names into a temporary array
' then after teh array has been read it will crate another array.  The new array is
' 2 dimensional, the first dimension is the pair number.   The second dimension
' is 0 for the keyname, 1 for the value.

   #If Win32 Then
     Dim xRet          As Long
   #Else
     Dim xRet          As Integer
   #End If

  Dim sReturnStr    As String
  Dim nStringLen    As Integer
  Dim nEndOfKey     As Integer
  Dim nNumKeys      As Integer
  Dim arrValues()   As Variant
  
  nStringLen = 5000        ' Must be big enough to hold all keys
  sReturnStr = String(nStringLen, Chr$(0))
  nNumKeys = -1
  xRet = GetPrivateProfileString(sSection, 0&, "", sReturnStr, nStringLen, sFileName)
  
  ' Parse the string, and add the elements to the array
  Do While (InStr(sReturnStr, Chr$(0)) > 1)
     ' Get each key in the section
     nEndOfKey = InStr(sReturnStr, Chr$(0))
     nNumKeys = nNumKeys + 1
     ReDim Preserve arrValues(nNumKeys)
     arrValues(nNumKeys) = Left$(sReturnStr, nEndOfKey - 1)
     sReturnStr = Mid(sReturnStr, nEndOfKey + 1)
  Loop

  If nNumKeys = -1 Then
     ' if no keys return an empty variant
     INIGetAllSetting = Empty
  Else
     ' Get the values for each key and return that, to maintain compliance with
     ' GetAllSettings
     ReDim arrFullArray(0 To nNumKeys, 0 To 1) As Variant
     For nNumKeys = LBound(arrValues) To UBound(arrValues)
         arrFullArray(nNumKeys, 0) = arrValues(nNumKeys)
         arrFullArray(nNumKeys, 1) = INIGetSetting(sFileName, sSection, arrValues(nNumKeys))
     Next nNumKeys
     INIGetAllSetting = arrFullArray
  End If
End Function

