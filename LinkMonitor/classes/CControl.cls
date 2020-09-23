VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CControl
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private goAdo As New CADO


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
   
End Sub


Private Sub SetMSHFlexGridHeaderText(ByRef pflexgrid As MSHFlexGrid)
   'Sets the text of the Header Columns of MSHFlexGrid
   'Call SetMSHFlexGridHeaderText(fgridAddresses)
    
   Dim sHeaderText(7) As String
   sHeaderText(0) = rTitle
   sHeaderText(1) = rAddress
   sHeaderText(2) = rCategory
   sHeaderText(3) = rUserName
   sHeaderText(4) = rPassword
   sHeaderText(5) = rLastDate
   sHeaderText(6) = rFirstDate
   sHeaderText(7) = rVisitCount
   For i = 0 To UBound(sHeaderText)
       pflexgrid.TextMatrix(0, i) = sHeaderText(i)
   Next
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''
'  MSHFlexGrid Functions  '
'''''''''''''''''''''''''''

Public Sub FillMSHFlexGrid(ByRef pflexgrid As MSHFlexGrid, ByVal pdbname As String, ByRef pRecordSet As Recordset, ByVal psqlcmd As String)
   Set pRecordSet = goAdo.GetRecordSetByFileName(pdbname, psqlcmd)
   Set pflexgrid.DataSource = pRecordSet
   Call SetMSHFlexGridHeaderText(pflexgrid)
End Sub

Public Sub SetFocusToMSHFlexGridHeader(ByRef pflexgrid As MSHFlexGrid)
   'Sets focus to MSHFlexGrid header column which the mouse is over
   
   'EXAMPLE:
   'Call goControl.SortMSHFlexGrid(flexgrid)
   
   pflexgrid.Row = 0
   pflexgrid.Col = pflexgrid.MouseCol
   pflexgrid.SetFocus
End Sub

Public Sub SortMSHFlexGrid(ByRef pflexgrid As MSHFlexGrid)
   'Sorts MSHFlexGrid alternatively.
   'Memo type cannot be sorted in fieldname consept.
   'Use this function in order to sort memo types in Access dbs.
   
   'EXAMPLE:
   'Call goControl.SortMSHFlexGrid(flexgrid)
   
   'Above can be used only sort for a fixed field.
   '(It sorts the grid according to first column & than others.)
   
   Static i As Integer
   
   Select Case i
      Case 0
            pflexgrid.Sort = flexSortGenericAscending
            i = 1
      Case 1
            pflexgrid.Sort = flexSortGenericDescending
            i = 0
   End Select
   
   'Set the text of the Header Columns of MSHFlexGrid
   Call SetMSHFlexGridHeaderText(pflexgrid)
End Sub

Public Sub SortMSHFlexGridByField(ByRef pflexgrid As MSHFlexGrid, ByVal pfieldName As String)
   'Sorts MSHFlexGrid alternatively according to a given field (also a column of the grid).
   'Memo type cannot be sorted in fieldname consept.
   
   'EXAMPLE:
   'Call goControl.SortMSHFlexGridByField(flexgrid, sfieldName)
   
   'Define a static variable in the sub scobe in order to remember the previous sorting type.
   Static isortType As Integer
   Dim oRecordSet As Recordset

   'Get a recordset object from the data of the grid.
   Set oRecordSet = pflexgrid.DataSource

   'Change the sorting type.
   Select Case isortType
      Case 0
         oRecordSet.Sort = pfieldName & " ASC"
         isortType = 1
      Case 1
         oRecordSet.Sort = pfieldName & " DESC"
         isortType = 0
   End Select
   
   'Set the datasource of the grid.
   'Set pflexgrid.DataSource = oRecordSet
      
   'Set the text of the Header Columns of MSHFlexGrid
   Call SetMSHFlexGridHeaderText(pflexgrid)
End Sub

'''''''''''''''''''''''''''
'  ComboBox Functions     '
'''''''''''''''''''''''''''

Public Sub FillComboBox(ByRef pComboBox As ComboBox, ByVal pdbname As String, ByRef pRecordSet As Recordset, ByVal psqlcmd As String)
   'Adds an item to a combobox if the item has not already included in the combobox.
   
   'EXAMPLE:
   'sSql = "SELECT DISTINCT Category FROM Links"
   'Call goControl.FillComboBox(cmbBrowseCat, gDBName, goRecordset, sSql)
   
   Set pRecordSet = goAdo.GetRecordSetByFileName(pdbname, psqlcmd)
   Do While Not pRecordSet.EOF
      pComboBox.AddItem (pRecordSet!Category)
      pRecordSet.MoveNext
   Loop
End Sub

Public Sub CopyComboBox(ByVal pSource As ComboBox, ByRef pTarget As ComboBox)
   'Copies all items of a combobox to another combobox
   
   'EXAMPLE:
   'Call cmbCategory.FillComboBoxFromComboBox(cmbCategory)
   
   Dim size As Integer
   size = pSource.ListCount - 1
   
   For i = 0 To size
      pTarget.AddItem (pSource.List(i))
   Next
   
   'Select the same item with the source combox
   pTarget.ListIndex = pSource.ListIndex
End Sub

Public Sub SelectValueComboBox(ByRef pSource As ComboBox, ByVal pitem As String)
   'Selects the combobox item whose text is equal to pitem
   
   'EXAMPLE:
   'Call
   
   Dim size As Integer
   size = pSource.ListCount - 1
   
   For i = 0 To size
      If (pSource.List(i) = pitem) Then
         pSource.ListIndex = i
      End If
   Next
End Sub

Public Sub AddItemToComboBoxUnique(ByRef pComboBox As ComboBox, pstrItem As String)
   'Adds an item to a combobox if the item has not already included in the combobox.
   
   'EXAMPLE:
   'Call goControl.AddItemToComboBoxUnique(cmbBrowseAddress, LCase(cmbBrowseAddress.Text))
   
   'Control if combobox has an item equal to pstrItem
   For i = 0 To pComboBox.ListCount - 1
      If LCase(pComboBox.List(i)) = LCase(pstrItem) Then
         Exit Sub
      End If
   Next
   
   'Add item into the combobox
   pComboBox.AddItem pstrItem
End Sub

Public Sub RemoveItemFromComboBox(ByRef pComboBox As ComboBox, pstrItem As String)
   'Removes an item from a combobox whose text is given.
   
   'EXAMPLE:
   'Call goControl.RemoveItemFromComboBox(cmbBrowseAddress, LCase(cmbBrowseAddress.Text))
   
   'Control if combobox has an item equal to pstrItem
   'Trim is omitted
   For i = 0 To pComboBox.ListCount - 1
      If LCase(pComboBox.List(i)) = LCase(pstrItem) Then
         pComboBox.RemoveItem (i)
      End If
   Next
End Sub


'''''''''''''''''''''''''
'  StatusBar Functions  '
'''''''''''''''''''''''''

Public Sub AddPanelsToStatusBar(ByRef pStatusBar As StatusBar, ptotalPanel)
   'Adds panels into a statusbar control
   
   'EXAMPLE:
   'Call goControl.AddPanelsToStatusBar(sbarRowCount, 4)
   
   For i = 1 To ptotalPanel
      pStatusBar.Panels.Add (i)
   Next
End Sub



'''''''''''''''''''''''''''
'  ProgressBar Functions  '
'''''''''''''''''''''''''''

Public Sub ChangeProgressBarValue(ByRef pProgressBar As ProgressBar, ByVal pmaxValue As Double, ByVal pmaxCount As Double, ByRef pcurrent As Double)
   'Moves progress bar mark
   
   'EXAMPLE:
   '

   'Move progress bar.
   pProgressBar.Value = pmaxValue * pcurrent / pmaxCount
   pcurrent = pcurrent + 1
End Sub


'''''''''''''''''''''''''''
'  General Functions      '
'''''''''''''''''''''''''''


Public Sub CreateFlipFlop(pbtnFlip As CommandButton, pbtnFlop As CommandButton)
   'Modifies a button control pair as Flip Flop buttons

   'EXAMPLE:
   'Call goControl.CreateFlipFlop(btnEncryptDB, btnDecryptDB)
   

   If (pbtnFlip.Enabled) Then
      pbtnFlip.Enabled = False
      pbtnFlop.Enabled = True
   Else
      pbtnFlip.Enabled = True
      pbtnFlop.Enabled = False
   End If

End Sub







