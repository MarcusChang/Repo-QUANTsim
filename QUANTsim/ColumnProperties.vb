'USEUNIT TestUtilities
Option Explicit

' Open the column properties dialog box
Sub Open
  Dim MenuToolbar
  Dim MenuToolbarPopup
  Set MenuToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.MenuBar
  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
    
  Call OptionViewGrid.Click(480, 217)
  Call OptionViewGrid.Keys("~vc")

  Delay(100)
End Sub

Sub OpenDockingViewColumnProperties 
  Dim DockingView
  Set DockingView = Aliases.MarketView.wndAfx.DockingMachineView
  
  If DockingView.Visible = True Then
    Call DockingView.ClickR(10, 25)    
    Delay(100)
  Else
    Log.Warning("DockingView was not visible, cannot open column properties")
  End If
End Sub
  
' Sets the number of decimal places in Column Properties
Sub SetDecimals(ColumnName,DecimalsValue)
  Dim dlgColumnProperties
  Set dlgColumnProperties = Aliases.MarketView.dlgColumnProperties
  
  ' Get a reference to the relevant highlighted tab control   
  Dim dlgColumnPropertiesTab
  Select Case dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)
  Case "Calls"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Calls
  Case "Puts"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Puts
  Case Else
    Log.Error("SetDecimals : the selected tab """&dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)&""" does not have a Decimals list box")
    Exit Sub
  End Select

  ' Only click on the ColumnName if it is available in the selected fields list box
  If ItemInList(dlgColumnPropertiesTab.SelectedFields.wItemList,dlgColumnPropertiesTab.SelectedFields.wListSeparator,ColumnName) Then
    Call dlgColumnPropertiesTab.SelectedFields.ClickItem(ColumnName)
  Else
    Log.Error("SetDecimals : column """&ColumnName&""" was not found in Selected fields list box")
    Exit Sub
  End If
     
  ' In case the the value passed in is not a string this converts it to one 
  DecimalsValue = ""&DecimalsValue&"" 
    
  Call dlgColumnPropertiesTab.Decimals.ClickItem(DecimalsValue)
End Sub
  
' Sets the value of multiplier in Column Properties
Sub SetMultiplier(ColumnName,MultiplierValue)
  Dim dlgColumnProperties
  Set dlgColumnProperties = Aliases.MarketView.dlgColumnProperties
  
  ' Get a reference to the relevant highlighted tab control   
  Dim dlgColumnPropertiesTab
  Select Case dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)
  Case "Calls"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Calls
  Case "Puts"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Puts
  Case Else
    Log.Error("SetMultiplier : the selected tab """&dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)&""" does not have a Multiplier list box")
    Exit Sub
  End Select

  ' Only click on the ColumnName if it is available in the selected fields list box
  If ItemInList(dlgColumnPropertiesTab.SelectedFields.wItemList,dlgColumnPropertiesTab.SelectedFields.wListSeparator,ColumnName) Then
    Call dlgColumnPropertiesTab.SelectedFields.ClickItem(ColumnName)
  Else
    Log.Error("SetMultiplier : column """&ColumnName&""" was not found in Selected fields list box")
  End If
     
  ' In case the the value passed in is not a string this converts it to one 
  MultiplierValue = ""&MultiplierValue&""
  
  Call dlgColumnPropertiesTab.Multiplier.ClickItem(MultiplierValue)
End Sub
  
' Select a tab on the column properties form
Sub SelectTab(Tab)
  Dim dlgColumnProperties
  Set dlgColumnProperties = Aliases.MarketView.dlgColumnProperties
  
  ' Search for the specified tab name in the list of tab names in the form
  Dim i
  Dim Found
  Found = False
  For i = 0 To dlgColumnProperties.TabControl.wTabCount - 1
    If dlgColumnProperties.TabControl.wTabCaption(i) = Tab Then
    Found = True
    Exit For
    End If
  Next
  
  If Found Then
    Call dlgColumnProperties.TabControl.ClickTab(Tab)
  Else
    Log.Error("SelectTab : the tab """&Tab&""" was not available to select")
  End If
End Sub
  
' Check the value for Category is available in the column properties dialog
Public Sub SelectCategory(Category)
  Dim dlgColumnProperties
  Set dlgColumnProperties = Aliases.MarketView.dlgColumnProperties

  ' Find out which of the Calls or Puts tabs is selected and get a reference to the GUI object
  Dim dlgColumnPropertiesTab
  Select Case dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)
  Case "Calls"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Calls
  Case "Puts"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Puts
  Case "Machines"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Machines
  Case Else
    Log.Error("SelectCategory : the selected tab """&dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)&""" does not have Category list box")
    Exit Sub
  End Select
  
  Call dlgColumnPropertiesTab.Category.ClickItem(Category) 
End Sub

Public Sub EnableColumn(CategoryName, FieldName)
  '04/07/11 Looks like everything in Column Properties needs to be remapped
  'Enable a column by specifying a Category and Field
  Dim dlgColumnProperties
  Set dlgColumnProperties = Aliases.MarketView.dlgColumnProperties
    
       ' Get a reference to the relevant highlighted tab control   
  Dim dlgColumnPropertiesTab
  Select Case dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)
  Case "Calls"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Calls
  Case "Puts"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Puts
  Case "Machines"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Machines
  Case Else
    Log.Error("EnableColumn: the selected tab """&dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)&""" does not have columns list box controls")
    Exit Sub
  End Select
  
    'Select the Category by from the ComboListBox
  dlgColumnPropertiesTab.Category.ClickItem(CategoryName)
  
  'Select the field required - Taken from original code, decide later if it's good to reuse
  'If the column name already appears in the selected items list, then we can exit the sub
  If ItemInList(dlgColumnPropertiesTab.SelectedFields.wItemList,dlgColumnPropertiesTab.SelectedFields.wListSeparator,FieldName) Then
    Exit Sub
  End If
  
  ' Otherwise, see if the column is available to select and if it is double click on it to move it to selected items
  If ItemInList(dlgColumnPropertiesTab.AvailableFields.wItemList,dlgColumnPropertiesTab.AvailableFields.wListSeparator,FieldName) Then
    Call dlgColumnPropertiesTab.AvailableFields.DblClickItem(FieldName)
  Else
    Log.Error("EnableColumn : the column """&FieldName&""" was not available to enable")
  End If   
      
End Sub


' This enables a column in column properties
' The column properties dialog must be open, and either the Calls or Puts tab selected
Public Sub EnableColumn_ORG(ColumnName)
  Dim dlgColumnProperties
  Set dlgColumnProperties = Aliases.MarketView.dlgColumnProperties
   
  ' Get a reference to the relevant highlighted tab control   
  Dim dlgColumnPropertiesTab
  Select Case dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)
  Case "Calls"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Calls
  Case "Puts"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Puts
  Case "Machines"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Machines
  Case Else
    Log.Error("EnableColumn: the selected tab """&dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)&""" does not have columns list box controls")
    Exit Sub
  End Select
  
  ' If the column name already appears in the selected items list, then we can exit the sub
  If ItemInList(dlgColumnPropertiesTab.SelectedFields.wItemList,dlgColumnPropertiesTab.SelectedFields.wListSeparator,ColumnName) Then
    Exit Sub
  End If
  
  ' Otherwise, see if the column is available to select and if it is double click on it to move it to selected items
  If ItemInList(dlgColumnPropertiesTab.AvailableFields.wItemList,dlgColumnPropertiesTab.AvailableFields.wListSeparator,ColumnName) Then
    Call dlgColumnPropertiesTab.AvailableFields.DblClickItem(ColumnName)
  Else
    Log.Error("EnableColumn : the column """&ColumnName&""" was not available to enable")
  End If   
End Sub
  
' This disables a column in column properties
' The column properties dialog must be open, and either the Calls or Puts tab selected
Public Sub DisableColumn(ColumnName)
  Dim dlgColumnProperties
  Set dlgColumnProperties = Aliases.MarketView.dlgColumnProperties
   
  ' Get a reference to the relevant highlighted tab control   
  Dim dlgColumnPropertiesTab
  Select Case dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)
  Case "Calls"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Calls
  Case "Puts"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Puts
  Case "Machines"
    Set dlgColumnPropertiesTab = dlgColumnProperties.TabControl.Machines
  Case Else
    Log.Error("EnableColumn: the selected tab """&dlgColumnProperties.TabControl.wTabCaption(dlgColumnProperties.TabControl.wFocusedTab)&""" does not have columns list box controls")
    Exit Sub
  End Select
  
  ' If the column name already appears in the selected items list, then we can exit the sub
  If ItemInList(dlgColumnPropertiesTab.SelectedFields.wItemList,dlgColumnPropertiesTab.SelectedFields.wListSeparator,ColumnName) Then
    Call dlgColumnPropertiesTab.SelectedFields.DblClickItem(ColumnName)
  Else
    Log.Error("EnableColumn : the column """&ColumnName&""" was not available to disable")
  End If   
End Sub
  
' Gets a list of all the selected columns
Public Function GetSelectedColumns()
   Dim SelectedColumns
   
   SelectTab("Calls")
   
   Dim dlgColumnPropertiesTab
   Set dlgColumnPropertiesTab = Aliases.MarketView.dlgColumnProperties.TabControl.Calls
   
   SelectedColumns = dlgColumnPropertiesTab.SelectedFields.wItemList
   
   SelectTab("Puts")
   Set dlgColumnPropertiesTab = Aliases.MarketView.dlgColumnProperties.TabControl.Puts
   
   SelectedColumns = SelectedColumns&dlgColumnPropertiesTab.SelectedFields.wListSeparator&dlgColumnPropertiesTab.SelectedFields.wItemList
   
   GetSelectedColumns = SelectedColumns
End Function
  
' Close the column properties dialog box
Sub Close
  Dim dlgColumnProperties
  Set dlgColumnProperties = Aliases.MarketView.dlgColumnProperties
  dlgColumnProperties.btnApply.ClickButton
  dlgColumnProperties.btnOK.ClickButton
End Sub
