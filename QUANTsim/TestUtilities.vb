Option Explicit

Dim TestConfig
Set TestConfig = ProjectSuite.Variables

Private GridControl
Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl


' Search for a given item in a supplied ItemList delimited by Separator
Const VALUE_UNDEF = 999999999998
Function ItemInList(ItemList, Separator, Item)
  Dim i
  Dim ItemListArray
  ItemListArray = Split(ItemList, Separator)
  
  ' As soon we have found the item in the list set return value to true and exit the function
  ItemInList = False 
  For Each i In ItemListArray
    If i = Item Then
      ItemInList = True
      Exit For
    End If
  Next 
End Function

Function NewItems(ItemListOld, ItemListNew)
  Dim ItemOld,ItemNew,Found,NewItemArray
  Redim NewItemArray(0)
  For Each ItemNew In ItemListNew
    Found = False
    For Each ItemOld In ItemListOld
      If ItemNew = ItemOld Then
        Found = True
        Exit For
      End If
    Next
    If Not Found Then
      NewItemArray(UBound(NewItemArray)) = ItemNew
      Redim Preserve NewItemArray(UBound(NewItemArray)+1)
    End If
  Next
  Redim Preserve NewItemArray(UBound(NewItemArray)-1)
  NewItems = NewItemArray
End Function

Function ListCount(ItemList, Separator)
  Dim ItemListArray
  
  ' Split the list with the separator
  ItemListArray = Split(ItemList, Separator)
  
  ' Return the index of the last item + 1 for the number of items in the list
  ListCount = UBound(ItemListArray) + 1
End Function

Function GetNthField(ItemList, Separator, n)
  Dim ItemListArray
  
  ' Split the list with the separator
  ItemListArray = Split(ItemList, Separator)
  
  If n >= 0 And n <= UBound(ItemListArray) Then
    GetNthField = ItemListArray(n)
  Else
    Log.Error("GetNthField : supplied value n is outside of the boundary of the list")
  End If
End Function

' Put a log in the test result for a number check
Public Function CheckValue(Description, ExpectedResult, ActualResult)
  If GetVarType(ExpectedResult) <> GetVarType(ActualResult) Then
    Log.Error("Type of ExpectedResult ("&GetVarType(ExpectedResult)&") does not match type of actual result ("&GetVarType(ActualResult)&")")
    Exit Function
  End If
  
  If GetVarType(ExpectedResult) = varBoolean And GetVarType(ActualResult) = varBoolean Then
    Dim ExpStr, ActStr
    
    If ExpectedResult = True Then
      ExpStr = "True"
    Else
      ExpStr = "False"
    End If
    
    If ActualResult = True Then
      ActStr = "True"
    Else
      ActStr = "False"
    End If
    
    If ExpectedResult = ActualResult Then
      'Log.Message(Description&" passed - Expected Result = """&ExpStr&""" Actual Result = """&ActStr&"""")
      Log.Checkpoint("[Pass] " &Description&" - Expected Result: """&ExpStr&""" Actual Result: """&ActStr&"""")
      CheckValue = True
    Else
      Log.Error("[Fail] " &Description&" - Expected Result: """&ExpStr&""" Actual Result: """&ActStr&"""")
      CheckValue = False
    End If
    
    Exit Function
  End If
  
  'If ""&ExpectedResult&"" = ""&ActualResult&"" Then
  If ExpectedResult = ActualResult Then
    'Log.Message(Description&" passed - Expected Result = """&ExpectedResult&""" Actual Result = """&ActualResult&"""")
    Log.Checkpoint("[Pass] " &Description&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")
    CheckValue = True
  Else
    Log.Error("[Fail] " &Description&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")
    CheckValue = False
  End If
End Function

Function IsFloatEquals(ExpectedResult,ActualResult)
 If (GetVarType(ExpectedResult) = varDouble Or GetVarType(ExpectedResult) = varSingle) And (GetVarType(ActualResult) = varDouble Or GetVarType(ActualResult) = varSingle) Then 
   If ActualResult > ExpectedResult - 0.0000000001 And ActualResult < ExpectedResult + 0.0000000001 Then
     IsFloatEquals = True
   Else
     IsFloatEquals = False
   End If   
 Else
 End If
End Function

Function CheckFloat(Description, ExpectedResult, ActualResult)
  If (GetVarType(ExpectedResult) = varDouble Or GetVarType(ExpectedResult) = varSingle) And (GetVarType(ActualResult) = varDouble Or GetVarType(ActualResult) = varSingle) Then
    If ActualResult > ExpectedResult - 0.0000000001 And ActualResult < ExpectedResult + 0.0000000001 Then
      Log.Checkpoint("[Pass] " &Description&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")
      CheckFloat = True
   Else
     Log.Error("[Fail] " &Description&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")
     CheckFloat = False
   End If
  Else
    Log.Error("Var type of results not correct GetVarType(ExpectedResult) = "&GetVarType(ExpectedResult)&", GetVarType(ActualResult) = "&GetVarType(ActualResult))
  End If
End Function

Public Function CheckType(Description, ExpectedResult, ActualResult)
  If GetVarType(ExpectedResult) = GetVarType(ActualResult) Then
    Log.Message(Description&" passed - Expected Result = """&GetTypeName(ExpectedResult)&""" Actual Result = """&GetTypeName(ActualResult)&"""")
    CheckType = True
  Else
    Log.Error(Description&" failed - Expected Result = """&GetTypeName(ExpectedResult)&""" Actual Result = """&GetTypeName(ActualResult)&"""")
    CheckType = False
  End If
End Function

' Put a log in the test result for a number check
Public Function CheckValueList(Description, ExpectedResultList, ActualResult)
  Dim ExpectedResult
  Dim Found
  Found = False
  For Each ExpectedResult In Split(ExpectedResultList,"|")
    If ""&ActualResult&"" = ""&ExpectedResult&"" Then
      Found = True
      Exit For
    End If  
  Next
  If Found = True Then
    Log.Message(Description&" passed - Expected Result = """&Replace(ExpectedResultList,"|",", or ")&""" Actual Result = """&ActualResult&"""")
    CheckValueList = True
  Else
    Log.Error(Description&" failed - Expected Result = """&Replace(ExpectedResultList,"|",", or ")&""" Actual Result = """&ActualResult&"""")
    CheckValueList = False
  End If
End Function

Public Function CheckRegExp(Description, Pattern, ActualResult)
  Dim RegExpObj
  Set RegExpObj = New RegExp
    
  RegExpObj.Pattern = Pattern
  
  If RegExpObj.Test(ActualResult) Then
    Log.Message(Description&" passed - Pattern = """&Pattern&""" Actual Result = """&ActualResult&"""")
    CheckRegExp = True
  Else
    Log.Error(Description&" failed - Pattern = """&Pattern&""" Actual Result = """&ActualResult&"""")
    CheckRegExp = False
  End If
End Function

Public Function RegExpMatch(Pattern, Result)
  Dim RegExpObj
  Set RegExpObj = New RegExp
    
  RegExpObj.Pattern = Pattern
  
  If RegExpObj.Test(Result) Then
    RegExpMatch = True
  Else
    RegExpMatch = False
  End If
End Function

' Replace regular expression metacharacters with their escaped character codes
Public Function RegexEscapeReplace(StringName)
  Dim Temp
  Temp = StringName
  Temp = Replace(Temp,".","\.")
  Temp = Replace(Temp,"[","\[")
  Temp = Replace(Temp,"]","\]")
  RegexEscapeReplace = Temp
End Function

Public Sub SelectItemInListBox(ListBox, Item)
  Dim SelectedIndex
  Dim ComboBoxListArray
  SelectedIndex     = ListBox.wSelectedItem
  ComboBoxListArray = Split(ListBox.wItemList, vbCrLf)
    
  ' Work out where in the list the option we want is
  Dim RequiredIndex
  RequiredIndex = 0
  Dim c
  For Each c In ComboBoxListArray
    If c = Item Then
      Exit For
    End If
    RequiredIndex=RequiredIndex+1         
  Next
    
  Dim i
  ' Use the cursor keys to select the value we want
  If RequiredIndex-SelectedIndex > 0 Then
    For i = 0 To (RequiredIndex-SelectedIndex)-1
      Call ListBox.Keys("[Down]")
    Next
  ElseIf  SelectedIndex-RequiredIndex > 0 Then
    For i = 0 To (SelectedIndex-RequiredIndex)-1
      Call ListBox.Keys("[Up]")
    Next
  End If
End Sub

' Round up to nearest number of significance
' E.g. Number = 350.1567
'      Significance = 0.05
'      Floor = 350.2
Function Ceiling(Number, Significance)
  ' If we are already on the significant number boundary, return that
  If FMod(Number,Significance) = 0 Then
    Ceiling = Number
  ' Otherwise round up to nearest number of significance
  Else
    Ceiling = (Fix(Number/Significance)+1)*Significance
  End If
  
End Function

Function FMod(Number, Significance)
  Dim Temp,TempNew,i
  On Error Resume Next
  Temp = ""&Number/Significance&""
  TempNew = ""
  For i = 1 To Len(Temp)
    If Mid(Temp,i,1) = "." Then
      Exit For
    Else
     TempNew = TempNew & Mid(Temp,i,1) 
    End If
  Next
  FMod = Round((Number-(Significance*StrToFloat(TempNew)))*1e9,0)/1e9
End Function

' Round down to nearest number of significance
' E.g. Number = 350.1567
'      Significance = 0.05
'      Floor = 350.15
Function Floor(Number, Significance)
  ' If we are already on the significant number boundary, return that
  If FMod(Number,Significance) = 0 Then
    Floor = Number
  ' Otherwise round down to nearest number of significance
  Else
    Floor = (Fix(Number/Significance))*Significance
  End If
End Function

Function MRound(Number, Significance)
  Log.Message("MRound Number = "&Number&" Significance = "&Significance)
  MRound = (Round(Number/Significance,0))*Significance
End Function

Function FormatDecimals(Number,Decimals)
  Dim i
  FormatDecimals = ""&FormatNumber(Number,Decimals,-1,0,0)&""
  For i = Len(FormatDecimals) To 0 Step -1
    If Mid(FormatDecimals,i,1) = "0" Then
      FormatDecimals = Mid(FormatDecimals, 1,Len(FormatDecimals)-1)
    Else 
      Exit For 
    End If
  Next
  If Mid(FormatDecimals,Len(FormatDecimals),1) = "." Then
      FormatDecimals = Mid(FormatDecimals, 1,Len(FormatDecimals)-1)
  End If
End Function

Function GetTypeName(TypeID)
  GetTypeName = "Unknown"
  Select Case TypeID
     Case varEmpty    : GetTypeName = "Uninitialized" 
     Case varNull     : GetTypeName = "Null" 
     Case varSmallInt : GetTypeName = "Signed 16-bit integer" 
     Case varInteger  : GetTypeName = "Signed 32-bit integer" 
     Case varSingle   : GetTypeName = "Single-precision floating-point number" 
     Case varDouble   : GetTypeName = "Double-precision floating-point number" 
     Case varCurrency : GetTypeName = "Currency. High-precision floating-point number" 
     Case varDate     : GetTypeName = "Date/Time" 
     Case varOleStr   : GetTypeName = "String" 
     Case varDispatch : GetTypeName = "Automation object of IDispatch interface" 
     Case varError    : GetTypeName = "Code of an OS error" 
     Case varBoolean  : GetTypeName = "Boolean" 
     Case varVariant  : GetTypeName = "Variant" 
     Case varShortInt : GetTypeName = "Signed 8-bit integer" 
     Case varByte     : GetTypeName = "Unsigned 8-bit integer" 
     Case varWord     : GetTypeName = "Unsigned 16-bit integer" 
     Case varLongWord : GetTypeName = "Unsigned 32-bit integer" 
     Case varInt64    : GetTypeName = "Signed 64-bit integer"
  End Select

  If TypeID >= varArray And TypeID <= varArray + varInt64 Then
    GetTypeName = "Array of " & GetTypeName(TypeID - varArray)
  ElseIf TypeID >= varByRef And TypeID <= varByRef + varInt64 Then
    GetTypeName = "Reference to " & GetTypeName(TypeID - varByRef)
  End If
End Function

' Click on the GridObject
Sub ClickGrid(GridObject, Row, Col, Button)
  Dim Grid
  Set Grid = TestConfig.QuantCOREControl
    
   Call MakeCellVisible(GridObject, Row, Col)
  
  Dim CoordArray
    CoordArray = Split(Grid.GetCellCoordinates(GridObject.Handle,Row,Col), "?")
    
    If UBound(CoordArray) = 1 Then
      Select Case Button
      Case "Left"
        Call GridObject.Click(CoordArray(0)-GridObject.ScreenLeft, CoordArray(1)-GridObject.ScreenTop)
      Case "Right"
        If Aliases.MarketView.WaitAliasChild("dlgOrderTicket", 200).Exists Then
          Call Aliases.MarketView.dlgOrderTicket.btnCancel.Click
        End If
        
        If Aliases.MarketView1.WaitAliasChild("dlgOrderTicket", 200).Exists Then
          Call Aliases.MarketView1.dlgOrderTicket.btnCancel.Click
        End If
         
        Call GridObject.ClickR(CoordArray(0)-GridObject.ScreenLeft, CoordArray(1)-GridObject.ScreenTop)
      Case "Double"
        Call GridObject.DblClick(CoordArray(0)-GridObject.ScreenLeft, CoordArray(1)-GridObject.ScreenTop)
      ' Ctrl click for selecting multiple rows  
      Case "Ctrl"
        Call GridObject.Click(CoordArray(0)-GridObject.ScreenLeft, CoordArray(1)-GridObject.ScreenTop, skCtrl)
        
      Case "Middle"
        Sys.Desktop.MouseX = CoordArray(0)
        Sys.Desktop.MouseY = CoordArray(1)
        GridObject.Keys("[F9]")
      Case "Fourth"
        Sys.Desktop.MouseX = CoordArray(0)
        Sys.Desktop.MouseY = CoordArray(1)
        GridObject.Keys("[F11]")
      Case "Fifth"
        Sys.Desktop.MouseX = CoordArray(0)
        Sys.Desktop.MouseY = CoordArray(1)
        GridObject.Keys("[F12]")
      Case Else
        Log.Error("ClickGrid : Unrecognised value for Button : "&Button)
      End Select
      Delay(50)
    Else
      Log.Error("ClickGrid : problem with grid coordinates")
    End If
End Sub

' This function gets the text from the column name ColumnName on the Row number provided
Function GetTextFromRow(GridObject,Row,ColumnName,Instance)
  Dim Grid
  Set Grid = TestConfig.QuantCOREControl
  
  Dim Col
  Col = Grid.GetCellColumn(GridObject.Handle,ColumnName,Instance)
    
  GetTextFromRow = Grid.GetCellText(GridObject.Handle,Row,Col)
End Function

Function GetIntegerFromRow(GridObject,Row,ColumnName,Instance)
  Dim CellText
  CellText = GetTextFromRow(GridObject,Row,ColumnName,Instance)
  
  If RegExpMatch("^-?[0-9]+$",CellText) Then
    GetIntegerFromRow = StrToInt(CellText)
  Else
    GetIntegerFromRow = 0
  End If
End Function

Function GetFloatFromRow(GridObject,Row,ColumnName,Instance)
  Dim CellText
  CellText = GetTextFromRow(GridObject,Row,ColumnName,Instance)
  
  If RegExpMatch("^-?[0-9]+(\.[0-9]+)?$",CellText) Then
    GetFloatFromRow = StrToFloat(CellText)
  End If
End Function

Function GetCellTextColourFromRow(GridObject,Row,ColumnName,Instance)
  Dim Grid
  Set Grid = TestConfig.QuantCOREControl
  
  Dim Col
  Col = Grid.GetCellColumn(GridObject.Handle,ColumnName,Instance)
    
  GetCellTextColourFromRow = Grid.GetCellTextColour(GridObject.Handle,Row,Col)
End Function

Function GetCellBackgroundColourFromRow(GridObject,Row,ColumnName,Instance)
  Dim Grid
  Set Grid = TestConfig.QuantCOREControl
  
  Dim Col
  Col = Grid.GetCellColumn(GridObject.Handle,ColumnName,Instance)
    
  GetCellBackgroundColourFromRow = Grid.GetCellBackgroundColour(GridObject.Handle,Row,Col)
End Function

' Generic function for access the product selection dialog box
' This is used in several applications
Sub SelectProducts(ProductsRef, ProductType, Product, ProductMonth)

  Call ProductsRef.checkSimulation.ClickButton(cbChecked)

  Dim i
  Dim Count
  Dim Found
  Count = 0
  Found = False
  Do Until Found = True Or Count = 2  
    For i = 0 To ProductsRef.List1.wItemCount - 1
      'MS - Product Type list e.g. the left 
      If ProductsRef.List1.wItem(i) = ProductType Then
        Found = True
        Exit For
      End If
    Next
    Count = Count + 1
    Delay(100)
  Loop
  
  If Found = False Then
    Log.Error("AddProduct : did not find product type "&ProductType&" in the list of available product types")
    ProductsRef.btnCancel.Click
    Exit Sub
  End If
    
  ' Click on the product type
  Call ProductsRef.List1.ClickItem(ProductType, 0)
  Delay(500)
    
  ' This keeps trying to select |Product|ProductMonth e.g. |AEX|SEP2010 until wSelection
  ' matches the pattern.  It does a pattern match as the whole string is |AEX|SEP2010     FUTURE but we're only interested in |AEX|SEP2010
  Count = 0
  Found = False
  
    
' Only try three times to find the product - Mei

  Do
    Delay(100)
'    For i = 0 To 30
'      Call ProductsRef.Tree1.Keys("[BS]")    ' Press backspace in GUI to delete what is selected already in case it's wrong
'    Next
    
    Call ProductsRef.Tree1.Keys(Product)
    Call ProductsRef.Tree1.Keys("[Right]")          
    Call ProductsRef.Tree1.Keys(ProductMonth)
             
    Count = Count + 1  
    If RegExpMatch("^\|"&Product&"\|"&ProductMonth, ProductsRef.Tree1.wSelection) = True Then
      Found = True
    End If 
  Loop Until Found = True Or Count = 3
  
 ' If the product is not found then log the error message - Mei
  If Not Found Then
    Log.Message("The Product '" &Product& " " &ProductMonth & "' can not be found.")
    ProductsRef.btnOK.ClickButton
    Exit Sub
  End If  
  
  Call CheckRegExp("Check product is highlighted","^\|"&Product&"\|"&ProductMonth, ProductsRef.Tree1.wSelection)

  ' If the product is found then select the product and log the message - Mei     
  Call ProductsRef.Tree1.Keys(" ")
  Log.Message("The Product '" &Product& " " &ProductMonth & "' has been selected.") 
  Delay(100)            

  ProductsRef.btnOK.ClickButton
End Sub


' Check simulation
Sub SelectProductSet(ProductsRef,ProductType,Product)
  Call ProductsRef.checkSimulation.ClickButton(cbChecked)

  Dim i
  Dim Count
  Dim Found
  Count = 0
  Found = False
  Do Until Found = True Or Count = 2  
    For i = 0 To ProductsRef.List1.wItemCount - 1
      If ProductsRef.List1.wItem(i) = ProductType Then
        Found = True
        Exit For
      End If
    Next
    Count = Count + 1
    Delay(100)
  Loop
  
  If Found = False Then
    Log.Error("SelectProductSet : did not find product type "&ProductType&" in the list of available product types")
    ProductsRef.btnCancel.Click
    Exit Sub
  End If

  Call ProductsRef.List1.ClickItem(ProductType, 0)
  Delay(500)
  
  ' This keeps trying to select |Product|ProductMonth e.g. |AEX|SEP2010 until wSelection
  ' matches the pattern.  It does a pattern match as the whole string is |AEX|SEP2010     FUTURE but we're only interested in |AEX|SEP2010
  Count = 0
  Do
    Delay(100)
    For i = 0 To 30
      Call ProductsRef.Tree1.Keys("[BS]")    ' Press backspace in GUI to delete what is selected already in case it's wrong
    Next
    
    Call ProductsRef.Tree1.Keys(Product)
             
    Count = Count + 1  
  Loop Until RegExpMatch("^\|"&Product, ProductsRef.Tree1.wSelection) = True Or Count = 100
  
  Call CheckRegExp("Check product is highlighted","^\|"&Product, ProductsRef.Tree1.wSelection)
  
  Call ProductsRef.Tree1.Keys(" ")
  Delay(100)
         
  ProductsRef.btnOK.ClickButton
End Sub

Sub WaitUntilAliasVisible(Parent, AliasName, Timeout)

  Dim Count
  Count = 0
  Do Until Parent.WaitAliasChild(AliasName,Timeout).Visible = True Or Count = Round(Timeout / 100, 0)
    Delay(200)
    Count = Count + 1
  Loop
End Sub


Function IsBCGButtonTicked(ButtonObject)
  Dim PictTicked, PictUnticked, PictCurrent, PictResult
  Set PictTicked   = Utils.Picture
  Set PictUnticked = Utils.Picture
  
  Call PictTicked.LoadFromFile("..\BCGButtonTicked.bmp")
  Call PictUnticked.LoadFromFile("..\BCGButtonCross.bmp")
  Set PictCurrent = Sys.Desktop.Picture(ButtonObject.ScreenLeft + 4,ButtonObject.ScreenTop + 4, 20, ButtonObject.Height-8, False)
  
  Call Log.Picture(PictCurrent,"Current image for button "&ButtonObject.MappedName)
  
  Set PictResult = PictTicked.Difference(PictCurrent)
  If PictResult Is Nothing Then
    IsBCGButtonTicked = True
    Exit Function   
  Else
    Set PictResult = PictUnticked.Difference(PictCurrent)  
    If PictResult Is Nothing Then
      IsBCGButtonTicked = False
      Exit Function 
    Else
      Log.Error("IsBCGButtonTicked : state not recognised")
      Exit Function
    End If
  End If
End Function

Sub MakeWindowVisible(WindowObject)
  If WindowObject.Height > (Sys.Desktop.Height - Aliases.Explorer.wndShell_TrayWnd.Height) Then
    Log.Error("Height of the window "&WindowObject.MappedName&" is greater than the desktop height, it cannot be moved so it will all be visible")
    Exit Sub
  End If
  
  Call WindowObject.Position(0,0,WindowObject.Width,WindowObject.Height)
End Sub

Sub LogBoolean(Description, Value)
  If Value = True Then
    Log.Message(Description&" = True")
  Else
    Log.Message(Description&" = False")
  End If
End Sub

Sub WaitForCellValue(GridObject,Row,Col,Value,Timeout)
  Dim Count
  Count = 0
  Do Until QuantCOREControl.GetCellText(GridObject.Handle,Row,Col) = Value Or Count = Round(Timeout / 100, 0)
    Delay(200)
    Count = Count + 1
  Loop
End Sub

Sub MakeCellVisible(GridObject, Row, Col)
  'MS - Updated the function to do a check to see if a Cell is already visible as we don't need to always reset to the start
  'of the Grid, this saves a lot of time.
  
  Dim CoordArray
  Dim pos
  Dim a
               
  If Row < 0 Then
    Log.Error("MakeCellVisible : invalid value for row")
    Exit Sub
  End If
  
  If Col < 0 Then
    Log.Error("MakeCellVisible : invalid value for col")
    Exit Sub
  End If
 
  CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row,Col),"?")
   'Debug - Do not remove please, I need to turn this on for error checking every now and then
   'Log.Message("Cell initial Position " & CoordArray(0) & ", " & CoordArray(1))
   Dim LimitX, LimitY
   LimitX = (GridObject.ScreenLeft + GridObject.Width - 50)
   LimitY = (GridObject.ScreenTop + GridObject.Height - 50)
   'Debug - Do not remove please, I need to turn this on for error checking every now and then
   'Log.Message("Limit X,Y " & LimitX & ", " & LimitY) 
   'Log.Message("GridObject.ScreenLeft+3= " & (GridObject.Screenleft + 3) & " GridObject.ScreenTop+3= " & (GridObject.ScreenTop + 3))   
   
  
  'Check if out of bounds - check if X is less than GridObject+3 (bounds), this usually means Y is in the same boat
  If StrToInt(CoordArray(0)) < (GridObject.ScreenLeft + 3 ) Then
    'Reset to top of the Grid
    GridObject.VScroll.Pos = GridObject.VScroll.Min
    GridObject.HScroll.Pos = GridObject.HScroll.Min   
      'Scroll from top to bottom and stop looking when the coordinates are visible in the grid window
      For pos = GridObject.VScroll.Min To GridObject.VScroll.Max
        GridObject.VScroll.Pos = pos
        CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row,Col),"?")
        'Debug - Do not remove please, I need to turn this on for error checking every now and then
        'Log.Message("Searching Cell Y Position " & CoordArray(0) & ", " & CoordArray(1))        
        If UBound(CoordArray) <> 1 Then
          a=1
        End If      
        If StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + 3 )_ 
        And StrToInt(CoordArray(1)) > (GridObject.ScreenTop + 3)_
        And StrToInt(CoordArray(1)) < (GridObject.ScreenTop + GridObject.Height - 50) Then
          Exit For
        End If
      Next    
      'Debug - Do not remove please, I need to turn this on for error checking every now and then
      'Log.Message("Cell Y was out of bounds")
      'Log.Message("Cell Y Position " & CoordArray(0) & ", " & CoordArray(1))      
          
      'If X is still out of bounds, meaning it couldn't find Y, do a search for X instead
      If StrToInt(CoordArray(0)) < (GridObject.ScreenLeft + 3 ) Then
        'Reset to top of the Grid
        'Debug - Do not remove please, I need to turn this on for error checking every now and then
        'Log.Message("Resetting to the top of the Grid again as Y could not be found, searching for X instead")
        GridObject.VScroll.Pos = GridObject.VScroll.Min
        GridObject.HScroll.Pos = GridObject.HScroll.Min   
        For pos = GridObject.HScroll.Min To GridObject.HScroll.Max
        GridObject.HScroll.Pos = pos
        CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row, Col),"?")
        'Debug - Do not remove please, I need to turn this on for error checking every now and then
        'Log.Message("Searching Cell X Position " & CoordArray(0) & ", " & CoordArray(1))
        If   StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + 3 )_ 
          And StrToInt(CoordArray(1)) > (GridObject.ScreenTop + 3)_
          And StrToInt(CoordArray(0)) < (GridObject.ScreenLeft + GridObject.Width - 50) Then
          Exit For
        End If
        Next
      End If
      
    
    'If X coordinate not in position but within bounds as Y was probably already found on screen
     If StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + GridObject.Width - 50) Then                                             
        GridObject.HScroll.Pos = GridObject.HScroll.Min         
        For pos = GridObject.HScroll.Min To GridObject.HScroll.Max
        GridObject.HScroll.Pos = pos
        CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row, Col),"?")
        If   StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + 3 )_ 
          And StrToInt(CoordArray(1)) > (GridObject.ScreenTop + 3)_
          And StrToInt(CoordArray(0)) < (GridObject.ScreenLeft + GridObject.Width - 50) Then
          Exit For
        End If
        Next
      End If
  
  'If X is already on screen but Y coordinate not in position but not out of bounds
  ElseIf StrToInt(CoordArray(1)) > (GridObject.ScreenTop + GridObject.Height - 50)Then
  'Debug
  'Call Log.Warning("X is already on screen but Y coordinate not in position but within bounds", , 100)
    'Start scrolling downwards from the current Vertical Scroll position    
    For pos = GridObject.VScroll.Pos To GridObject.VScroll.Max
        GridObject.VScroll.Pos = pos
        CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row,Col),"?")
        'Debug - Do not remove please, I need to turn this on for error checking every now and then
        'Log.Message("Searching Cell Y Position " & CoordArray(0) & ", " & CoordArray(1))
        
        If UBound(CoordArray) <> 1 Then
          a=1
        
        End If
      
        If StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + 3 )_ 
        And StrToInt(CoordArray(1)) > (GridObject.ScreenTop + 3)_
        And StrToInt(CoordArray(1)) < (GridObject.ScreenTop + GridObject.Height - 50) Then
          'Debug - Do not remove please, I need to turn this on for error checking every now and then
          'Log.Message("Cell Y Position " & CoordArray(0) & ", " & CoordArray(1)) 
          Exit For
        End If
      Next
      
      If StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + GridObject.Width - 50) Then                                                 
        For pos = GridObject.HScroll.Pos To GridObject.HScroll.Max
        GridObject.HScroll.Pos = pos
        CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row, Col),"?")
        If   StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + 3 )_ 
          And StrToInt(CoordArray(1)) > (GridObject.ScreenTop + 3)_
          And StrToInt(CoordArray(0)) < (GridObject.ScreenLeft + GridObject.Width - 50) Then
          Exit For
        End If
        Next
      End If          
  
  ElseIf StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + GridObject.Width - 50) Then 
  
    For pos = GridObject.HScroll.Pos To GridObject.HScroll.Max
        GridObject.HScroll.Pos = pos
        CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row, Col),"?")
        If   StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + 3 )_ 
          And StrToInt(CoordArray(1)) > (GridObject.ScreenTop + 3)_
          And StrToInt(CoordArray(0)) < (GridObject.ScreenLeft + GridObject.Width - 50) Then
          Exit For
        End If
    Next 
  
  End If

End Sub 
