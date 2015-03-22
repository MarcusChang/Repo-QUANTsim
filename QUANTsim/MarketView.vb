'USEUNIT TestUtilities
'USEUNIT ExcelDriver
Option Explicit

Dim TestConfig
Set TestConfig = ProjectSuite.Variables

Private GridControl
Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl

Private BTN_PULL_ALL_ORDERS
BTN_PULL_ALL_ORDERS = 34915

Private MarketView
Set MarketView = Aliases.MarketView

' MS - Initialize routine is redundant when the TestConfig are declared as above
Sub Initialize
  BTN_PULL_ALL_ORDERS = 34915
  Set GridControl = TestConfig.QuantCOREControl
End Sub

Sub Login(Username, Password)

  Call TestedApps.MarketView.Run(1, True)
  
  Call WaitUntilAliasVisible(Aliases.MarketView, "dlgLogin", 20000)

  Dim dlgLogin
  Set dlgLogin = Aliases.MarketView.dlgLogin
  
  Call dlgLogin.Username.Click
  Call dlgLogin.Username.Keys("[Home]![End][Del]"&Username)
  Call dlgLogin.Password.Keys("[Home]![End][Del]"&Password)
  
  dlgLogin.ComboBox.ClickItem("New Workspace")
  
  dlgLogin.btnOK.ClickButton
  
  Call WaitUntilAliasVisible(Aliases.MarketView, "wndAfx", 20000)
End Sub

'---------------------------------------------------------------------------------------------------------------------
'Sub NewView(ViewType)
'Description:
'---------------------------------------------------------------------------------------------------------------------
Sub NewView(ViewType)
  Dim MenuToolbar
  Set MenuToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.MenuBar

  ' This is not an ideal solution... however, when you open market view the toolbar first looks like this:
  ' File, Edit, View, Window, Help
  ' When you open the first view, they all shift right one, and you get
  ' DOC, File, Edit, View, Window, Help
  ' When you open a second view, the DOC icon disappears again and you get
  ' File, Edit, View, Window, Help
  ' There is not an elegant solution as TestComplete cannot get the text of the buttons in the menu bar
'  Dim ViewButtonId
'  If MenuToolBar.wButtonText(0) = "" Then
'    ViewButtonId = 1
'  ElseIf MenuToolBar.wButtonText(0) = "&File" Then
'    ViewButtonId = 0
'  End If
'  
'  Call MenuToolbar.ClickItem(ViewButtonId, True)
'  
'   ' Open the New View control
'  Call Aliases.MarketView.wndAfx1.BCGPToolBar40000081001510.ClickItem("&New"&Chr(9)&"Ctrl+N", False)
  
  Dim MDIClient
  Set MDIClient = Aliases.MarketView.wndAfx.MDIClient
  Call MDIClient.Keys("^n")

  
  Dim dlgNew
  Set dlgNew = Aliases.MarketView.dlgNew
  ' Check if ViewType exists in the list of view types
  If Not ItemInList(dlgNew.NewViewListBox.wItemList,dlgNew.NewViewListBox.wListSeparator,ViewType) Then
    Log.Error("NewView : ViewType is not recognised : "&ViewType)
    dlgNew.btnCancel.ClickButton  
    Exit Sub
  End If
  
  Call dlgNew.NewViewListBox.ClickItem(ViewType)
  dlgNew.btnOK.ClickButton
  Delay(2000)
End Sub

'---------------------------------------------------------------------------------------------------------------------
'Sub OpenSheet(Filename)
'Description:
'---------------------------------------------------------------------------------------------------------------------
Sub OpenSheet(AppName, Filename)
  
  'Check if MarketView is running by scanning for the main window 'wndAfx'
    Dim MDIClient, dlgOpen
    Dim MarketView, OptionView  
    
    
   Select case AppName
   
    Case "MarketView"
        If Not Aliases.MarketView.WaitAliasChild("wndAfx",200).Exists Then
          Log.Error("Error: Cannot find an instance of MarketView. Check if it is running.")
          Exit Sub
        End If
  
 
        
  
         Set MDIClient = Aliases.MarketView.wndAfx.MDIClient
         Set dlgOpen = Aliases.MarketView.dlgOpen
   
       
         Set MarketView = Aliases.MarketView
         Set OptionView = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
   
  
        Call MDIClient.Click(480, 217)
        Call MDIClient.Keys("^o")
        Call dlgOpen.FilePath.SetText(Filename)
        dlgOpen.btnOpen.ClickButton

        Delay(1000)
        
        If Not OptionView.Exists Then
          If MarketView.dlgMarketViewName.Exists Then         
              Log.Error("The sheet '" &Filename & "' was not found.")
              Call MarketView.dlgMarketViewName.btnOK.Click 
            ElseIf   MarketView.dlgInvalidFileType.Exists Then
              Log.Error("The sheet '" &Filename & "' was not found.")
              Call MarketView.dlgInvalidFileType.btnOK.Click
           End If  
        End If  
        
        
    Case "MarketView1"
    
        If Not Aliases.MarketView1.WaitAliasChild("wndAfx",200).Exists Then
          Log.Error("Error: Cannot find an instance of MarketView. Check if it is running.")
          Exit Sub
        End If
  
        Set MDIClient = Aliases.MarketView1.wndAfx.MDIClient
        Set dlgOpen = Aliases.MarketView1.dlgOpen
    
        Set MarketView = Aliases.MarketView1
        Set OptionView = Aliases.MarketView1.wndAfx.MDIClient.OptionView1.OptionViewGrid
   
  
        Call MDIClient.Click(480, 217)
        Call MDIClient.Keys("^o")
        Call dlgOpen.FilePath.SetText(Filename)
        dlgOpen.btnOpen.ClickButton

        Delay(1000)
   
      If Not OptionView.Exists Then   
        If MarketView.dlgMarketViewName.Exists Then         
            Log.Error("The sheet '" &Filename & "' was not found.")
            Call MarketView.dlgMarketViewName.btnOK.Click 
          ElseIf   MarketView.dlgInvalidFileType.Exists Then
            Log.Error("The sheet '" &Filename & "' was not found.")
            Call MarketView.dlgInvalidFileType.btnOK.Click
        End If 
     End If   
        
  End Select
  
End Sub


'---------------------------------------------------------------------------------------------------------------------
'Sub RemoveAllProducts()
'Description: Remove all products from Product Selection
'---------------------------------------------------------------------------------------------------------------------
Sub RemoveAllProducts()

  Dim MenuToolBar, MenuToolbarPopup
  Set MenuToolBar = MarketView.wndAfx.BCGPDockBar.MenuBar
'  Set MenuToolBarPopup = MarketView.wndAfx1.BCGPToolBar40000081001510 
'
'  Dim ViewButtonId
'  If MenuToolBar.wButtonText(0) = "" Then
'    ViewButtonId = 3
'  ElseIf MenuToolBar.wButtonText(0) = "&File" Then
'    ViewButtonId = 2
'  End If
'  
'  'Open up Product Selection menu
'  Call MenuToolbar.ClickItem(ViewButtonId, True)   
'  Call MenuToolbarPopup.ClickItem("&Product Selection", False)

  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
    
  Call OptionViewGrid.Click(480, 217)
  Call OptionViewGrid.Keys("~vp")

  
  Call WaitUntilAliasVisible(MarketView, "dlgProducts", 5000)
  
  Dim ProductSelection, ProductList, i
  Set ProductSelection = MarketView.dlgProducts1
  Set ProductList = MarketView.dlgProducts.List1

  If ProductList.wItemCount = 0 Then
    Log.Message("No Products to delete")
    ProductSelection.btnClose.Click
    Exit Sub
  End If
    
  For i = 0 To ProductList.wItemCount - 1
    Call ProductList.ClickItem(ProductList.wItem(i), , skShift)
  Next
  
  ProductSelection.btnDelete.Click
  ProductSelection.btnClose.Click  
    
End Sub

'----------------------------------------
'04/11/2011
Sub ProductSelection(State)
  
  If State = "Open" Then
  
    Dim MenuToolBar, MenuToolbarPopup
    Set MenuToolBar = MarketView.wndAfx.BCGPDockBar.MenuBar
'    Set MenuToolBarPopup = MarketView.wndAfx1.BCGPToolBar40000081001510 
'
'    Dim ViewButtonId
'    If MenuToolBar.wButtonText(0) = "" Then
'      ViewButtonId = 3
'    ElseIf MenuToolBar.wButtonText(0) = "&File" Then
'      ViewButtonId = 2
'    End If
'  
'  'Open up Product Selection menu
'    Call MenuToolbar.ClickItem(ViewButtonId, True)   
'    Call MenuToolbarPopup.ClickItem("&Product Selection", False)

  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
    
  Call OptionViewGrid.Click(480, 217)
  Call OptionViewGrid.Keys("~vp")
  Call WaitUntilAliasVisible(MarketView, "dlgProducts1", 5000)
  
  ElseIf State = "Close" Then
  
    Dim ProductSelection
    Set ProductSelection = MarketView.dlgProducts1
      
    ProductSelection.btnClose.Click 
  
  End If

End Sub

Sub AddProductMultiples(ProductType, Product, ProductMonth)
    
    Dim ProductsList
    Dim ProductsAdd
    
    Set ProductsList = Aliases.MarketView.dlgProducts1
    Set ProductsAdd = Aliases.MarketView.dlgProducts
    Call ProductsList.btnAdd.ClickButton
    Call WaitUntilAliasVisible(Aliases.MarketView, "dlgProducts", 5000)
    
'    Set ProductsAdd = Aliases.MarketView.dlgProducts1
    
    If ProductType = "EQUITY" Then
      Call SelectProductSet(ProductsAdd,ProductType,Product)    
    Else    
      Call SelectProducts(ProductsAdd, ProductType, Product, ProductMonth)
    End If
    
 'If "Selection Conflicts" error pops up, close the error message
    If Aliases.MarketView.WaitAliasChild("dlgMarketViewName", 200).Exists Then
                  
          Log.Message("There was an attempt to add a Product that conflicted with existing Selections")
          Call Aliases.MarketView.dlgMarketViewName.btnOK.Click
    End If
    
    'ProductsList.btnClose.ClickButton
    Delay(750)

End Sub

Sub AddProduct(ProductType, Product, ProductMonth)
    Dim MenuToolbar
    Dim MenuToolbarPopup
    Dim ProductsList
    Dim ProductsAdd

    Set MenuToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.MenuBar
'    Set MenuToolbarPopup = Aliases.MarketView.wndAfx1.BCGPToolBar40000081001510

    ' This is not an ideal solution... however, when you open market view the toolbar first looks like this:
    ' File, Edit, View, Window, Help
    ' When you open the first view, they all shift right one, and you get
    ' DOC, File, Edit, View, Window, Help
    ' When you open a second view, the DOC icon disappears again and you get
    ' File, Edit, View, Window, Help
    ' There is not an elegant solution as TestComplete cannot get the text of the buttons in the menu bar
'    Dim ViewButtonId
'    If MenuToolBar.wButtonText(0) = "" Then
'      ViewButtonId = 3
'    ElseIf MenuToolBar.wButtonText(0) = "&File" Then
'      ViewButtonId = 2
'    End If
'    
'    Call MenuToolbar.ClickItem(ViewButtonId, True)
'    
'    Call MenuToolbarPopup.ClickItem("&Product Selection", False)
'             
'    Call WaitUntilAliasVisible(Aliases.MarketView, "dlgProducts", 5000)
'    
'    Set  ProductsList = Aliases.MarketView.dlgProducts
'    
'    Call ProductsList.btnAdd.ClickButton

  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid

  Call OptionViewGrid.Keys("~vp")
  Set  ProductsList = Aliases.MarketView.dlgProducts1
  Call ProductsList.btnAdd.ClickButton

    Call WaitUntilAliasVisible(Aliases.MarketView, "dlgProducts", 5000)
    
    'MS - Products Dialog Box - the one where you select Product Type, then product, and check Simulation etc
    Set  ProductsAdd = Aliases.MarketView.dlgProducts
    
    'MS - function from TestUtilities
    Call SelectProducts(ProductsAdd, ProductType, Product, ProductMonth)

    'If "Selection Conflicts" error pops up, close the error message
    If Aliases.MarketView.WaitAliasChild("dlgMarketViewName", 200).Exists Then
                  
          Log.Message("There was an attempt to add a Product that conflicted with existing Selections")
          Call Aliases.MarketView.dlgMarketViewName.btnOK.Click
    End If
    
    
    ProductsList.btnClose.ClickButton
    
    Delay(1000)
End Sub

 
Sub AddProductSet(ProductType, Product)
    Dim MenuToolbar
    Dim MenuToolbarPopup
    Dim ProductsList
    Dim ProductsAdd

    Set MenuToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.MenuBar
'    Set MenuToolbarPopup = Aliases.MarketView.wndAfx1.BCGPToolBar40000081001510
'
'    ' This is not an ideal solution... however, when you open market view the toolbar first looks like this:
'    ' File, Edit, View, Window, Help
'    ' When you open the first view, they all shift right one, and you get
'    ' DOC, File, Edit, View, Window, Help
'    ' When you open a second view, the DOC icon disappears again and you get
'    ' File, Edit, View, Window, Help
'    ' There is not an elegant solution as TestComplete cannot get the text of the buttons in the menu bar
'    Dim ViewButtonId
'    If MenuToolBar.wButtonText(0) = "" Then
'      ViewButtonId = 3
'    ElseIf MenuToolBar.wButtonText(0) = "&File" Then
'      ViewButtonId = 2
'    End If
'    
'    Call MenuToolbar.ClickItem(ViewButtonId, True)
'    
'    Call MenuToolbarPopup.ClickItem("&Product Selection", False)

  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
    
  Call OptionViewGrid.Click(480, 217)
  Call OptionViewGrid.Keys("~vp")


    Call WaitUntilAliasVisible(Aliases.MarketView, "dlgProducts", 5000)
       
    Set  ProductsList = Aliases.MarketView.dlgProducts1
    
    Call ProductsList.btnAdd.ClickButton

    Call WaitUntilAliasVisible(Aliases.MarketView, "dlgProducts", 5000)
    
    Set  ProductsAdd = Aliases.MarketView.dlgProducts 
     
    Call SelectProductSet(ProductsAdd,ProductType,Product)
    
    ProductsList.btnClose.ClickButton
    
    Delay(1000) 
End Sub

Function ReplaceStrikes(str)
    Dim rex
    Set rex= New RegExp ' create the regular expression instance
    rex.IgnoreCase = True ' set whether ignore the Caps Lock
    rex.Global = True ' set whether it's global affacted
    rex.Pattern = "^(\d*)\.(\d*)$"
    ReplaceStrikes = rex.replace(str, "$1") 'replace the strikes without decimals   
End Function


Public Function GetOptionStrikes(Product, ProductMonth)
  Dim OptionViewGrid, i, Strikes, StrikeText, StrikeText_1
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
  
  Dim MatchString_0, MatchString_1
  MatchString_0 = "^SIM\.O\."&Product&"\."&ProductMonth&".*C\.0$"
  MatchString_1 = "^SIM\.O\."&Product&"\."&ProductMonth&".*C\.1$"

  Dim DecimalsColumns, Row_DecimalsTable
  Set DecimalsColumns = ExcelDriver.GetDataTable(TestConfig.xlConfigFile,"MarketView","MarketView_Columns")
  
 
  
  ' Search through all the strikes in the series Product / Product Month and get the strike
  For i = 1 To GridControl.GetRowCount(OptionViewGrid.Handle)
    If RegExpMatch(MatchString_0, GetTextFromRow(OptionViewGrid,i,"ProductID",1)) = True Then
    
      For Row_DecimalsTable = 2 to DecimalsColumns.ListRows.Count + 1 
        Dim Name, Decimals, NeedReg
        Name = DecimalsColumns.Range.Cells(Row_DecimalsTable, 1).Value
        Decimals = DecimalsColumns.Range.Cells(Row_DecimalsTable, 2).Value
        If Name = "Strike" and Decimals <> "" Then
          NeedReg = True          
          Exit For
        End If        
      Next
        
        
        If NeedReg Then
          
          StrikeText = GetTextFromRow(OptionViewGrid, i, "Strike", 1)
      
          Call ReplaceStrikes(StrikeText)              
      
          Strikes = Strikes & ReplaceStrikes(StrikeText) & "|"
                        
        Else   
          
          Strikes = Strikes & GetTextFromRow(OptionViewGrid,i,"Strike",1)&"|"          
          
        End If
        
    
    ElseIf RegExpMatch(MatchString_1, GetTextFromRow(OptionViewGrid, i, "ProductID", 1)) = True Then
           
      For Row_DecimalsTable = 2 to DecimalsColumns.ListRows.Count + 1 
        Dim Name_1, Decimals_1, NeedReg_1
        Name_1 = DecimalsColumns.Range.Cells(Row_DecimalsTable, 1).Value
        Decimals_1 = DecimalsColumns.Range.Cells(Row_DecimalsTable, 2).Value
        If Name_1 = "Strike" and Decimals_1 <> "" Then
          NeedReg_1 = True          
          Exit For
        End If        
      Next
        
        If NeedReg_1 Then
            
          StrikeText_1 = GetTextFromRow(OptionViewGrid, i, "Strike", 1)
      
          Call ReplaceStrikes(StrikeText_1)              
      
          Strikes = Strikes & ReplaceStrikes(StrikeText_1) & "|" 
        
        Else   
          Strikes = Strikes & GetTextFromRow(OptionViewGrid,i,"Strike",1)&"|"
          
        End If
            
    End If
  Next
  

  GetOptionStrikes = Mid(Strikes,1,Len(Strikes)-1)  
  
End Function
  
Public Function GetAtTheMoneyStrike(Product, ProductMonth, CallPut, Strikes)
  Dim ColumnInstance
  Select Case CallPut
  Case "Call"
    ColumnInstance = 1
  Case "Put"
    ColumnInstance = 2
  Case Else
    Call Log.Error("GetAtTheMoneyStrike : unknown value for CallPut - """&CallPut&"""")
    Exit Function
  End Select 
  
  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
    
  Dim Strike, Row, Col, Delta, DeltaError, ATMStrike, ATMDelta, TargetDelta
  DeltaError = 2
  TargetDelta = 0.5
  
  ' Cycle through every strike in the list
  For Each Strike In Split(Strikes, "|")
    ' Get the delta for the current strike     Name          &" "&ProductMonth&" "&Strike&"C"
    Row = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID", "SIM.O."&Product&"."&ProductMonth&".C.0", 1)
    Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Delta", ColumnInstance)
    Delta = StrToFloat(GridControl.GetCellText(OptionViewGrid.Handle,Row,Col))
    
    ' Calculate whether it is the ATM strike
    ' This is defined as being:
    ' The strike which is closest to 0.5
    ' The strike should also be on the 'In The Money' side if it is not exactly 0 
    ' e.g. for Calls, a delta of 0.52 could be ATM
    '    for Puts, a delta of -0.52 could be ATM
    ' Abs function is used to switch sign of negative delta values, so that the same
    ' check works for calls and puts
    If (Abs(Delta) - TargetDelta) < DeltaError And (Abs(Delta) - TargetDelta) > 0  Then
      ATMStrike  = Strike
      ATMDelta = Delta
      DeltaError = Abs(Delta)- TargetDelta
    End If
  Next

  Call Log.Message("GetAtTheMoneyStrike for "&Product&", "&ProductMonth&", "&CallPut&" : ATMStrike = "&ATMStrike&", ATMDelta = "&ATMDelta)
  GetAtTheMoneyStrike = ATMStrike
End Function   
  
Sub MakeCellVisible(GridObject, Row, Col)
  'MS - need to do a check if the cell is already visible so you dont waste time resetting the Scroll positions
  
  Dim CoordArray
  Dim pos
  Dim a
  
  'MS - Temporary Declaration
  Dim GridControl
  Set GridControl = TestConfig.QuantCOREControl
                
  If Row < 0 Then
    Log.Error("MakeCellVisible : invalid value for row")
    Exit Sub
  End If
  
   If Col < 0 Then
    Log.Error("MakeCellVisible : invalid value for col")
    Exit Sub
  End If
 
  CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row,Col),"?")
  Log.Message("Cell initial Position " & CoordArray(0) & ", " & CoordArray(1))
  Log.Message("GridOBjectScreenLeft " & GridObject.ScreenLeft)
  Log.Message("GridObject.ScreenTop " & GridObject.ScreenTop)
  Log.Message("GridObject.Height " & GridObject.Height)   
  Log.Message("GridObject.VScroll.Pos " & GridObject.VScroll.Pos)  
    
      'Reset to top left of grid
      If GridObject.VScroll.Pos <> 0 Then
        GridObject.VScroll.Pos = GridObject.VScroll.Min
      End If
  
      If GridObject.HScroll.Pos <> 0 Then
        GridObject.HScroll.Pos = GridObject.HScroll.Min
      End If
      
      CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row,Col),"?")
      Log.Message("Cell Position after reset " & CoordArray(0) & ", " & CoordArray(1))
  
      ' Scroll from top to bottom and stop looking when the coordinates are visible in the grid window
      If GridObject.VScroll.Pos <> 0 Then
        For pos = GridObject.VScroll.Min To GridObject.VScroll.Max
        GridObject.VScroll.Pos = pos
        CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row,Col),"?")
        If UBound(CoordArray) <> 1 Then
        a=1
        End If
        If   StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + 3 )_ 
          And StrToInt(CoordArray(1)) > ( + 3)_
          And StrToInt(CoordArray(1)) < (GridObject.ScreenTop + GridObject.Height - 50) Then
           Log.Message("Cell Y Position " & CoordArray(0) & ", " & CoordArray(1))
           
           Log.Message("GridObject.VScroll.Pos " & GridObject.VScroll.Pos)
          Exit For
        End If
        Next
      End If
  
      If GridObject.HScroll.Pos <> 0 Then
        ' Scroll from left to right and stop looking when the coordinates are visible in the grid window
        For pos = GridObject.HScroll.Min To GridObject.HScroll.Max
        GridObject.HScroll.Pos = pos
        CoordArray = Split(GridControl.GetCellCoordinates(GridObject.Handle,Row, Col),"?")
        If   StrToInt(CoordArray(0)) > (GridObject.ScreenLeft + 3 )_ 
          And StrToInt(CoordArray(1)) > (GridObject.ScreenTop + 3)_
          And StrToInt(CoordArray(0)) < (GridObject.ScreenLeft + GridObject.Width - 50) Then
            Log.Message("Cell X Position " & CoordArray(0) & ", " & CoordArray(1))
            
           Log.Message("GridObject.HScroll.Pos " & GridObject.HScroll.Pos)
          Exit For
        End If
        Next
      End If
End Sub 
  
Public Sub ClickTab(TabName)
    'Temporary Declaration - need to make this global?
  Set TestConfig.QuantCOREControl = QuantCOREControl 
  Dim GridControl, OptionViewGrid
  Set GridControl = TestConfig.QuantCOREControl
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
  
  
  Dim MarketViewWindow
  Set MarketViewWindow = Aliases.MarketView.wndAfx
  
  Dim ViewNames, CoordArray
  ViewNames = GridControl.GetViewNames(MarketViewWindow.Handle)
    
  Dim Item
  Dim Found
  Found = False
  For Each Item In Split(ViewNames,"?")
    If RegExpMatch("^"&TabName&"( \([0-9]+ new row[s]?\))?$",Item) Then
    
    Call Log.Message("Clicking on the tab """&Item&"""")
    
    CoordArray = Split(GridControl.GetViewTabCoordinates(MarketViewWindow.Handle, Item),"?")

    If UBound(CoordArray) = 1 Then 
      Call MarketViewWindow.Click(CoordArray(0)-MarketViewWindow.ScreenLeft, CoordArray(1)-MarketViewWindow.ScreenTop)
    End If
    Found = True
    Exit For
    End If
  Next
  If Not Found Then
    Log.Error("ClickTab : There is not a tab called "&TabName&" available to click")
  End If
End Sub
  
Public Sub OpenDepth(ProductID, Instance)
  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid

  'MS - Temporary Declaration     
  Dim GridControl
  Set GridControl = TestConfig.QuantCOREControl
      
  Dim Row, Col, Depth1

  Row = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID", ProductID, Instance)
  Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Depth Stack Position", Instance)
  'Log.Message("OpenDepth " & Row & ", " & Col & ", " & Instance)
  
  Depth1 = GridControl.GetCellText(OptionViewGrid.Handle,Row+1,Col)
  
  If Depth1 <> "1" Then
    'Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Strike", Instance)
    
    Call ClickGrid(OptionViewGrid, Row, Col, "Double")
  End If  
End Sub
  
Public Sub CloseDepth(ProductID)
  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
    
  Dim Row, Col, Depth1

  Row = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID", ProductID, 1)
  Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Depth Stack Position", 1)
  Depth1 = StrToInt(GridControl.GetCellText(OptionViewGrid.Handle,Row+1,Col))
  
  If Depth1 = "1" Then
    Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Strike", 1)
    
    Call ClickGrid(OptionViewGrid, Row, Col, "Double")
  End If  
End Sub

Function GetTickSize(ProductID)
  Dim OptionViewGrid, Row, Col
  
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
  
  Row = GridControl.GetCellRow(OptionViewGrid.Handle, "ProductID",ProductID, 1)          
  Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Bid", 1)
  
  Call TestUtilities.MakeCellVisible(OptionViewGrid, Row, Col)
  
  Call ClickGrid(OptionViewGrid, Row, Col, "Right")
  Delay(300)

  Dim dlgOrderTicket
  
  If Aliases.MarketView.WaitAliasChild("dlgOrderTicket", 50).Exists Then
    Set dlgOrderTicket = Aliases.MarketView.dlgOrderTicket
  ElseIf Aliases.MarketView.WaitAliasChild("dlgAmendOrder", 50).Exists Then
    Set dlgOrderTicket = Aliases.MarketView.dlgAmendOrder
  Else
    Log.Error("Order ticket did not appear")
    Exit Function
  End If

  dlgOrderTicket.PriceArrow.Up
  Dim Before, After
  Before = dlgOrderTicket.Price.wText
  dlgOrderTicket.PriceArrow.Up
  After = dlgOrderTicket.Price.wText
  dlgOrderTicket.btnCancel.ClickButton
  
  GetTickSize = (StrToFloat(After) - StrToFloat(Before))
  GetTickSize = StrToFloat(FormatDecimals(GetTickSize,5))
End Function

Sub EnableDockingMachineView
  Dim MenuToolbar
  Dim MenuToolbarPopup
  Dim DockingView 
  Set MenuToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.MenuBar
  Set MenuToolbarPopup = Aliases.MarketView.wndAfx1.BCGPToolBar40000081001510
  Set DockingView = Aliases.MarketView.wndAfx.DockingMachineView
  
  If Aliases.MarketView.wndAfx.WaitAliasChild("DockingMachineView").Visible = False Then
    Dim ViewButtonId
  
    If MenuToolBar.wButtonText(0) = "" Then
      ViewButtonId = 3
    ElseIf MenuToolBar.wButtonText(0) = "&File" Then
      ViewButtonId = 2
    End If
  
    Call MenuToolbar.ClickItem(ViewButtonId, True)
    
    Call WaitUntilAliasVisible(Aliases.MarketView.wndAfx1, "BCGPToolBar40000081001510", 10000)
  
    Call MenuToolbarPopup.ClickItem("&Docking View", False)
    
    Call WaitUntilAliasVisible(Aliases.MarketView.wndAfx, "DockingMachineView", 10000)
  End If
End Sub

Sub DisableDockingMachineView
  Dim MenuToolbar
  Dim MenuToolbarPopup
  Dim DockingView 
  Set MenuToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.MenuBar
  Set MenuToolbarPopup = Aliases.MarketView.wndAfx1.BCGPToolBar40000081001510
  Set DockingView = Aliases.MarketView.wndAfx.DockingMachineView
  
  If Aliases.MarketView.WaitAliasChild("dlgOrderTicket").Visible = False Then
    Dim ViewButtonId
  
    If MenuToolBar.wButtonText(0) = "" Then
      ViewButtonId = 3
    ElseIf MenuToolBar.wButtonText(0) = "&File" Then
      ViewButtonId = 2
    End If
  
    Call MenuToolbar.ClickItem(ViewButtonId, True)
    
    Call WaitUntilAliasVisible(Aliases.MarketView.wndAfx1, "BCGPToolBar40000081001510", 1000) 
  
    Call MenuToolbarPopup.ClickItem("&Docking View", False)
    
    Call WaitUntilAliasVisible(Aliases.MarketView.wndAfx, "DockingMachineView", 1000)
  End If
End Sub

Function GetMachineCalculatingState(MachineSpecName)
  Dim DockingView
  Set DockingView = Aliases.MarketView.wndAfx.DockingMachineView
  
  Dim Calculating, Row, Col
  Row = GridControl.GetCellRow(DockingView.Grid.Handle,"Name",MachineSpecName,1)
  Col = GridControl.GetCellColumn(DockingView.Grid.Handle,"Calculating",1)
  Calculating = GridControl.GetCellText(DockingView.Grid.Handle,Row,Col)

  GetMachineCalculatingState = Calculating
End Function

Sub WaitUntilMachineStopsCalculating(MachineSpecName, MilliSeconds)
  Dim Count, i
  Count = Round(Milliseconds / 1000,0)
  For i = 0 To Count
    If MarketView.GetMachineCalculatingState(MachineSpecName) = "FALSE" Then
      Exit For
    End If
    Delay(1000)
  Next
  
  If MarketView.GetMachineCalculatingState(MachineSpecName) = "TRUE" Then
    Log.Error("Machine is still calculating")
  End If
End Sub 

Sub WaitUntilMachineStartsCalculating(MachineSpecName, MilliSeconds)
  Dim Count, i
  Count = Round(Milliseconds / 1000,0)
  For i = 0 To Count
    If MarketView.GetMachineCalculatingState(MachineSpecName) = "TRUE" Then
      Exit For
    End If
    Delay(1000)
  Next
  
  If MarketView.GetMachineCalculatingState(MachineSpecName) = "FALSE" Then
    Log.Error("Machine is not calculating")
  End If
End Sub

Sub PullAllOrders
  Call Aliases.MarketView.wndAfx.BCGPDockBar.OptionViewToolbar.ClickItem(BTN_PULL_ALL_ORDERS, False)
End Sub

Sub OpenUserManual(ProductID)
  Dim OptionViewGrid, Row, Col
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
  
  Row = GridControl.GetCellRow(OptionViewGrid.Handle, "ProductID",ProductID, 1)          
  Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"User Manual", 1)
  
  Call MakeCellVisible(OptionViewGrid, Row, Col)
  
  Call ClickGrid(OptionViewGrid, Row, Col, "Left")
  
  Call WaitUntilAliasVisible(Aliases.MarketView,"dlgUserManualEntry",1000)
End Sub


Public Sub SetUserManualPrice(Price, Method)
  Dim CurrentQuantity, Count
  
  Dim dlgUserManualEntry
  Set dlgUserManualEntry = Aliases.MarketView.dlgUserManualEntry
  
  If Method = "Keys" Then
    Call dlgUserManualEntry.Price.Keys("[Home]![End][Del]"&Price)
  
  ElseIf Method = "MouseWheel" Then
    If FMod(Price, 1) <> 0 Then
      Log.Error("SetUserManualPrice - when using MouseWheel, the Price should be specified in whole number of notches to move the mouse wheel")
      Exit Sub
    End If
  
    dlgUserManualEntry.Price.Click
    
    Dim i
    For i = 0 To Abs(Price) - 1
      If Price > 0 Then
        dlgUserManualEntry.Price.MouseWheel(1)
      Else
        dlgUserManualEntry.Price.MouseWheel(-1)
      End If
    Next
  Else
    Log.Error("SetUserManualPrice : unknown value for Method")
  End If    
  
  Call dlgUserManualEntry.btnSend.ClickButton
End Sub

'---------------------------------------------------------------------------------------------------------------------
'Sub CloseSheet(Filename)
'Description:
'---------------------------------------------------------------------------------------------------------------------
Sub CloseSheet
  
  'Check if MarketView is running by scanning for the main window 'wndAfx'
  If Not MarketView.WaitAliasChild("wndAfx",200).Exists Then
    Log.Error("Error: Cannot find an instance of MarketView. Check if it is running.")
    Exit Sub
  End If
  
  Dim MenuToolbar
  Set MenuToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.MenuBar

  ' This is not an ideal solution... however, when you open market view the toolbar first looks like this:
  ' File, Edit, View, Window, Help
  ' When you open the first view, they all shift right one, and you get
  ' DOC, File, Edit, View, Window, Help
  ' When you open a second view, the DOC icon disappears again and you get
  ' File, Edit, View, Window, Help
  ' There is not an elegant solution as TestComplete cannot get the text of the buttons in the menu bar
  Dim ViewButtonId
  If MenuToolBar.wButtonText(0) = "" Then
    ViewButtonId = 1
  ElseIf MenuToolBar.wButtonText(0) = "&File" Then
    ViewButtonId = 0
  End If
  
  'Click on File
  Call MenuToolbar.ClickItem(ViewButtonId, True)

  'Open the Open Dialog box
  Call Aliases.MarketView.wndAfx1.BCGPToolBar40000081001510.ClickItem("&Close", False)  
  
  'IF the 'Save Changes to XXX' dialog appears
    If Aliases.MarketView.WaitAliasChild("dlgMarketView", 200).Exists Then                         
          Call Aliases.MarketView.dlgMarketView.btnNo.Click
    End If
  
  
End Sub

