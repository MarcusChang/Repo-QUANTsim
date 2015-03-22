' Class used for controlling OrderView
'USEUNIT TestUtilities
'USEUNIT MarketView
''USEUNIT MarketView1


Private GridControl
Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl

Private OrderView, OrderViewGrid, dlgColumnSetup, dlgAmendOrder, dlgFilters
Set OrderView = Aliases.OrderView
Set OrderViewGrid = Aliases.OrderView.wndAfx.OrderViewGrid
Set dlgColumnSetup = Aliases.OrderView.dlgColumnSetup
Set dlgAmendOrder = Aliases.OrderView.dlgAmendOrder
Set dlgFilters = Aliases.OrderView.dlgFilters


Sub Login(Username, Password)
  
  Call TestedApps.OrderView.Run(1, True)
  Call WaitUntilAliasVisible(OrderView, "dlgLogin", 20000)   
  Call OrderView.dlgLogin.Username.Click
  Call OrderView.dlgLogin.Username.Keys("[Home]![End][Del]"&Username)
  Call OrderView.dlgLogin.Password.Keys("[Home]![End][Del]"&Password)
  OrderView.dlgLogin.btnOK.ClickButton
  Call WaitUntilAliasVisible(OrderView, "wndAfx", 10000)

End Sub
  
' ------------------------------------------------------------------------------------
' Get a list of all the order ids in OrderView
' ------------------------------------------------------------------------------------
Public Function GetCurrentOrderIDs
  Dim i, OrderArray
  
  '   Make space for the list of order ids
  Redim OrderArray(GridControl.GetRowCount(OrderViewGrid.Handle)-1)

  For i = 0 To GridControl.GetRowCount(OrderViewGrid.Handle) - 1
    OrderArray(i) = GridControl.GetCellText(OrderViewGrid.Handle,i+1,1)
  Next
  
  GetCurrentOrderIDs = OrderArray
End Function
  
' ------------------------------------------------------------------------------------
' Get the order details
' ------------------------------------------------------------------------------------
Public Function GetOrder(OrderID)
  Dim Row, Col
  
  Dim MyOrder
  Set MyOrder = New Order                                                          
  
  Row = GridControl.GetCellRow(OrderViewGrid.Handle, "Order ID",OrderID,1)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Order ID",1)
  MyOrder.OrderID = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Theo",1)
  MyOrder.Theo = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Product Name",1)
  MyOrder.ProductName = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "ProductID",1)
  MyOrder.ProductID = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Account",1)
  MyOrder.Account = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Buy/Sell",1)
  MyOrder.BuySell = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Order Status",1)
  MyOrder.OrderStatus = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Date",1)
  MyOrder.Date = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Order Type",1)
  MyOrder.OrderType = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Volume",1)
  MyOrder.Volume = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Executed Volume",1)
  MyOrder.ExecutedVolume = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Residual Volume",1)
  MyOrder.ResidualVolume = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Price",1)
  MyOrder.Price = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Strike",1)
  MyOrder.Strike = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "User",1)
  MyOrder.User = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Order Restriction",1)
  MyOrder.Restriction = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Month",1)
  MyOrder.Month = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Time",1)
  MyOrder.Time = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "AS Price",1)
  MyOrder.ASPrice = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "AC Price",1)
  MyOrder.ACPrice = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "AC Offset",1)
  MyOrder.ACOffset = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "TOM Price",1)
  MyOrder.TOMPrice = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "TOM Type",1)
  MyOrder.TOMType = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "TOM Note",1)
  MyOrder.TOMNote = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  
  Set GetOrder = MyOrder
End Function
  
' ------------------------------------------------------------------------------------
' Enable a column
' ------------------------------------------------------------------------------------
Public Sub EnableColumn(ColumnName)
   
  ' If the column name already appears in the selected items list, then we can exit the sub
  If ItemInList(dlgColumnSetup.SelectedFields.wItemList,dlgColumnSetup.SelectedFields.wListSeparator,ColumnName) Then
    Exit Sub
  End If
  
  ' Otherwise, see if the column is available to select and if it is double click on it to move it to selected items
  If ItemInList(dlgColumnSetup.AvailableFields.wItemList,dlgColumnSetup.AvailableFields.wListSeparator,ColumnName) Then
    Call dlgColumnSetup.AvailableFields.DblClickItem(ColumnName)
  Else
    Log.Error("EnableColumn : the column """&ColumnName&""" was not available to enable")
  End If   
End Sub
 
' ------------------------------------------------------------------------------------
' Disable a column
' ------------------------------------------------------------------------------------
Public Sub DisableColumn(ColumnName)
   
  ' If the column name already appears in the selected items list, then we can exit the sub
  If ItemInList(dlgColumnSetup.SelectedFields.wItemList,dlgColumnSetup.SelectedFields.wListSeparator,ColumnName) Then
    Call dlgColumnSetup.SelectedFields.DblClickItem(ColumnName)
  Else
    Log.Error("EnableColumn : the column """&ColumnName&""" was not available to disable")
  End If   
End Sub
  
' ------------------------------------------------------------------------------------
' Set the decimal places of a column
' ------------------------------------------------------------------------------------
Public Sub SetColumnDecimals(ColumnName,DecimalsValue)
    
  ' Only click on the ColumnName if it is available in the selected fields list box
  If ItemInList(dlgColumnSetup.SelectedFields.wItemList,dlgColumnSetup.SelectedFields.wListSeparator,ColumnName) Then
    Call dlgColumnSetup.SelectedFields.ClickItem(ColumnName)
  Else
    Log.Error("SetDecimals : column """&ColumnName&""" was not found in Selected fields list box")
    Exit Sub
  End If
     
  ' In case the the value passed in is not a string this converts it to one 
  DecimalsValue = ""&DecimalsValue&"" 
  
  Call dlgColumnSetup.Decimals.ClickItem(DecimalsValue)
End Sub

' ------------------------------------------------------------------------------------
' Get the decimal places
' ------------------------------------------------------------------------------------
Public Function GetColumnDecimals(ColumnName)
    
  ' Only click on the ColumnName if it is available in the selected fields list box
  If ItemInList(dlgColumnSetup.SelectedFields.wItemList,dlgColumnSetup.SelectedFields.wListSeparator,ColumnName) Then
    Call dlgColumnSetup.SelectedFields.ClickItem(ColumnName)
  Else
    Log.Error("GetDecimals : column """&ColumnName&""" was not found in Selected fields list box")
    Exit Function
  End If
  
  GetColumnDecimals = dlgColumnSetup.Decimals.wText
End Function
  
' ------------------------------------------------------------------------------------
' Set the multiplier
' ------------------------------------------------------------------------------------
Public Sub SetColumnMultiplier(ColumnName,MultiplierValue)
  
  ' Only click on the ColumnName if it is available in the selected fields list box
  If ItemInList(dlgColumnSetup.SelectedFields.wItemList,dlgColumnSetup.SelectedFields.wListSeparator,ColumnName) Then
    Call dlgColumnSetup.SelectedFields.ClickItem(ColumnName)
  Else
    Log.Error("SetMultiplier : column """&ColumnName&""" was not found in Selected fields list box")
  End If
     
  ' In case the the value passed in is not a string this converts it to one 
  MultiplierValue = ""&MultiplierValue&""
  
  Call dlgColumnSetup.Multiplier.ClickItem(MultiplierValue)
End Sub
  
' ------------------------------------------------------------------------------------
' Get the multiplier
' ------------------------------------------------------------------------------------
Public Function GetColumnMultiplier(ColumnName)
 
  ' Only click on the ColumnName if it is available in the selected fields list box
  If ItemInList(dlgColumnSetup.SelectedFields.wItemList,dlgColumnSetup.SelectedFields.wListSeparator,ColumnName) Then
    Call dlgColumnSetup.SelectedFields.ClickItem(ColumnName)
  Else
    Log.Error("SetMultiplier : column """&ColumnName&""" was not found in Selected fields list box")
    Exit Function
  End If

  GetColumnMultiplier = dlgColumnSetup.Multiplier.wText
End Function
  
Public Sub OpenColumnProperties
  OrderView.wndAfx.MainMenu.Click("View|Column Properties...")
End Sub
  
Public Sub OKColumnProperties
  dlgColumnSetup.btnApply.Click
  dlgColumnSetup.btnOK.Click
End Sub
  
Public Sub CancelColumnProperties
  dlgColumnSetup.btnCancel.Click
End Sub
  
Public Sub ApplyColumnProperties
  dlgColumnSetup.btnApply.Click
End Sub
  
Public Sub OpenFilters
  
  Call OrderView.wndAfx.MainMenu.Click("View|Filters...")
  Call WaitUntilAliasVisible(Aliases.OrderView,"dlgFilters",10000)
  
End Sub
  
Public Sub SetFiltersWorkingCheckbox(value)
  
  If value = True Then
    Call dlgFilters.checkWorking.ClickButton(cbChecked)  
  ElseIf value = False Then
    Call dlgFilters.checkWorking.ClickButton(cbUnChecked)
  Else
    Log.Error("SetFiltersWorkingCheckbox value is incorrect") 
    Exit Sub
  End If
End Sub
  
Public Sub SetFiltersFilledCheckbox(value)
 
  If value = True Then
    Call dlgFilters.checkFilled.ClickButton(cbChecked)  
  ElseIf value = False Then
    Call dlgFilters.checkFilled.ClickButton(cbUnChecked)
  Else
    Log.Error("SetFiltersFilledCheckbox value is incorrect") 
    Exit Sub
  End If
End Sub
  
Public Sub SetFiltersHeldCheckbox(value)
  
  If value = True Then
    Call dlgFilters.checkHeld.ClickButton(cbChecked)  
  ElseIf value = False Then
    Call dlgFilters.checkHeld.ClickButton(cbUnChecked)
  Else
    Log.Error("SetFiltersHeldCheckbox value is incorrect") 
    Exit Sub
  End If
  
End Sub
  
Public Sub SetFiltersOthersCheckbox(value)
  
  If value = True Then
    Call dlgFilters.checkOthers.ClickButton(cbChecked)  
  ElseIf value = False Then
    Call dlgFilters.checkOthers.ClickButton(cbUnChecked)
  Else
    Log.Error("SetFiltersOthersCheckbox value is incorrect") 
    Exit Sub
  End If
  End Sub
  
Public Sub SetFiltersMassQuoteCheckbox(value)
  
  If value = True Then
    Call dlgFilters.checkMassQuote.ClickButton(cbChecked)  
  ElseIf value = False Then
    Call dlgFilters.checkMassQuote.ClickButton(cbUnChecked)
  Else
    Log.Error("SetFiltersassQuotebox value is incorrect") 
    Exit Sub
  End If
End Sub
  
Public Sub SetFiltersUsers(value)
 
  If Not ItemInList(dlgFilters.Users.wSelectedItems,dlgFilters.Users.wListSeparator,value) Then
    Call dlgFilters.Users.ClickItem(value)
  End If

End Sub
  
Public Sub FiltersOK
    dlgFilters.btnOK.ClickButton
End Sub

'-----------------------------------------------------------------------------------------------------------------
Public Sub SetFiltersProduct(ProductType, Product)

  Call dlgFilters.btnAdd.Click
  Call WaitUntilAliasVisible(OrderView, "dlgProducts", 5000)
  
  Dim ProductsList
  Set ProductsList = OrderView.dlgProducts
  
  Call SelectProductSet(ProductsList,ProductType,Product)
    
End Sub


 
Public Sub FiltersCancel
    dlgFilters.btnCancel.ClickButton
End Sub

Class Order
  Dim OrderID
  Dim Theo
  Dim ProductName
  Dim ProductID
  Dim Account
  Dim BuySell
  Dim OrderStatus
  Dim Date
  Dim OrderType
  Dim Volume
  Dim ExecutedVolume
  Dim ResidualVolume
  Dim Price
  Dim Strike
  Dim User
  Dim Restriction
  Dim Month
  Dim Time
  Dim ASPrice
  Dim ACPrice
  Dim ACOffset
  Dim TOMPrice
  Dim TOMType
  Dim TOMNote
End Class

Sub Pull(OrderID)
  Dim OrderViewGrid, Row, Col, i
  Set OrderViewGrid = Aliases.OrderView.wndAfx.OrderViewGrid
  
  Row = GridControl.GetCellRow(OrderViewGrid.Handle, "Order ID", OrderID, 1)
    
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Order ID", 1)
  
  Call ClickGrid(OrderViewGrid, Row, Col,"Middle")
  
  Dim Count
  Count = 0
  Do
    Delay(TestConfig.fl_delay_OrderView_updates)
    Row = GridControl.GetCellRow(OrderViewGrid.Handle, "Order ID", OrderID, 1)
    Count = Count + 1
  Loop Until GetTextFromRow(OrderViewGrid, Row, "Order Status", 1) = "Pulled" Or Count = 50
    
  'Call CheckValue("Check order "&OrderID&" status is Pulled","Pulled", GetTextFromRow(OrderViewGrid, Row, "Order Status", 1))
End Sub

Function Amend(OrderID, OrderFilled)             
  
  Dim Row, Col, i 
   
  Row = GridControl.GetCellRow(OrderViewGrid.Handle, "Order ID", OrderID, 1) 
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Order ID", 1)
  'Debug
  'Log.Message("OrderView Row,Col " & Row & " ," & Col)
  
  Dim Col_OrderStatus, OrderStatus
  Col_OrderStatus = GridControl.GetCellColumn(OrderViewGrid.Handle, "Order Status", 1)
  OrderStatus = GridControl.GetCellText(OrderViewGrid.Handle, Row, Col_OrderStatus)
  
  
  Call TestUtilities.MakeCellVisible(OrderViewGrid, Row, Col)
    
  If Not Aliases.OrderView.WaitAliasChild("dlgAmendOrder", 200).Exists Then
    Call ClickGrid(OrderViewGrid, Row, Col,"Right")
    
    If OrderStatus = "Filled" Or OrderStatus = "Pulled" Or OrderStatus = "Rejected by risk" Then 
      Log.Message("The Order ID: """&OrderID&""" is a filled/Pulled/Rejected order and the test step is try to amend it")
      OrderFilled = True
    Else
      Log.Message("The Order ID: """&OrderID&""" is not a filled order")     
    End If
      
  End If                                         
    
End Function


Sub HighlightOrder(OrderID, FindOrder, Click)          
  
  Dim Row, Col, i  

  Row = GridControl.GetCellRow(OrderViewGrid.Handle, "Order ID", OrderID, 1) 
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle, "Order ID", 1)

  If Row = -1 Or Col = -1 Or Row = "" Or Col = "" Then
    Log.Error("The OrderID or one of the orderID for multiple orders cannot be found on the OrderView.")
    
    FindOrder = False        
           
    Exit Sub  
  End If
  
  FindOrder = True
    
  Call ClickGrid(OrderViewGrid, Row, Col, Click)    
  
End Sub
               


Public Sub SetPrice(Price, Method)
  
  If dlgAmendOrder.Exists <> True then
    Log.Message("The amend order dialog box does not appear in the OrderView.")
    Exit Sub
  End If


  If Method = "Keys" Then
    Call dlgAmendOrder.Price.Keys("[Home]![End][Del]"&Price)
  ElseIf Method = "Arrows" Then  
    If TestUtilities.FMod(Price, 1) <> 0 Then
      Log.Error("SetPrice - when using Arrows, the Price should be specified in whole number of clicks on the arrow")
      Exit Sub
    End If
    
    For i = 0 To Abs(Price) - 1
      If Price > 0 Then
        dlgAmendOrder.PriceArrow.Up
      Else
        dlgAmendOrder.PriceArrow.Down
      End If
    Next
  ' For MouseWheel events, the ACOffset should be specified in terms of the number of mouse wheel clicks
  ElseIf Method = "MouseWheel" Then
    If TestUtilities.FMod(Price, 1) <> 0 Then
      Log.Error("SetPrice - when using MouseWheel, the Price should be specified in whole number of notches to move the mouse wheel")
      Exit Sub
    End If
  
    dlgAmendOrder.Price.Click
    
    For i = 0 To Abs(Price) - 1
      If Price > 0 Then
        dlgAmendOrder.Price.MouseWheel(1)
      Else
        dlgAmendOrder.Price.MouseWheel(-1)
      End If
    Next
  Else
    Log.Error("SetPrice : unknown value for Method")
  End If    
End Sub
  
Public Sub SetQuantity(Quantity, Method)
  Dim CurrentQuantity, Count, dlgAmendOrder
  Set dlgAmendOrder = Aliases.OrderView.dlgAmendOrder

  If Method = "Keys" Then
    Call dlgAmendOrder.Quantity.Keys("[Home]![End][Del]"&Quantity)
  ElseIf Method = "Arrows" Then  
    If TestUtilities.FMod(Quantity, 1) <> 0 Then
      Log.Error("SetQuantity - when using Arrows, the Quantity should be specified in whole number of clicks on the arrow")
      Exit Sub
    End If
    
    For i = 0 To Abs(Quantity) - 1
      If Quantity > 0 Then
        dlgAmendOrder.QuantityArrow.Up
      Else
        dlgAmendOrder.QuantityArrow.Down
      End If
    Next
  ' For MouseWheel events, the ACOffset should be specified in terms of the number of mouse wheel clicks
  ElseIf Method = "MouseWheel" Then
    If TestUtilities.FMod(Quantity, 1) <> 0 Then
      Log.Error("SetQuantity - when using MouseWheel, the Quantity should be specified in whole number of notches to move the mouse wheel")
      Exit Sub
    End If
  
    dlgAmendOrder.Quantity.Click
    
    For i = 0 To Abs(Quantity) - 1
      If Quantity > 0 Then
        dlgAmendOrder.Quantity.MouseWheel(1)
      Else
        dlgAmendOrder.Quantity.MouseWheel(-1)
      End If
    Next
  ElseIf Method = "Buttons" Then
    If FMod(Quantity, 1) <> 0 Then
    Log.Error("SetQuantity : Cannot set quantity using buttons as it is not a multiple of 1 ("&Quantity&")")
    Exit Sub
    End If
    
    Count = 0
    dlgAmendOrder.btnC.ClickButton
    dlgAmendOrder.Quantity.Keys("0")
    
    Do Until StrToInt(dlgAmendOrder.Quantity.wText) = Quantity Or Count = 100
    CurrentQuantity = StrToInt(dlgAmendOrder.Quantity.wText)
    If (Quantity-CurrentQuantity) >= 100 Then
      dlgAmendOrder.btn100.ClickButton
    ElseIf (Quantity-CurrentQuantity) >= 50 Then
      dlgAmendOrder.btn50.ClickButton
    ElseIf (Quantity-CurrentQuantity) >= 10 Then
      dlgAmendOrder.btn10.ClickButton
    ElseIf (Quantity-CurrentQuantity) >= 5 Then
      dlgAmendOrder.btn5.ClickButton
    ElseIf (Quantity-CurrentQuantity) >= 0 Then
      dlgAmendOrder.btn1.ClickButton
    End If
    Count = Count + 1
    Loop
  Else
    Log.Error("SetQuantity : unknown value for Method")
  End If
End Sub


' ------------------------------------------------------------------------------------
' Get the price from the order ticket
' ------------------------------------------------------------------------------------
Public Function GetPrice  
  Dim dlgAmendOrder
  Set dlgAmendOrder = Aliases.OrderView.dlgAmendOrder

  GetPrice = StrToFloat(dlgAmendOrder.Price.wText)  
End Function
  
' ------------------------------------------------------------------------------------
' Get the quantity from the order ticket
' ------------------------------------------------------------------------------------
Public Function GetQuantity  
  Dim dlgAmendOrder
  Set dlgAmendOrder = Aliases.OrderView.dlgAmendOrder

  GetQuantity = StrToInt(dlgAmendOrder.Quantity.wText)  
End Function

Public Sub SetACOffset(ACOffset, Method)
  Dim CurrentQuantity, Count
  
  Dim dlgAmendOrder
  Set dlgAmendOrder = Aliases.OrderView.dlgAmendOrder
  
  If Aliases.MarketView.dlgAmendOrder.Offset.Exists = False Then
    Log.Error("SetACOffset : Cannot set AC offset as the AC offset field is not enabled")
    Exit Sub
  End If

  If Aliases.MarketView1.dlgAmendOrder.Offset.Exists = False Then
    Log.Error("SetACOffset : Cannot set AC offset as the AC offset field is not enabled")
    Exit Sub
  End If
  
  
  If Method = "Keys" Then
    Call dlgAmendOrder.Offset.Keys("[Home]![End][Del]"&ACOffset)
  ElseIf Method = "Arrows" Then  
    If TestUtilities.FMod(ACOffset, 1) <> 0 Then
      Log.Error("SetACOffset - when using Arrows, the ACOffset should be specified in whole number of clicks on the arrow")
      Exit Sub
    End If
    
    For i = 0 To Abs(ACOffset) - 1
      If ACOffset > 0 Then
        dlgAmendOrder.OffsetArrow.Up
      Else
        dlgAmendOrder.OffsetArrow.Down
      End If
    Next
  ' For MouseWheel events, the ACOffset should be specified in terms of the number of mouse wheel clicks
  ElseIf Method = "MouseWheel" Then
    If TestUtilities.FMod(ACOffset, 1) <> 0 Then
      Log.Error("SetACOffset - when using MouseWheel, the ACOffset should be specified in whole number of notches to move the mouse wheel")
      Exit Sub
    End If
  
    dlgAmendOrder.Offset.Click
    
    For i = 0 To Abs(ACOffset) - 1
      If ACOffset > 0 Then
        dlgAmendOrder.Offset.MouseWheel(1)
      Else
        dlgAmendOrder.Offset.MouseWheel(-1)
      End If
    Next
  Else
    Log.Error("SetACOffset : unknown value for Method")
  End If    
End Sub
  
Public Function GetACOffset
  Dim dlgAmendOrder
  Set dlgAmendOrder = Aliases.OrderView.dlgAmendOrder
  GetACOffset = dlgAmendOrder.Offset.wText
End Function

Sub ClickToolbar(Button)
  Dim MenuBar
  Set MenuBar = Aliases.OrderView.wndAfx.Dockbar.Menubar

  Select Case Button
  Case "New"
    Call Menubar.Click(15, 12)  ' New
  Case "Open"
    Call Menubar.Click(35, 12)  ' Open
  Case "Save"  
    Call Menubar.Click(63, 12)  ' Save
  Case "Save all"  
    Call Menubar.Click(84, 12)  ' Save all
  Case "Zoom In"    
    Call Menubar.Click(152, 12) ' Zoom in
  Case "Zoom Out"
    Call Menubar.Click(175, 12) ' Zoom out
  Case "Sorting"
    Call Menubar.Click(209, 12) ' Sorting
  Case "Sorting"
    Call Menubar.Click(221, 12) ' Filters
  Case "Unhold"
    Call Menubar.Click(258, 12) ' Unhold
  Case "Auto-unhold"
    Call Menubar.Click(278, 12) ' Auto unhold
  Case "Pull"
    Call Menubar.Click(308, 12) ' Pull selected order
  Case "Pull All"
    Call Menubar.Click(328, 12) ' Pull all orders
  Case "Report"
    Call Menubar.Click(355, 12) ' Report selected order
  Case Else
    Log.Error("ClickMenubar: Unknown value for Button")
  End Select
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------------------------------------------------
Public Sub OpenViewSetup
  OrderView.wndAfx.MainMenu.Click("View|View Setup...")
End Sub

Public Sub SetSorting(Column)

  'Open up the Sorting properties  
  OrderView.wndAfx.MainMenu.Click("View|Sorting...")
  
  Dim dlgSorting
  Set dlgSorting = Aliases.OrderView.dlgSorting
  
  ' If the column name already appears in the selected items list, then we can exit the sub
  'If ItemInList(dlgColumnSetup.SelectedFields.wItemList,dlgColumnSetup.SelectedFields.wListSeparator,ColumnName) Then
   ' Exit Sub
  'End If
  
  ' Otherwise, see if the column is available to select and if it is double click on it to move it to selected items
  'If ItemInList(dlgColumnSetup.AvailableFields.wItemList,dlgColumnSetup.AvailableFields.wListSeparator,ColumnName) Then
   ' Call dlgColumnSetup.AvailableFields.DblClickItem(ColumnName)
  'Else
   ' Log.Error("EnableColumn : the column """&ColumnName&""" was not available to enable")
  'End If 
  
  '  If ItemInList(dlgSorting.ListBox.wItemList,dlgSorting.ListBox.wListSeparator,Column) Then
'    Call dlgSorting.ListBox.ClickItem(Column)
'    Call dlgSorting.btnAdd.Click
'  Else
'    Log.Error("Somethign Very Bad Happened")
'  End if

  While  dlgSorting.ListBox1.wItemCount > 0   
    Call dlgSorting.ListBox1.ClickItemXY(0, 32, 9)
    dlgSorting.btnRemove.ClickButton
  Wend  
 
  Call dlgSorting.ListBox.ClickItem("Time")
  dlgSorting.btnAdd.ClickButton
  dlgSorting.btnOK.ClickButton
  
End Sub


'-------------------------------------------------------------------------------------------------------------------------
' Enable columns
'-------------------------------------------------------------------------------------------------------------------------
Public Sub EnableColumns(ColumnNames)

'MS - I reckon it's better to pass an Array that has a list of columns and then go through each one
  
  For Each x in ColumnNames   
    ' If the column name already appears in the selected items list, then we can exit the sub
 '    If ItemInList(dlgColumnSetup.SelectedFields.wItemList,dlgColumnSetup.SelectedFields.wListSeparator,x) Then
'      Exit Sub
'    End If
  
    ' Otherwise, see if the column is available to select and if it is double click on it to move it to selected items
    If ItemInList(dlgColumnSetup.AvailableFields.wItemList,dlgColumnSetup.AvailableFields.wListSeparator,x) Then
      Call dlgColumnSetup.AvailableFields.DblClickItem(x)
'    Else
'      Log.Error("EnableColumn : the column """&x&""" was not available to enable")
    End If   
  
  Next
'    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Function: OVOrderInstance
'-------------------------------------------------------------------------------------------------------------------------
Function OVOrderClassInstance

  Set OVOrderClassInstance = New Order

End Function

'-------------------------------------------------------------------------------------------------------------------------
'Function: GetColumnValue
'-------------------------------------------------------------------------------------------------------------------------
Function GetColumnValue(OrderID, ColumnName)
 
  Dim Row, Col
  
  'Dim MyOrder
  'Set MyOrder = New Order                                                          

  'Dim OrderView
  'Set OrderView = Aliases.OrderView
  
  Row = GridControl.GetCellRow(OrderViewGrid.Handle, "Order ID",OrderID,1)
  
  Col = GridControl.GetCellColumn(OrderViewGrid.Handle,ColumnName,1)
  
  GetColumnValue = GridControl.GetCellText(OrderViewGrid.Handle,Row,Col)
  

End Function

Public Sub PressAmend
  
  If dlgAmendOrder.Exists <> True then
    Log.Message("The amend order dialog box does not appear in the OrderView.")
  Else  
    dlgAmendOrder.btnSubmit.Click
  End If   

End Sub

Public Sub PressCancel
  
  'Dim dlgAmendOrder
  'Set dlgAmendOrder = Aliases.OrderView.dlgAmendOrder
  
  dlgAmendOrder.btnCancel.Click
  
End Sub

Sub PullAllOrders

  Dim  OrderViewMenuBar

  Set OrderViewMenuBar = Aliases.OrderView.wndAfx.OrderViewDockBar.OrderViewMenuBar
  Call OrderViewMenuBar.Click(330, 20) 
  
End sub
