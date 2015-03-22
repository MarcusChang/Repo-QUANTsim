'USEUNIT TestUtilities
'USEUNIT OrderView
'USEUNIT MarketView
'USEUNIT MarketView1
'USEUNIT Order

Private GridControl, OptionViewGrid, OrderViewGrid
Private FsaValue
Private HeldValue

Dim TestConfig
Set TestConfig = ProjectSuite.Variables

Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl
Set OptionViewGrid = Aliases.MarketView1.wndAfx.MDIClient.OptionView1.OptionViewGrid
Set OrderViewGrid = Aliases.OrderView.wndAfx.OrderViewGrid

'Private OrderView
'Set OrderView = Aliases.OrderView

Public Sub Initialize
  Set GridControl = TestConfig.QuantCOREControl
  Set OptionViewGrid = Aliases.MarketView1.wndAfx.MDIClient.OptionView1.OptionViewGrid
  Set OrderViewGrid = Aliases.OrderView.wndAfx.OrderViewGrid
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Class: MarketViewOrder
' Description
'-------------------------------------------------------------------------------------------------------------------------
Class MarketViewOrder
  
  Dim OrderName
  Dim ProductType
  Dim PriceFormula
  Dim Quantity
  Dim BidAsk
  Dim Month
  Dim OrderStatus

  Dim Series
  Dim Strike
  Dim OrderRestriction
  Dim Price
  Dim PriceBefore 'Original price before an amendment
  Dim OrderID
  Dim Strategy
  Dim StrategyType
  Dim StrategyName
  Dim StrategyID
  Dim Held
  Dim AnotherSameGroupUser
  
  
    
  'Create the ProductID based on ProductType, Product, and Long Month e.g. SIM.F.XJO.OCT2011 
  Public Property Get ProductID
  
  Dim Row, Col
    Select Case ProductType
      Case "Future"
          'Note the month has to be written in LONG format in the spreadsheet
        ProductID = "SIM.F." &TestConfig.FutureProduct &"." &Month
      Case "Strategy"
          'Note the month has to be written in LONG format in the spreadsheet
        ProductID = StrategyID 
      Case "Call"
        'Note the month has to be written in LONG format in the spreadsheet
        ProductID = "SIM.O." &TestConfig.OptionProduct &"." &Month &"." &Strike &".C.0"
      Case "Put"
        'Note the month has to be written in LONG format in the spreadsheet
        ProductID = "SIM.O." &TestConfig.OptionProduct &"." &Month &"." &Strike &".P.0"  
      Case "TMCStrategy"
        ProductID = StrategyID
            
      Case Else
        Log.Error("Invalid ProductType Specified from Worksheet")            
    End Select
  End Property
  
  'Set the OrderTickSize baded on ProductType, need to add cases for Equities, Strategies
  Public Property Get OrderTickSize
    Select Case ProductType
      Case "Future"
        OrderTickSize = TestConfig.FutureTickSize
      Case "Call"
        OrderTickSize = TestConfig.OptionTickSize
      Case "Put"
        OrderTickSize = TestConfig.OptionTickSize 
      Case "Strategy", "TMCStrategy"
        OrderTickSize = TestConfig.OptionTickSize
      
      Case Else
        Log.Error("Invalid ProductType Specified from Worksheet")            
    End Select
  End Property
  
End Class


' ------------------------------------------------------------------------------------
' Returns a reference for the type of order ticket that has opened
' ------------------------------------------------------------------------------------
'Private Function GetOrderTicket
Public Function GetOrderTicket 
	On Error Resume Next
  If Aliases.MarketView1.WaitAliasChild("dlgOrderTicket").Exists Then
    Set GetOrderTicket = Aliases.MarketView1.dlgOrderTicket
  ElseIf Aliases.MarketView1.WaitAliasChild("dlgAmendOrder").Exists Then
    Set GetOrderTicket = Aliases.MarketView1.dlgAmendOrder
  Else
    Log.Error("GetOrderTicket : Order ticket is not on screen")
    Exit Function
  End If
End Function
  
' ------------------------------------------------------------------------------------
' Open a new ticket order form
' ------------------------------------------------------------------------------------
Public Sub Ticket(OptionViewGrid, Row, Col)
  ' Reset the value for held state when we open a new ticket
  HeldValue = False
  
  'Call MarketView.MakeCellVisible(OptionViewGrid, Row, Col)
  Call TestUtilities.MakeCellVisible(OptionViewGrid, Row, Col)
  
  Call ClickGrid(OptionViewGrid, Row, Col, "Right")
  
  On Error Resume Next  
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  If dlgOrderTicket.WndCaption <> "Order Ticket" Then
    'Call Log.Warning("The order ticket that popped up was not a new order ticket.  It has the caption """&dlgOrderTicket.WndCaption&""".")
  End If
End Sub
  
' ------------------------------------------------------------------------------------
' Click trade an order
' ------------------------------------------------------------------------------------
Public Sub Click(OptionViewGrid, Row, Col)
  Call ClickGrid(OptionViewGrid, Row, Col, "Left")
End Sub
  
' ------------------------------------------------------------------------------------
' Pull an order at given row / column
' ------------------------------------------------------------------------------------
Public Sub Pull(OptionViewGrid, Row, Col)
  'MS - Added 24/10/11 
  Call TestUtilities.MakeCellVisible(OptionViewGrid, Row, Col)
  Call ClickGrid(OptionViewGrid, Row, Col, "Middle")
End Sub
  
Public Sub TickOrderBetter(OptionViewGrid, Row, Col)
  Call ClickGrid(OptionViewGrid, Row, Col, "Fourth")
End Sub
  
Public Sub TickOrderWorse(OptionViewGrid, Row, Col)
  Call ClickGrid(OptionViewGrid, Row, Col, "Fifth")
End Sub
  
Public Sub Amend(OptionViewGrid, Row, Col)
  ' Reset the value for held state when we open a new ticket
  HeldValue = False
  
  Call ClickGrid(OptionViewGrid, Row, Col, "Right")
  
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  If dlgOrderTicket.WndCaption <> "Amend Order" Then
    Call Log.Warning("The order ticket that popped up was not an amend order ticket.  It has the caption """&dlgOrderTicket.WndCaption&""".")
  End If
End Sub
  
' ------------------------------------------------------------------------------------
' Set the price on the order ticket
' Can be set by using arrows, or keyboard presses
' ------------------------------------------------------------------------------------
Public Sub SetPrice(Price, Method)

  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket

  If Method = "Keys" Then
    Call dlgOrderTicket.Price.Click
    Call dlgOrderTicket.Price.Keys("[Home]![End][Del]"&Price)
  ElseIf Method = "Arrows" Then
    If FMod(Price,GetArrowObjectStepSize(dlgOrderTicket.PriceArrow,dlgOrderTicket.Price)) <> 0 Then
    Log.Error("SetPrice : Cannot set price using arrows as it is not on a tick boundary ("&Price&")")
    Exit Sub
    End If
    
    Dim Count
    Count = 0
    Do Until StrToInt(dlgOrderTicket.Price.wText) = Price Or Count = 10000
    If Price > StrToInt(dlgOrderTicket.Price.wText) Then
      dlgOrderTicket.PriceArrow.Up
    Else
      dlgOrderTicket.PriceArrow.Down
    End If
    Loop
  Else
    Log.Error("SetPrice : unknown value for Method")
  End If    
End Sub
  
' ------------------------------------------------------------------------------------
' Set the quantity on the order ticket
' Can be set by using buttons, arrows, or keyboard presses
' ------------------------------------------------------------------------------------
Public Sub SetQuantity(Quantity, Method)
  Dim CurrentQuantity, Count
  
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket

  If Method = "Keys" Then
    Call dlgOrderTicket.Quantity.Click
    Call dlgOrderTicket.Quantity.Keys("[Home]![End][Del]"&Quantity)
  ElseIf Method = "Buttons" Then
    If FMod(Quantity, 1) <> 0 Then
    Log.Error("SetQuantity : Cannot set quantity using buttons as it is not a multiple of 1 ("&Quantity&")")
    Exit Sub
    End If
    
    Count = 0
    dlgOrderTicket.btnC.ClickButton
    dlgOrderTicket.Quantity.Keys("0")
    
    Do Until StrToInt(dlgOrderTicket.Quantity.wText) = Quantity Or Count = 100
    CurrentQuantity = StrToInt(dlgOrderTicket.Quantity.wText)
    If (Quantity-CurrentQuantity) >= 100 Then
      dlgOrderTicket.btn100.ClickButton
    ElseIf (Quantity-CurrentQuantity) >= 50 Then
      dlgOrderTicket.btn50.ClickButton
    ElseIf (Quantity-CurrentQuantity) >= 10 Then
      dlgOrderTicket.btn10.ClickButton
    ElseIf (Quantity-CurrentQuantity) >= 5 Then
      dlgOrderTicket.btn5.ClickButton
    ElseIf (Quantity-CurrentQuantity) >= 0 Then
      dlgOrderTicket.btn1.ClickButton
    End If
    Count = Count + 1
    Loop
  ElseIf Method = "Arrows" Then
    If FMod(Quantity, 1) <> 0 Then
    Log.Error("SetQuantity : Cannot set quantity using arrows as it is not a multiple of 1 ("&Quantity&")")
    Exit Sub
    End If
    
    Count = 0
    Do Until StrToInt(dlgOrderTicket.Quantity.wText) = Quantity Or Count = 1000
    If Quantity > StrToInt(dlgOrderTicket.Quantity.wText) Then
      dlgOrderTicket.QuantityArrow.Up
    Else
      dlgOrderTicket.QuantityArrow.Down
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
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket

  GetPrice = StrToFloat(dlgOrderTicket.Price.wText)  
End Function
  
' ------------------------------------------------------------------------------------
' Get the quantity from the order ticket
' ------------------------------------------------------------------------------------
Public Function GetQuantity  
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket

  GetQuantity = StrToInt(dlgOrderTicket.Quantity.wText)  
End Function
  
' ------------------------------------------------------------------------------------
' Switch the buy sell indicator
' ------------------------------------------------------------------------------------
Public Sub SwitchBuySell
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  dlgOrderTicket.btnBuySell.Click
End Sub
  
' ------------------------------------------------------------------------------------
' Get the product ID from the order ticket         
' ------------------------------------------------------------------------------------ 
Public Function GetProductID
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  GetProductID = dlgOrderTicket.ProductID
  End Function
  
' ------------------------------------------------------------------------------------
' Sets the state of the H button in the order ticket                 
' ------------------------------------------------------------------------------------ 
Public Sub SetHeld(Value)
  If Value <> True And Value <> False Then
    Log.Error("SetHeld: Invalid value for Value "&Value)
    Exit Sub
  End If
  
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
    
  If dlgOrderTicket.Held.Enabled Then
    If Value = True And HeldValue = False Then
    dlgOrderTicket.Held.Click
    ElseIf Value = False And HeldValue = True Then
    dlgOrderTicket.Held.Click
    End If
  Else
    Log.Message("SetHeld : Did not click held button as it was disabled")
  End If
End Sub
  
' ------------------------------------------------------------------------------------
' Sets the state of the F button in the order ticket                 
' ------------------------------------------------------------------------------------ 
Public Sub SetFsaIndicator(FsaValue) 
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  ' Check to see if the FSA tab is active
  If Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("FsaTab", 50).Visible And FsaValue = False Then
    dlgOrderTicket.btnF.Click
  ElseIf Not Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("FsaTab", 50).Visible And FsaValue = True Then
    dlgOrderTicket.btnF.Click
   End If
End Sub
  
' ------------------------------------------------------------------------------------
' Sets the state of the AC button in the order ticket                 
' ------------------------------------------------------------------------------------ 
Public Sub SetAC(ACEnabled)
  ' Get reference to the type of open order ticket  
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  If ACEnabled = False Then
    If Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("AOM",50).Visible = True Then
      If dlgOrderTicket.AOM.Offset.Enabled = True Then
        dlgOrderTicket.btnAC.Click
      End If
    End If
  ElseIf ACEnabled = True Then
    If Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("AOM",50).Visible = False Then
      dlgOrderTicket.btnAC.Click
    ElseIf Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("AOM",50).Visible = True And dlgOrderTicket.AOM.Offset.Enabled = False Then
      dlgOrderTicket.btnAC.Click
    End If
  End If
End Sub
  
' ------------------------------------------------------------------------------------
' Sets the state of the AS button in the order ticket                 
' ------------------------------------------------------------------------------------   
Public Sub SetAS(ASEnabled)
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  If ASEnabled = False Then
    If Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("AOM", 50).Visible = True Then
      If dlgOrderTicket.AOM.Trigger.Enabled = True  Then
        dlgOrderTicket.btnAS.Click
      End If
    End If
  ElseIf ASEnabled = True And Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("AOM", 50).Visible = False Then
    dlgOrderTicket.btnAS.Click
  ElseIf ASEnabled = True And Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("AOM", 50).Visible = True And dlgOrderTicket.AOM.Trigger.Enabled = False Then
    dlgOrderTicket.btnAS.Click
  End If
End Sub
  
Public Sub SetACOffset(ACOffset, Method)
  Dim CurrentQuantity, Count, i
  
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  If dlgOrderTicket.WaitAliasChild("AOM",50).Visible = False Or dlgOrderTicket.AOM.Offset.Enabled = False Then
    Log.Error("SetACOffset : Cannot set AC offset as the AC offset field is not enabled")
    Exit Sub
  End If

  If Method = "Keys" Then
    Call dlgOrderTicket.AOM.Offset.Keys("[Home]![End][Del]"&ACOffset)
  ElseIf Method = "Arrows" Then  
    If TestUtilities.FMod(ACOffset, 1) <> 0 Then
      Log.Error("SetACOffset - when using Arrows, the ACOffset should be specified in whole number of clicks on the arrow")
      Exit Sub
    End If
    
    For i = 0 To Abs(ACOffset) - 1
      If ACOffset > 0 Then
        dlgOrderTicket.AOM.OffsetArrow.Up
      Else
        dlgOrderTicket.AOM.OffsetArrow.Down
      End If
    Next
  ' For MouseWheel events, the ACOffset should be specified in terms of the number of mouse wheel clicks
  ElseIf Method = "MouseWheel" Then
    If TestUtilities.FMod(ACOffset, 1) <> 0 Then
      Log.Error("SetACOffset - when using MouseWheel, the ACOffset should be specified in whole number of notches to move the mouse wheel")
      Exit Sub
    End If
  
    Aliases.MarketView1.dlgOrderTicket.AOM.Offset.Click
    
    For i = 0 To Abs(ACOffset) - 1
      If ACOffset > 0 Then
        Aliases.MarketView1.dlgOrderTicket.AOM.Offset.MouseWheel(1)
      Else
        Aliases.MarketView1.dlgOrderTicket.AOM.Offset.MouseWheel(-1)
      End If
    Next
  Else
    Log.Error("SetACOffset : unknown value for Method")
  End If    
End Sub
  
Public Function GetACOffset
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  GetACOffset = dlgOrderTicket.AOM.Offset.wText
End Function
  
Public Sub SetASTrigger(ASTrigger, Method)
  Dim CurrentQuantity, Count, i
  
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  If Aliases.MarketView1.dlgOrderTicket.WaitAliasChild("AOM",50).Visible = False Or dlgOrderTicket.AOM.Trigger.Enabled = False Then
    Log.Error("SetACOffset : Cannot set AS offset as the AS trigger field is not enabled")
    Exit Sub
  End If
  
  If dlgOrderTicket.WndCaption = "Amend Order" Then
    Log.Error("SetASTrigger : cannot set the AS trigger value on an order amendement ticket")
    Exit Sub
  End If

  If Method = "Keys" Then
    Call dlgOrderTicket.AOM.Trigger.Keys("[Home]![End][Del]"&ASTrigger)
  ElseIf Method = "Arrows" Then
    If TestUtilities.FMod(ASTrigger, 1) <> 0 Then
      Log.Error("SetASTrigger - when using Arrows, the ASTrigger should be specified in whole number of clicks on the arrow")
      Exit Sub
    End If
    
    For i = 0 To Abs(ASTrigger) - 1
      If ASTrigger > 0 Then
        Aliases.MarketView1.dlgOrderTicket.AOM.TriggerArrow.Up
      Else
        Aliases.MarketView1.dlgOrderTicket.AOM.TriggerArrow.Down
      End If
    Next
  ' For MouseWheel events, the ACOffset should be specified in terms of the number of mouse wheel clicks
  ElseIf Method = "MouseWheel" Then
    If TestUtilities.FMod(ASTrigger, 1) <> 0 Then
      Log.Error("SetASTrigger - when using MouseWheel, the AS Trigger should be specified in whole number of notches to move the mouse wheel")
      Exit Sub
    End If
  
    Aliases.MarketView1.dlgOrderTicket.AOM.Trigger.Click
    
    For i = 0 To Abs(ASTrigger) - 1
      If ASTrigger > 0 Then
        Aliases.MarketView1.dlgOrderTicket.AOM.Trigger.MouseWheel(1)
      Else
        Aliases.MarketView1.dlgOrderTicket.AOM.Trigger.MouseWheel(-1)
      End If
    Next
  Else
    Log.Error("SetASTrigger : unknown value for Method")
  End If    
End Sub
  
Public Function GetASTrigger
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  GetASTrigger = dlgOrderTicket.AOM.Trigger.wText
End Function

Function Submit(CancelDialog)
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket
  
  Dim OldOrderIDs
  OldOrderIDs = OrderView.GetCurrentOrderIDs
            
  Dim OrderIDCount
  OrderIDCount = GridControl.GetRowCount(OrderViewGrid.Handle)
  
  dlgOrderTicket.btnSubmit.ClickButton
  
  ' In case the warning "Theo not available.  Are you sure you want to submit these orders?" appears
  ' We click "Yes" on the dialog box
  If Aliases.MarketView1.WaitAliasChild("dlgOrderWarning").Exists Then
    Dim dlgOrderWarning
    Set dlgOrderWarning = Aliases.MarketView1.dlgOrderWarning
    Call Aliases.MarketView1.dlgOrderWarning.btnYes.ClickButton    
  End If
  
  Dim Count
  Count = 0
  While OrderIDCount = GridControl.GetRowCount(OrderViewGrid.Handle) And Count < 50
    Count = Count + 1
    Delay(100)
  WEnd
  
  Dim CurrentOrderIDs
  CurrentOrderIDs = OrderView.GetCurrentOrderIDs
  
  If dlgOrderTicket.WndCaption = "Order Ticket" Then
    If UBound(NewItems(OldOrderIDs, CurrentOrderIDs)) = 0 Then
      Submit = NewItems(OldOrderIDs, CurrentOrderIDs)(0)
      
      Log.Message("Order ID = "&NewItems(OldOrderIDs, CurrentOrderIDs)(0))
    Else
      Log.Error("Submit : Did not get 1 order ID returned... size of NewItems = "&UBound(NewItems(OldOrderIDs, CurrentOrderIDs)))
    End If
  End If

  If CancelDialog = True Then
    dlgOrderTicket.btnCancel.ClickButton
  End If
End Function
  
Public Sub Cancel
  Dim dlgOrderTicket
  Set dlgOrderTicket = GetOrderTicket

  dlgOrderTicket.btnCancel.ClickButton
End Sub
  
' This function gets the change in value applied to a cell for one click of 
' its associated arrow key
Private Function GetArrowObjectStepSize(ArrowObject, ArrowObjectValue)
  ' Get reference value
  ArrowObject.Up
  GetArrowObjectStepSize = StrToFloat(ArrowObjectValue.wText)
  
  ' Get value for one step up
  ArrowObject.Up
  GetArrowObjectStepSize = StrToFloat(ArrowObjectValue.wText) - GetArrowObjectStepSize
  
  ' Go back to original value
  ArrowObject.Down
  ArrowObject.Down
End Function
  
' Returns Order ID if successful
Function SubmitOrder(GridObject, ProductID, CallPut, BidAsk, Price, Quantity, Held)
  Dim Row, Col, Instance

  ' Work out the instance of the column
  Select Case CallPut
  Case "Call", "Future", "Strategy"
    Instance = 1
  Case "Put"
    Instance = 2
  Case Else
    Log.Error("SubmitOrder : invalid value for CallPut - "&CallPut)
  End Select
  
  ' Check BidAsk is valid
  If Not ItemInList("Bid|Bid Qty|Ask|Ask Qty","|",BidAsk) Then
    Log.Error("SubmitOrder : invalid value for BidAsk - "&BidAsk)
  End If
  
  Log.Message("Submit order: "&" "&ProductID&" "&CallPut&" "&Replace(BidAsk," Qty","")&" "&Quantity&" @ "&Price)
  
  Row = GridControl.GetCellRow(GridObject.Handle, "ProductID", ProductID, 1)
  Col = GridControl.GetCellColumn(GridObject.Handle, BidAsk, Instance)
   
  'MS - This is where the values are being entered into the order ticket
  Call Order1.Ticket(GridObject, Row, Col)
  Call Order1.SetQuantity(Quantity,"Keys")
  Call Order1.SetPrice(Price,"Keys")
  Call Order1.SetAC(False)
  Call Order1.SetAS(False)
  Call Order1.SetHeld(Held)

  SubmitOrder = Order1.Submit(True)
End Function

' Returns Order ID if successful
Function SubmitACOrder(GridObject, ProductID, CallPut, BidAsk, Price, Quantity, Offset, Held)
  Dim Row, Col, Instance

  ' Work out the instance of the column
  Select Case CallPut
  Case "Call", "Future", "Strategy"
    Instance = 1
  Case "Put"
    Instance = 2
  Case Else
    Log.Error("SubmitACOrder : invalid value for CallPut - "&CallPut)
  End Select
  
  ' Check BidAsk is valid
  If Not ItemInList("Bid|Bid Qty|Ask|Ask Qty","|",BidAsk) Then
    Log.Error("SubmitACOrder : invalid value for BidAsk - "&BidAsk)
  End If
  
  Log.Message("Submit AC order: "&" "&ProductID&" "&CallPut&" "&Replace(BidAsk," Qty","")&" "&Quantity&" @ "&Price&" with offset of "&Offset)
  
  Row = GridControl.GetCellRow(GridObject.Handle, "ProductID", ProductID, 1)
  Col = GridControl.GetCellColumn(GridObject.Handle, BidAsk, Instance)
   
  Call Order1.Ticket(GridObject, Row, Col)
  Call Order1.SetQuantity(Quantity,"Keys")
  Call Order1.SetPrice(Price,"Keys")
  Call Order1.SetAC(True)
  Call Order1.SetACOffset(Offset,"Keys")
  Call Order1.SetHeld(Held)
  
  SubmitACOrder = Order1.Submit(True)
End Function

' Returns Order ID if successful
Function SubmitASOrder(GridObject, ProductID, CallPut, BidAsk, Restriction, Price, Quantity, Trigger)
  Dim Row, Col, Instance

  ' Work out the instance of the column
  Select Case CallPut
  Case "Call", "Future", "Strategy"
    Instance = 1
  Case "Put"
    Instance = 2
  Case Else
    Log.Error("SubmitASOrder : invalid value for CallPut - "&CallPut)
  End Select
  
  ' Check BidAsk is valid
  If Not ItemInList("Bid|Bid Qty|Ask|Ask Qty","|",BidAsk) Then
    Log.Error("SubmitASOrder : invalid value for BidAsk - "&BidAsk)
  End If
  
  Log.Message("Submit AS order: "&" "&ProductID&" "&CallPut&" "&Replace(BidAsk," Qty","")&" "&Quantity&" @ "&Price&" with trigger of "&Trigger)

  Row = GridControl.GetCellRow(GridObject.Handle, "ProductID", ProductID, 1)
  Col = GridControl.GetCellColumn(GridObject.Handle, BidAsk, Instance)

  Call Order1.Ticket(GridObject, Row, Col)
  Call Order1.SetQuantity(Quantity,"Keys")
  Call Order1.SetPrice(Price,"Keys")
  Call Order1.SetAS(True)
  Call Order1.SetRestriction(Restriction)
  Call Order1.SetASTrigger(Trigger,"Keys")
  
  SubmitASOrder = Order1.Submit(True)
End Function

' Returns Order ID if successful
Function SubmitASCOrder(GridObject, ProductID, CallPut, BidAsk, Price, Quantity, Trigger, Offset)
  Dim Row, Col, Instance

  ' Work out the instance of the column
  Select Case CallPut
  Case "Call", "Future", "Strategy"
    Instance = 1
  Case "Put"
    Instance = 2
  Case Else
    Log.Error("SubmitASCOrder : invalid value for CallPut - "&CallPut)
  End Select
  
  ' Check BidAsk is valid
  If Not ItemInList("Bid|Bid Qty|Ask|Ask Qty","|",BidAsk) Then
    Log.Error("SubmitASCOrder : invalid value for BidAsk - "&BidAsk)
  End If
  
  Log.Message("Submit ASC order: "&" "&ProductID&" "&CallPut&" "&Replace(BidAsk," Qty","")&" "&Quantity&" @ "&Price&" with trigger of "&Trigger&" and offset of "&Offset)
  
  Row = GridControl.GetCellRow(GridObject.Handle, "ProductID", ProductID, 1)
  Col = GridControl.GetCellColumn(GridObject.Handle, BidAsk, Instance)
  
  Call Order1.Ticket(GridObject, Row, Col)
  Call Order1.SetQuantity(Quantity,"Keys")
  Call Order1.SetPrice(Price,"Keys")
  Call Order1.SetAC(True)
  Call Order1.SetAS(True)
  Call Order1.SetACOffset(Offset,"Keys")
  Call Order1.SetASTrigger(Trigger,"Keys")
  
  SubmitASCOrder = Order1.Submit(True)
End Function



Sub SetRestriction(Restriction)
  Dim dlgOrderTicket
  Set dlgOrderTicket = Aliases.MarketView1.dlgOrderTicket
  
  Select Case Restriction
  Case "GFD"
    dlgOrderTicket.radioDay.ClickButton
  Case "IOC"
    dlgOrderTicket.radioIOC.ClickButton
  Case "Open"
    dlgOrderTicket.radioOpen.ClickButton
  Case "Ice/Close"
    dlgOrderTicket.radioIceClose.ClickButton
  Case Else
    Log.Error("SetRestriction : unknown value for Restriction")
  End Select
End Sub



Sub ClickOrderForTrade(TheCurrentOrder)        
  Dim Row_Click, Col_Click
  
  Select Case TheCurrentOrder.ProductType
  Case "Future" 
    Instance = 1
  Case "Strategy"
    Instance = 1    
  Case "Call"
    Instance = 1  
  Case "Put"
    Instance = 2
  Case Else
    Log.Error("SubmitOrder : invalid value for CallPut - "&CallPut)
  End Select

  Row_Click = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",TheCurrentOrder.ProductID,Instance)
  
  If TheCurrentOrder.BidAsk = "Bid" Then
    Col_Click = GridControl.GetCellColumn(OptionViewGrid.Handle,"Bid",Instance)
    Delay(100)
  ElseIf TheCurrentOrder.BidAsk = "Ask" Then
    Col_Click = GridControl.GetCellColumn(OptionViewGrid.Handle,"Ask",Instance)
    Delay(100)
  End If
  
   Call ClickGrid(OptionViewGrid, Row_Click, Col_Click, "Left")

End Sub                                        





'-------------------------------------------------------------------------------------------------------------------------
'Name: Order_MS()
'Arguments: NewOrder[Order object] 
'Description:
'New function to submit orders in MarketView, I will rename this later to something more meaningful
'The difference between this and original that I am passing an Order object with all order details.
'-------------------------------------------------------------------------------------------------------------------------
Public Function Order_MS(NewOrder, NeedHeld)
  
  
     
  Dim QtyCell(1) '(row,col) location of the Bid/Ask Qty cell for the product. Submitting new orders will be using this                      Cell
  Dim Instance 'This determines whether the QtyCell should be taken from the Call or Put side of the OptionViewSheet
  Dim TheoCell(1), LastCell_1 '(Row,Col) location of the Theo cell for this product
  Dim Theo, LastValue_1, TheoValue_1, Last
     
  ' Work out the instance of the column
  Select Case NewOrder.ProductType
  Case "Future" 
      Instance = 1
  	  SearchString = NewOrder.Product & "|" & NewOrder.Month & "|FUTURE"
      Log.Message("SearchString =" & SearchString)
      Log.Message("NewOrder.Product=" & NewOrder.Product)
      QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Product Type",SearchString,Instance)
    Case "Strategy","TMCStrategy"
      Instance = 1
  	  QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",NewOrder.ProductID,Instance)
    Case "Call"
      Instance = 1
      SearchString = NewOrder.Product & "|" & NewOrder.Month & "|" & NewOrder.Strike
      Log.Message("SearchString =" & SearchString)
      QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Strike",SearchString,Instance)
      Log.Message("QtyCell(0) =" & QtyCell(0))  
    Case "Put"
      Instance = 2
      SearchString = NewOrder.Product & "|" & NewOrder.Month & "|" & NewOrder.Strike
      QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Strike",SearchString,Instance)
    Case "Equity" 
      Instance = 1
      SearchString = NewOrder.Product
      QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product",SearchString,Instance)      
      'QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",NewOrder.ProductID,Instance)  
    Case Else
      Log.Error("SubmitOrder : invalid value for CallPut - "&CallPut)
  End Select
  
  'Determine the row of the product and also the TheoCell location
  'QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",NewOrder.ProductID,Instance)
  TheoCell(0) = QtyCell(0)
  If TheoCell(0) = -1 then 
   	Log.Error("[Fail]:  could not find the Product "& NewOrder.ProductID) 
    Exit Function
  End IF  
  TheoCell(1) = GridControl.GetCellColumn(OptionViewGrid.Handle,"Theo",Instance)
  LastCell_1 = GridControl.GetCellColumn(OptionViewGrid.Handle, "Last", Instance)
  
  'Current Theo of the Product
  TheoValue_1 = GridControl.GetCellText(OptionViewGrid.Handle,TheoCell(0),TheoCell(1))
  LastValue_1 = GridControl.GetCellText(OptionViewGrid.Handle,TheoCell(0),LastCell_1)
  
  If TheoValue_1 <> "" Then
    Theo = StrToFloat(TheoValue_1)
  Else 
    log.Error ("The theo for product " & NewOrder.ProductID & " is not available.")
    Exit Function
  End If

  If LastValue_1 <> "" Then
    Last = StrToFloat(LastValue_1)
  ElseIf LastValue_1 = "" Then
    Last = "Blank"        
  End If
  
  
  'Determine the column location of the Bid or Ask Qty cell
    If NewOrder.BidAsk = "Bid" Then   
      QtyCell(1) = GridControl.GetCellColumn(OptionViewGrid.Handle,"Bid Qty",Instance)      
    ElseIf NewOrder.BidAsk = "Ask" Then
      QtyCell(1) = GridControl.GetCellColumn(OptionViewGrid.Handle,"Ask Qty",Instance)
    End If
  
  'Open the order ticket
  'Log.Message("Order Ticket coordinates are: " & QtyCell(0) & ", " & QtyCell(1) & ", " & Instance)
  Call Ticket(OptionViewGrid,QtyCell(0),QtyCell(1)) 
   
  'Calculate the Price based on the Theo
  NewOrder.Price = Order1.CalculatePriceAmend(Last, NewOrder.PriceFormula,Theo,NewOrder.OrderTickSize)
  
  'Set the Price, Qty, Order Restriction in the Ticket
  Call Order1.SetPrice( NewOrder.Price,"Keys")
  Call Order1.SetQuantity(NewOrder.Quantity,"Keys")
  
  Dim dlgOrderTicket1
  Set dlgOrderTicket1 = Order1.GetOrderTicket
  Dim dlgOrderTicket2
  Set dlgOrderTicket2 = Aliases.MarketView.dlgOrderTicket
  
  If NewOrder.Held = "Yes" Then   
    Log.Message("The """&NewOrder.OrderName&""" is a Held Order~!")    
    dlgOrderTicket1.btnH.Click
    'If Aliases.MarketView.WaitAliasChild("dlgOrderTicket").Exists Then
      'Log.Message("Begin to check whether the WaitAliasChild exist")
      'dlgOrderTicket2.Held.Click
    'End If
    NeedHeld = True   
  Else
    Log.Message("The """&NewOrder.OrderName&""" is not a Held Order~!")
  End If
  
  
  Dim dlgOrderTicket
  Set dlgOrderTicket = Aliases.MarketView1.dlgOrderTicket
  Call dlgOrderTicket.btnSubmit.ClickButton
  
  If Aliases.MarketView1.WaitAliasChild("dlgSubmitOrder", 200).Exists Then
    Call Aliases.MarketView1.dlgSubmitOrder.btnOK.ClickButton
  End If
  
  Call SetRestriction(NewOrder.OrderRestriction)  
    
  
  'Check the number of rows in orderview before submission
  Dim OrderViewRowCountBefore, OrderViewRowCountAfter
    
  OrderViewRowCountBefore = GridControl.GetRowCount(OrderViewGrid.Handle)
  'Log.Message("Number of Rows in OrderView before submission is " & OrderViewRowCountBefore)  
  
  'Create a reference to the order ticket so you can click on either submit or cancel 
  
  
 
 
  'dlgOrderTicket.btnCancel.Click
   
  'After an Order submission, there is a small delay before OrderView is updated
  'Adding a loop until OrderView is updated, if not updated within 5 seconds then halt and send an error message
  Dim Timeout
  Timeout = DateAdd("s",5,Now)
     
  'Do
    Delay(700)
    OrderViewRowCountAfter = GridControl.GetRowCount(OrderViewGrid.Handle)
    If TimeOut = Now Then
      Log.Error("Timeout Error: It's been 5 seconds and OrderView has not updated a new order row")
      Log.Message("Order has not been added to OrderView, cannot verify in MarketView")           
    End If            
  'Loop Until OrderViewRowCountAfter <> OrderViewRowCountBefore
  
  'Log.Message("Number of Rows in OrderView after submission is " & OrderViewRowCountAfter)  
  'Log.Message("Order has been successfully added to OrderView")
      
  'Get the ORDER ID from OrderView
  Dim OrderID
  Dim OrderIDColumn
  
  OrderIDColumn = GridControl.GetCellColumn(OrderViewGrid.Handle,"Order ID",1)  
  OrderID = GridControl.GetCellText(OrderViewGrid.Handle,OrderViewRowCountAfter,OrderIDColumn)
  'Log.Message("OrderView OrderID for Order Price=" & NewOrder.Price & ", Qty=" & NewOrder.Quantity & " is " & OrderID)
  'GetOrderViewDetails(OrderID)
  
  Timeout = DateAdd("s",5,Now)
  Dim OrderStatus, OrderConfirmed
  
  OrderConfirmed = False
  OrderStatus = GetColumnValue(OrderID,"Order Status")
  
  If OrderStatus = "Processing" Or OrderStatus = "Submitted" Then
    Do
      OrderStatus = GetColumnValue(OrderID,"Order Status")
        If OrderStatus = "Processing" Or OrderStatus = "Submitted" Then
          OrderConfirmed = False 
            If TimeOut = Now Then
              Log.Error("Timeout Error: It's been 5 seconds and Order status is still " & OrderStatus)
            End If
        Else
          OrderConfirmed = True
        End If  
    Loop Until OrderConfirmed = True 
  End If
'    
'
  Dim AdditionalInfo
  AdditionalInfo = ("Name: " & NewOrder.OrderName & VBNewLine _ 
  & "Order ID: " & OrderID & " " & VBNewLine _
  & "Price: " & NewOrder.Price & VBNewLine _   
  & "Theo: " & GetColumnValue(OrderID,"Theo") & VBNewLine _ 
  & "Order Status: " & GetColumnValue(OrderID,"Order Status") & VBNewLine _
  & "Volume: " & GetColumnValue(OrderID,"Volume") & VBNewLine _
  & "Residual Volume: " & GetColumnValue(OrderID,"Residual Volume") & VBNewLine _
  & "Executed Volume: " & GetColumnValue(OrderID,"Executed Volume") & VBNewLine _  
  & "Product Name: " & GetColumnValue(OrderID,"Product Name") & VBNewLine _
  & "Price (OrderView): " & GetColumnValue(OrderID,"Price") & VBNewLine _  
  & "Time: " & GetColumnValue(OrderID,"Time"))
  
  Call Log.Message("Submitting Order: " & NewOrder.OrderName & " (Click Here for Order Details)", AdditionalInfo)
  
  NewOrder.OrderStatus = GetColumnValue(OrderID,"Order Status") 
  NewOrder.OrderID = OrderID
  Order_MS = OrderID 
                     
End Function


'-------------------------------------------------------------------------------------------------------------------------
'Public Function CalculatePriceAmend(PriceFormula,StartPrice,TickSize)
'-------------------------------------------------------------------------------------------------------------------------
Public Function CalculatePriceAmend(LastPrice, PriceFormula,Theo,TickSize)

  Dim AmendValue
  
  If LastPrice = "Blank" Then
    Log.Message("The Last Column is empty, can not pick up the price based on Last Column on MV, trying to replace the price with the Theo column's value")
    LastPrice = Theo
  End If
    
  Theo = Round(Theo, 2)

  If Left(PriceFormula,1) = "T" Or Left(PriceFormula,1) = "t" Then
    If Mid(PriceFormula,2,1) = "+" Then
      AmendValue = CInt(Mid(PriceFormula,2))
      CalculatePriceAmend = Round((Floor(Theo,TickSize) + (AmendValue * TickSize)),2) 
    ElseIf Mid(PriceFormula,2,1) = "-" Then
      AmendValue = CInt(Mid(PriceFormula,2))
      CalculatePriceAmend = Round((Floor(Theo,TickSize) + (AmendValue * TickSize)),2) 
    Else
      Log.Message("Invalid Operator specified: must either be + or - ")
    End If    
  
  ElseIf Left(PriceFormula,1) = "P" Or Left(PriceFormula,1) = "p" Then
    If Mid(PriceFormula,2,1) = "+" Then
      AmendValue = CInt(Mid(PriceFormula,2))
      CalculatePriceAmend = Round((Theo + (AmendValue * TickSize)),2)  
    ElseIf Mid(PriceFormula,2,1) = "-" Then
      AmendValue = CInt(Mid(PriceFormula,2))
      CalculatePriceAmend = Round((Theo + (AmendValue * TickSize)),2)  
    Else
      Log.Message("Invalid Operator specified: must either be + or - ")
    End If 
  
  ElseIf left(PriceFormula, 1) = "L" Or Left(PriceFormula,1) = "l" Then
    If Mid(PriceFormula, 2,1) = "+" Then
      AmendValue = CInt(Mid(PriceFormula,2))
      CalculatePriceAmend = Round((LastPrice + (AmendValue * TickSize)),2)
    ElseIf Mid(PriceFormula,2,1) = "-" Then
      AmendValue = CInt(Mid(PriceFormula,2))
      CalculatePriceAmend = Round((LastPrice + (AmendValue * TickSize)),2)  
    Else  
      Log.Message("Invalid Operator specified: must either be + or -")
    End If  
    
  
  ElseIf PriceFormula <> "" Then
    'For the "Trade" action only
    CalculatePriceAmend = StrToFloat(PriceFormula) 
  
  Else
    Log.Error("[Error]:Price Formula should start with T or t") 
    On Error Resume Next
       
  End If 
 
End Function



'-----------------------------------------------------------------------------------------------
'Function VerifyOrderView(OrderID, ColumnName, ExpectedResult)
'
'-----------------------------------------------------------------------------------------------
Function VerifyOrderView(MyOrder, CurrentAction, ColumnName, ExpectedResult) 
  
  'There's a slight delay in the update of Residual Volume when pulling an order through MarketView
  'In future need to replace the Delays with a better way of compensating
  Delay(150)
  
  Dim EventLogGrid, EventMessage, MultiOrders(), Result', OrderNumber
  'Error checking
  If MyOrder.OrderID = "" Then
    Log.Error("Cannot verify this order as it does not exist: No OrderID available")
  End If
  
    
  
  If CurrentAction = "MultiOrderPull" Then
    If ColumnName = "Order Status" and ExpectedResult = "Pulled" or ColumnName = "Residual Volume" and ExpectedResult = "0" Then
     MultiOrders = GetOrderNames(Myorder.OrderName)
     Result = True
'     OrderNumber = Ubound(MultiOrders)
     For each Order in MultiOrders
        ActualResult = OrderView.GetColumnValue(MyOrder.OrderID, ColumnName)
        If ActualResult <> ExpectedResult Then 
          Result = False
          Log.Error("[Fail] " & Order & "'s " & ColumnName &" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")  
        End If
      Next 
      If  Result = True Then
        Log.Checkpoint("[Pass] " &ColumnName&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""") 
      End If
      Exit Function
    End If
  End If      
      
    
  
  If ColumnName = "Price" Then
    'Here I'm actually 'recalculating' the Price that is entered in the Order ticket as the user is required to
    'pass the ExpectedResult (Price Formula) as VBScript does not support optional paramaters/arguments 
    
    Dim OrderViewTheo, ExpectedPrice, ActualPrice
    OrderViewTheo = StrToFloat(OrderView.GetColumnValue(MyOrder.OrderID, "Theo"))
    'Debug
    'Log.Message("OverViewTheo: " & OrderViewTheo)
    
    If Left(ExpectedResult,1) = "T" Or Left(ExpectedResult,1) = "t" Then
      ExpectedPrice = Order.CalculatePriceAmend(ExpectedResult,OrderViewTheo,MyOrder.OrderTickSize)
    ElseIf Left(ExpectedResult,1) = "P" Or Left(ExpectedResult,1) = "p" Then
      ExpectedPrice = Order.CalculatePriceAmend(ExpectedResult,MyOrder.PriceBefore,MyOrder.OrderTickSize)      
    Else
      ExpectedPrice = ExpectedResult      
    End If
    
    ActualPrice = StrToFloat(OrderView.GetColumnValue(MyOrder.OrderID,ColumnName))
    Call TestUtilities.CheckFloat(ColumnName, ExpectedPrice, ActualPrice)    
  
  ElseIf MyOrder.OrderRestriction = "IOC" and CurrentAction = "Amend" and ColumnName = "Order Status" Then
    Set EventLogGrid =  Aliases.OrderView.wndAfx.EventLogControlBar.EventLog.EventLogGrid
    EventMessage = GridControl.GetCellText(EventLogGrid.Handle,1,3)
    If EventMessage = "Order cannot be amended as it is no longer active or doesn't support it." Then
      ActualResult = "Cannot amend"
    Else ActualResult = EventMessage
    End If
    Call TestUtilities.CheckValue(ColumnName, ExpectedResult, ActualResult)
    
  Else
    Dim ActualResult
    ActualResult = OrderView.GetColumnValue(MyOrder.OrderID, ColumnName)
    ExpectedResult = CStr(ExpectedResult)
    Call TestUtilities.CheckValue(ColumnName, ExpectedResult, ActualResult)
  
  End If

  
End Function

   

Function GetOrderNames (OrderNames)

  Dim OrderArray, i, Order, OrderNumber

'  Dim OrderNames
'  OrderNames = "Order2, Order3, Order4, Order5"



  OrderArray = split (OrderNames, ",")
  OrderNumber = Ubound(OrderArray)
  
  For i = 0 to OrderNumber
    OrderArray(i) = Trim(OrderArray(i))
  Next  
  
  GetOrderNames = OrderArray

End Function