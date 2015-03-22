'USEUNIT TestUtilities
'USEUNIT OrderView
'USEUNIT MarketView
'USEUNIT ProductDefaults

Private GridControl, OptionViewGrid, OrderViewGrid
Private FsaValue
Private HeldValue

Dim TestConfig
Set TestConfig = ProjectSuite.Variables

Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl
Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
Set OrderViewGrid = Aliases.OrderView.wndAfx.OrderViewGrid

Public Sub Initialize
  Set GridControl = TestConfig.QuantCOREControl
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
  Set OrderViewGrid = Aliases.OrderView.wndAfx.OrderViewGrid
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Class: MarketViewOrder
' Description
'-------------------------------------------------------------------------------------------------------------------------
Class MarketViewOrder
  
  Dim OrderName
  Dim ProductType 'Future, Option, Strategy, Equity
  Dim PriceFormula
  Dim Quantity
  Dim BidAsk
  Dim Month
  Dim OrderStatus
  Dim ProductTestID
  
  Dim Series
  Dim Strike
  Dim OrderRestriction
  Dim Price
  Dim PriceBefore 'Original price before an amendment
  Dim OrderID
  Dim LastPrice
  Dim StrategyNumber
  Dim StrategyName
  Dim StrategyID
  Dim Held
  Dim AnotherSameGroupUser
  Dim UnderlyingTheo
  Dim Theo 'MarketView theo of the product under test at the time of Order submission
      
  Dim m_Product
  Dim m_TickSize
  Dim m_ProductPath
  
      
  'Set the Product (e.g. XJO, NK) that is being tested
  'This is retrieved from the Excel Order Table, Product column
  'Mainly used when testing multiple products 
  Public Property Let Product(x)      
    If x = "" Then      
      Select Case ProductType
        Case "Future"
            m_Product = TestConfig.FutureProduct
        Case "Strategy"
            'This won't work for Option strategies, need to update later
            Product = TestConfig.FutureProduct 
        Case "Call"
            m_Product = TestConfig.OptionProduct
        Case "Put"        
            m_Product = TestConfig.OptionProduct 
        Case "Equity"
            m_Product = TestConfig.EquityProduct 
        Case "TMCStrategy"
        'This won't work for Option strategies, need to update later
            m_Product = TestConfig.FutureProduct
        Case Else
          Log.Error("Invalid ProductType Specified from Worksheet")             
      End Select
    Else
          m_Product = x    
    End If          
  End Property
  
  'Get the Product being tested e.g. XJO, NK
  Public Property Get Product     
    If m_Product = "" Then
        Select Case ProductType
        Case "Future"
            Product = TestConfig.FutureProduct
        Case "Strategy"
            'This won't work for Option strategies, need to update later
            Product = TestConfig.FutureProduct 
        Case "Call"
            Product = TestConfig.OptionProduct
        Case "Put"        
            Product = TestConfig.OptionProduct 
        Case "Equity"
            Product = TestConfig.EquityProduct 
        Case "TMCStrategy"
        'This won't work for Option strategies, need to update later
            Product = TestConfig.FutureProduct
        Case Else
          Log.Error("Invalid ProductType Specified from Worksheet")             
      End Select   
      Else
        Product = m_Product
      End If                  
  End Property 
    
  'Create the Product Path (this will be used for top level selection in Product Defaults)
  'ProductPath is created based on the ProductType
  Public Property Get ProductPath 
         
    If m_Product = "" Then
        Select Case ProductType
        Case "Future"
            m_Product = TestConfig.FutureProduct
        Case "Strategy"
            'This won't work for Option strategies, need to update later
            m_Product = TestConfig.FutureProduct 
        Case "Call"
            m_Product = TestConfig.OptionProduct
        Case "Put"        
            m_Product = TestConfig.OptionProduct 
        Case "Equity"
            m_Product = TestConfig.EquityProduct 
        Case "TMCStrategy"
        'This won't work for Option strategies, need to update later
            m_Product = TestConfig.FutureProduct
        Case Else
          Log.Error("Invalid ProductType Specified from Worksheet")             
      End Select   
  
      End If                  
   
    Select Case ProductType
      Case "Future"
          ProductPath = "SIM.F." &m_Product &".>"
      Case "Strategy"
          If Left(StrategyNumber,1)= "F" Then

            ProductPath = "SIM.F." &TestConfig.FutureProduct &".STRATEGIES.>" 
          Else   
            ProductPath = "SIM.O." &TestConfig.OptionProduct &".STRATEGIES.>"
          End If  

      Case "Call"
          ProductPath = "SIM.O." &m_Product &".>"
      Case "Put"        
          ProductPath = "SIM.O." &m_Product &".>" 
      Case "Equity"
          ProductPath = "SIM.E." &m_Product 
      Case "TMCStrategy"
      'This won't work for Option strategies, need to update later
          ProductPath = "SIM.F." &TestConfig.FutureProduct &".>"
        Case Else
        Log.Error("Invalid ProductType Specified from Worksheet")            
    End Select
  End Property
    
  'ProductID as indicated in MarketView 
  'Create the ProductID based on ProductType, Product, and Long Month e.g. SIM.F.XJO.OCT2011 
  Public Property Get ProductID
    Dim Row, Col
      Select Case ProductType
        Case "Future"
            'Note the month has to be written in LONG format in the spreadsheet
          ProductID = "SIM.F." &m_Product &"." &Month
        Case "Strategy", "TMCStrategy"
            'Note the month has to be written in LONG format in the spreadsheet
          ProductID = StrategyID
        Case "Call"
          'Note the month has to be written in LONG format in the spreadsheet
           Row = GridControl.GetCellRow(OptionViewGrid.Handle,"Series|Strike", Month & "|" & Strike,1)
	        Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"ProductID",1)
    	    ProductID = GridControl.GetCellText(OptionViewGrid.Handle,Row,Col)
        Case "Put"
          'Note the month has to be written in LONG format in the spreadsheet    
	      Row = GridControl.GetCellRow(OptionViewGrid.Handle,"Series|Strike", Month & "|" & Strike,1)
    	    Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"ProductID",2)
        	ProductID = GridControl.GetCellText(OptionViewGrid.Handle,Row,Col)
        Case "Equity"
          ProductID = "SIM.E." &Product               
      Case Else
        Log.Error("Invalid ProductType Specified from Worksheet")            
    End Select
  End Property
  
  'Set the OrderTickSize based on ProductType, need to add cases for Equities, Strategies
  Public Property Get OrderTickSize
    If m_TickSize = "" Then
      Select Case ProductType
        Case "Future"
          OrderTickSize = TestConfig.FutureTickSize
        Case "Equity"
          OrderTickSize = TestConfig.EquityTickSize
        Case "Call"
          OrderTickSize = TestConfig.OptionTickSize
        Case "Put"
          OrderTickSize = TestConfig.OptionTickSize 
        Case "Strategy", "TMCStrategy"
          OrderTickSize = TestConfig.OptionTickSize      
        Case Else
          Log.Error("Invalid ProductType Specified from Worksheet")            
      End Select
    Else
      OrderTickSize = m_TickSize 
    End If
  End Property

  Public Property Let OrderTickSize(x)      
    If x = "" Then      
      Select Case ProductType
        Case "Future"
            m_TickSize = TestConfig.FutureTickSize
        Case "Strategy"
            'This won't work for Option strategies, need to update later
             m_TickSize = TestConfig.FutureTickSize
        Case "Call"
             m_TickSize = TestConfig.OptionTickSize
        Case "Put"        
             m_TickSize = TestConfig.OptionTickSize
        Case "Equity"
            m_TickSize = TestConfig.EquityTickSize 
        Case "TMCStrategy"
        'This won't work for Option strategies, need to update later
             m_TickSize = TestConfig.FutureProduct
        Case Else
          Log.Error("Invalid ProductType Specified from Worksheet")             
      End Select
    Else
           m_TickSize = x    
    End If          
  End Property 
  
End Class


' ------------------------------------------------------------------------------------
' Returns a reference for the type of order ticket that has opened
' ------------------------------------------------------------------------------------
'Private Function GetOrderTicket
Public Function GetOrderTicket 
	On Error Resume Next
  If Aliases.MarketView.WaitAliasChild("dlgOrderTicket").Exists Then
    Set GetOrderTicket = Aliases.MarketView.dlgOrderTicket
  ElseIf Aliases.MarketView.WaitAliasChild("dlgAmendOrder").Exists Then
    Set GetOrderTicket = Aliases.MarketView.dlgAmendOrder
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
  If Aliases.MarketView.dlgOrderTicket.WaitAliasChild("FsaTab", 50).Visible And FsaValue = False Then
    dlgOrderTicket.btnF.Click
  ElseIf Not Aliases.MarketView.dlgOrderTicket.WaitAliasChild("FsaTab", 50).Visible And FsaValue = True Then
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
    If Aliases.MarketView.dlgOrderTicket.WaitAliasChild("AOM",50).Visible = True Then
      If dlgOrderTicket.AOM.Offset.Enabled = True Then
        dlgOrderTicket.btnAC.Click
      End If
    End If
  ElseIf ACEnabled = True Then
    If Aliases.MarketView.dlgOrderTicket.WaitAliasChild("AOM",50).Visible = False Then
      dlgOrderTicket.btnAC.Click
    ElseIf Aliases.MarketView.dlgOrderTicket.WaitAliasChild("AOM",50).Visible = True And dlgOrderTicket.AOM.Offset.Enabled = False Then
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
    If Aliases.MarketView.dlgOrderTicket.WaitAliasChild("AOM", 50).Visible = True Then
      If dlgOrderTicket.AOM.Trigger.Enabled = True  Then
        dlgOrderTicket.btnAS.Click
      End If
    End If
  ElseIf ASEnabled = True And Aliases.MarketView.dlgOrderTicket.WaitAliasChild("AOM", 50).Visible = False Then
    dlgOrderTicket.btnAS.Click
  ElseIf ASEnabled = True And Aliases.MarketView.dlgOrderTicket.WaitAliasChild("AOM", 50).Visible = True And dlgOrderTicket.AOM.Trigger.Enabled = False Then
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
  
    Aliases.MarketView.dlgOrderTicket.AOM.Offset.Click
    
    For i = 0 To Abs(ACOffset) - 1
      If ACOffset > 0 Then
        Aliases.MarketView.dlgOrderTicket.AOM.Offset.MouseWheel(1)
      Else
        Aliases.MarketView.dlgOrderTicket.AOM.Offset.MouseWheel(-1)
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
  
  If Aliases.MarketView.dlgOrderTicket.WaitAliasChild("AOM",50).Visible = False Or dlgOrderTicket.AOM.Trigger.Enabled = False Then
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
        Aliases.MarketView.dlgOrderTicket.AOM.TriggerArrow.Up
      Else
        Aliases.MarketView.dlgOrderTicket.AOM.TriggerArrow.Down
      End If
    Next
  ' For MouseWheel events, the ACOffset should be specified in terms of the number of mouse wheel clicks
  ElseIf Method = "MouseWheel" Then
    If TestUtilities.FMod(ASTrigger, 1) <> 0 Then
      Log.Error("SetASTrigger - when using MouseWheel, the AS Trigger should be specified in whole number of notches to move the mouse wheel")
      Exit Sub
    End If
  
    Aliases.MarketView.dlgOrderTicket.AOM.Trigger.Click
    
    For i = 0 To Abs(ASTrigger) - 1
      If ASTrigger > 0 Then
        Aliases.MarketView.dlgOrderTicket.AOM.Trigger.MouseWheel(1)
      Else
        Aliases.MarketView.dlgOrderTicket.AOM.Trigger.MouseWheel(-1)
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
  If Aliases.MarketView.WaitAliasChild("dlgOrderWarning").Exists Then
    Dim dlgOrderWarning
    Set dlgOrderWarning = Aliases.MarketView.dlgOrderWarning
    Call Aliases.MarketView.dlgOrderWarning.btnYes.ClickButton    
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
  Call Order.Ticket(GridObject, Row, Col)
  Call Order.SetQuantity(Quantity,"Keys")
  Call Order.SetPrice(Price,"Keys")
  Call Order.SetAC(False)
  Call Order.SetAS(False)
  Call Order.SetHeld(Held)

  SubmitOrder = Order.Submit(True)
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
   
  Call Order.Ticket(GridObject, Row, Col)
  Call Order.SetQuantity(Quantity,"Keys")
  Call Order.SetPrice(Price,"Keys")
  Call Order.SetAC(True)
  Call Order.SetACOffset(Offset,"Keys")
  Call Order.SetHeld(Held)
  
  SubmitACOrder = Order.Submit(True)
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

  Call Order.Ticket(GridObject, Row, Col)
  Call Order.SetQuantity(Quantity,"Keys")
  Call Order.SetPrice(Price,"Keys")
  Call Order.SetAS(True)
  Call Order.SetRestriction(Restriction)
  Call Order.SetASTrigger(Trigger,"Keys")
  
  SubmitASOrder = Order.Submit(True)
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
  
  Call Order.Ticket(GridObject, Row, Col)
  Call Order.SetQuantity(Quantity,"Keys")
  Call Order.SetPrice(Price,"Keys")
  Call Order.SetAC(True)
  Call Order.SetAS(True)
  Call Order.SetACOffset(Offset,"Keys")
  Call Order.SetASTrigger(Trigger,"Keys")
  
  SubmitASCOrder = Order.Submit(True)
End Function

Sub SetUnderlyingPrice(ProductID, RefPrice, Quantity)
  Dim Bid, Ask, Row, Col, TickSize
  TickSize = TestConfig.UnderlyingTickSize
  
  ' Calculate the bid and the ask
  Bid = Floor(RefPrice, TickSize)

  If Floor(RefPrice,TickSize) = Ceiling(RefPrice,TickSize) Then
    Ask = Ceiling(RefPrice + TickSize, TickSize)
  Else
    Ask = Ceiling(RefPrice, TickSize)
  End If
  
  Row = GridControl.GetCellRow(OptionViewGrid.Handle, "ProductID",ProductID,1)
  Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Bid",1)
  
  ' Open order ticket and set price
  Call Ticket(OptionViewGrid, Row, Col)
  Call SetPrice(Bid, "Keys")
  
  ' Submit the whole quantity
  Dim RemainingQuantity, OrderQuantity
  RemainingQuantity = Quantity
  Do
    If RemainingQuantity > TestConfig.RiskLimit Then
      OrderQuantity = TestConfig.RiskLimit
    Else
      OrderQuantity = RemainingQuantity
    End If
    
    If Order.GetQuantity <> OrderQuantity Then
      Call SetQuantity(OrderQuantity,"Keys")
    End If
    
    RemainingQuantity = RemainingQuantity - OrderQuantity
    
    Call Submit(False)
  Loop Until RemainingQuantity = 0 
  
  Call Order.Cancel
   
  Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Ask",1)

  ' Open order ticket and set price
  Call Ticket(OptionViewGrid, Row, Col) 
  Call SetPrice(Ask, "Keys")
 
  RemainingQuantity = Quantity 
  Do
    If RemainingQuantity > TestConfig.RiskLimit Then
      OrderQuantity = TestConfig.RiskLimit
    Else
      OrderQuantity = RemainingQuantity
    End If
    
    If Order.GetQuantity <> OrderQuantity Then
      Call SetQuantity(OrderQuantity,"Keys")
    End If
    
    RemainingQuantity = RemainingQuantity - OrderQuantity
    
    Call Submit(False)
  Loop Until RemainingQuantity = 0 
  
  Call Order.Cancel
  
  Dim Theo
  Row = GridControl.GetCellRow(OptionViewGrid.Handle, "ProductID",ProductID,1)
  Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Theo",1) 
  Theo = GridControl.GetCellText(OptionViewGrid.Handle,Row,Col)
  
  Log.Message("Set underlying - Bid = "&Bid&", Ask = "&Ask&", Theo = "&Theo)
End Sub
     
Sub NudgeUnderlying(ProductID, Direction)
  Dim BidPrice
  Dim AskPrice
  Dim BidColour
  Dim AskColour  
  Dim Bid1Colour
  Dim Ask1Colour
    
  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
  
  Dim Row, Col
  Row = GridControl.GetCellRow(OptionViewGrid.Handle, "ProductID", ProductID, 1)
  
  BidPrice = GetFloatFromRow(OptionViewGrid,Row,"Bid",1)    
  AskPrice = GetFloatFromRow(OptionViewGrid,Row,"Ask",1)
  BidColour = GetCellBackgroundColourFromRow(OptionViewGrid,Row,"Bid",1)
  AskColour = GetCellBackgroundColourFromRow(OptionViewGrid,Row,"Ask",1) 
  
  ' This is to make sure the number is calculated correctly
  ' When you get the number from the grid it can be 5.00000000000000000001 for example
  Dim DiffAskBid
  DiffAskBid = StrToFloat(TestUtilities.FormatDecimals(AskPrice - BidPrice, 7))
  
  Call MarketView.MakeCellVisible(OptionViewGrid,Row,QuantCOREControl.GetCellColumn(OptionViewGrid.Handle,"Theo",1))
  
  If Direction = "Down" Then 
    ' If the spread is only one tick, then move the bid down to lower the underlying
    ' If the cell is not green (00FF00) then move the ask down as the bid is not ours to move
    If DiffAskBid = TestConfig.UnderlyingTickSize And BidColour = "0000FF00" Then
      Col = GridControl.GetCellColumn(OptionViewGrid.Handle, "Bid", 1)
      Call Order.TickOrderWorse(OptionViewGrid, Row, Col)
    ElseIf AskColour = "0000FF00" Then
      Col = GridControl.GetCellColumn(OptionViewGrid.Handle, "Ask", 1)
      Call Order.TickOrderBetter(OptionViewGrid, Row, Col)
    Else
      Log.Error("Error moving underlying, Bid = "&BidPrice&" BidColour = "&BidColour&" Ask = "&AskPrice&" AskColour = "&AskColour)
    End If
  ElseIf Direction = "Up" Then
    ' If the spread is only one tick, then move the bid up to lower the underlying
    ' If the cell is not green (00FF00) then move the ask up as the bid is not ours to move
    If DiffAskBid = TestConfig.UnderlyingTickSize And AskColour = "0000FF00" Then
      Col = GridControl.GetCellColumn(OptionViewGrid.Handle, "Ask", 1)
      Call Order.TickOrderWorse(OptionViewGrid, Row, Col)
    ElseIf BidColour = "0000FF00" Then
      Col = GridControl.GetCellColumn(OptionViewGrid.Handle, "Bid", 1)
      Call Order.TickOrderBetter(OptionViewGrid, Row, Col)
    Else
      Log.Error("Error moving underlying, Bid = "&BidPrice&" BidColour = "&BidColour&" Ask = "&AskPrice&" AskColour = "&AskColour)
    End If
  End If
  
  Delay(500)
  
  Dim Count
  Count = 0
  Do
    Delay(250)
    BidColour = GetCellBackgroundColourFromRow(OptionViewGrid,Row,"Bid",1)
    AskColour = GetCellBackgroundColourFromRow(OptionViewGrid,Row,"Ask",1)
    Bid1Colour = GetCellBackgroundColourFromRow(OptionViewGrid,Row+1,"Bid",1)
    Ask1Colour = GetCellBackgroundColourFromRow(OptionViewGrid,Row+1,"Ask",1)
    Count = Count + 1
  Loop Until ((BidColour = "0000FF00" Or Bid1Colour = "0000FF00") And (AskColour = "0000FF00" Or Ask1Colour = "0000FF00")) Or Count = 60
  
  If ((BidColour = "0000FF00" Or Bid1Colour = "0000FF00") And (AskColour = "0000FF00" Or Ask1Colour = "0000FF00")) Then
    Log.Message("Bid and ask have correct colour") 
  Else
    Log.Message(BidColour&"  "&Bid1Colour&"  "&AskColour&"  "&Ask1Colour&"Bid and ask have correct colour")
  End If
End Sub

Sub SetRestriction(Restriction)
  Dim dlgOrderTicket
  Set dlgOrderTicket = Aliases.MarketView.dlgOrderTicket
  
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
  Case "Strategy","TMCStrategy"
    Instance = 1    
  Case "Call"
    Instance = 1  
  Case "Put"
    Instance = 2
  '19/06/12
   Case "Equity"
    Instance = 1
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
  Dim TheoCell(1) '(Row,Col) location of the Theo cell for this product
  Dim LastCell '(Row,Col) localtion of the Last price cell for the test product
  Dim Last
  Dim Theo
  Dim SearchSTring
  
  ' Work out the instance of the column
  Select Case NewOrder.ProductType
    Case "Future" 
      Instance = 1
  	  SearchString = NewOrder.Product & "|" & NewOrder.Month & "|FUTURE"
      QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Product Type",SearchString,Instance)
    Case "Strategy","TMCStrategy"
      Instance = 1
  	  QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",NewOrder.ProductID,Instance)
    Case "Call"
      Instance = 1
      SearchString = NewOrder.Product & "|" & NewOrder.Month & "|" & NewOrder.Strike
      QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Strike",SearchString,Instance)  
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
  LastCell = GridControl.GetCellColumn(OptionViewGrid.Handle, "Last", Instance) 
  'Current Theo of the Product - handle the error if the theo is not available
  
  Dim TheoValue, LastValue
  TheoValue = GridControl.GetCellText(OptionViewGrid.Handle,TheoCell(0),TheoCell(1))
  LastValue = GridControl.GetCellText(OptionViewGrid.Handle,TheoCell(0),LastCell)
  
  If TheoValue <> "" Then
    Theo = StrToFloat(TheoValue)
    NewOrder.Theo = Theo
  Else 
    log.Error ("The theo for product " & NewOrder.ProductID & " is not available.")
    Exit Function
  End If
  
  If LastValue <> "" Then
    Last = StrToFloat(LastValue)
    NewOrder.LastPrice = Last
  Else
    NewOrder.LastPrice = Theo
    'Log.Message("No Last Price available - using Theo instead")            
  End If 

 ''Submit trade fix  
  If NewOrder.ProductID = TestConfig.UnderlyingProductID Then
    If NewOrder.UnderlyingTheo = "" Then
      NewOrder.UnderlyingTheo = Theo
    End If
  End If
  
  '19/06/12
  If NewOrder.ProductType = "Equity" Then
    If NewOrder.UnderlyingTheo = "" Then
      NewOrder.UnderlyingTheo = Theo
    End If
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
  'NewOrder.Price = Order.CalculatePriceAmend(NewOrder.PriceFormula,Theo,NewOrder.OrderTickSize)
 ' NewOrder.Price = Order.CalculatePriceAmend(NewOrder.PriceFormula,Theo,NewOrder.OrderTickSize)
  
  'Calculate the Price based on the Theo
  'NewOrder.Price = Order.CalculatePriceAmend(Last, NewOrder.PriceFormula,Theo,NewOrder.OrderTickSize)
  
  'Calculate the Price based on the Theo if it has not been defined
  'If NewOrder.Price = "" And NewOrder.UnderlyingTheo = "" Then
    'Log.Message("The NewOrder.Price = blank")
    'Log.Message("The NewOrder.PriceFormula =" & NewOrder.PriceFormula)  
    'NewOrder.Price = Order.CalculatePriceAmend(Last,NewOrder.PriceFormula,Theo,NewOrder.OrderTickSize)
    NewOrder.Price = Order.CalculatePriceAmend(NewOrder, NewOrder.PriceFormula)
    'Log.Message("NewOrder.Price =" & NewOrder.Price)    
  '19/06/12
  'Else
    'NewOrder.Price = Order.CalculatePriceAmend(Last,NewOrder.PriceFormula,NewOrder.UnderlyingTheo,NewOrder.OrderTickSize)
    'Log.Message("Hi")
  'End If
  
  
    
  'Set the Price, Qty, Order Restriction in the Ticket
  Call Order.SetPrice( NewOrder.Price,"Keys")
  Call Order.SetQuantity(NewOrder.Quantity,"Keys")
  Call SetRestriction(NewOrder.OrderRestriction)    
  
  'Check the number of rows in orderview before submission
  Dim OrderViewRowCountBefore, OrderViewRowCountAfter
    
  OrderViewRowCountBefore = GridControl.GetRowCount(OrderViewGrid.Handle)
  'Log.Message("Number of Rows in OrderView before submission is " & OrderViewRowCountBefore)  
  
  'Create a reference to the order ticket so you can click on either submit or cancel
    
  
  Dim dlgOrderTicket
  Set dlgOrderTicket = Order.GetOrderTicket
  
  If NewOrder.Held = "Yes" Then   
    Log.Message("The """&NewOrder.OrderName&""" Held")
    dlgOrderTicket.Held.Click
    NeedHeld = True   
  Else
    'Log.Message("The """&NewOrder.OrderName&""" is not a Held Order~!")
  End If
 
  
  dlgOrderTicket.btnSubmit.Click
  'dlgOrderTicket.btnCancel.Click
  If Aliases.MarketView.WaitAliasChild("dlgInvalidOrder", 200).Exists Then
    Call Aliases.MarketView.dlgInvalidOrder.btnOK.ClickButton
  End If
  
  If Aliases.MarketView.WaitAliasChild("dlgAOMMachineCheck", 200).Exists Then
    Log.Message("The AC Ticket Default is enabled, skip the alert message to submit the AC order directly")
    Call Aliases.MarketView.dlgAOMMachineCheck.btnYes.ClickButton
    If Aliases.MarketView.WaitAliasChild("dlgAOMMachineCheck", 200).Exists Then
      Call Aliases.MarketView.dlgAOMMachineCheck.btnYes.ClickButton
    End If
  End If
   
  'After an Order submission, there is a small delay before OrderView is updated
  'Adding a loop until OrderView is updated, if not updated within 5 seconds then halt and send an error message
  Dim Timeout
  Timeout = DateAdd("s",5,Now)  
  
  Do
    'Delay(700)
    OrderViewRowCountAfter = GridControl.GetRowCount(OrderViewGrid.Handle)
    If TimeOut = Now Then
      Log.Error("Timeout Error: It's been 5 seconds and OrderView has not updated a new order row")
      Log.Message("Order has not been added to OrderView, cannot verify in MarketView")           
    End If            
  Loop Until OrderViewRowCountAfter <> OrderViewRowCountBefore
  
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
  
  NewOrder.OrderStatus = GetColumnValue(OrderID,"Order Status") 
  NewOrder.OrderID = OrderID
  
  LogOrderDetails(NewOrder)
  
  Order_MS = OrderID 
  
  
                     
End Function


'-------------------------------------------------------------------------------------------------------------------------
'Public Function CalculatePriceAmend(PriceFormula,StartPrice,TickSize)
'-------------------------------------------------------------------------------------------------------------------------
Public Function CalculatePriceAmend(MyOrder,PriceFormula)
  


  Dim AmendValue, Theo, TickSize, PriceReference ', LASTget
	Theo = Round(Theo, 2)
  TickSize = MyOrder.OrderTickSize 
    
  PriceReference = Left(PriceFormula,1)
  
  Select Case PriceReference 
    Case "T", "t"
      If MyOrder.UnderlyingTheo = "" Then
        Theo = MyOrder.Theo
        'Log.Message("MyOrder.Theo =" & Theo)
      Else
        Theo = MyOrder.UnderlyingTheo
        'Log.Message("MyOrder.UnderlyingTheo =" & Theo)
      End If
    Case "P", "p"
      Theo = MyOrder.PriceBefore
    Case "L", "l"      
      'Call GrabTheLast(LASTget, MyOrder)
      'Log.Message("LASTget =" & LASTget)
      Theo = MyOrder.LastPrice
    Case ""
      Log.Error("No Price formula")
    Case Else
      CalculatePriceAmend = StrToFloat(PriceFormula)  
  End Select

  
  '19/06/12
  If Mid(PriceFormula,2,1) = "+" Then  
      AmendValue = CInt(Mid(PriceFormula,2))
      CalculatePriceAmend = Round((Floor(Theo,TickSize) + (AmendValue * TickSize)),2)
      'Log.Message("CalculatePriceAmend =" & CalculatePriceAmend)
  ElseIf Mid(PriceFormula,2,1) = "-" Then
      AmendValue = CInt(Mid(PriceFormula,2))
      CalculatePriceAmend = Round((Floor(Theo,TickSize) + (AmendValue * TickSize)),2)
      'Log.Message("CalculatePriceAmend =" & CalculatePriceAmend) 
  Else

      'Log.Message("Invalid Operator specified: must either be + or - ")
  End If
 
End Function



'-------------------------------------------------------------------------------------------------------------------------
'Function: CreateNewOrder
'-------------------------------------------------------------------------------------------------------------------------
Function CreateNewOrder(ProductType, Month, Strike, BidAsk, OrderRestriction, PriceFormula, Quantity, StrategyID)

  'Set GetNewOrderObject = New MarketViewOrder
  Set NewOrderObject = New MarketViewOrder

  NewOrderObject.ProductType = ProductType
  NewOrderObject.Month = Month
  NewOrderObject.BidAsk = BidAsk
  NewOrderObject.Strike = Strike
  NewOrderObject.OrderRestriction = OrderRestriction
  NewOrderObject.PriceFormula = PriceFormula  
  NewOrderObject.Quantity = Quantity 
  NewOrderObject.StrategyID = StrategyID
  NewOrderObject.LastPrice = PriceFormula
  
  Set CreateNewOrder = NewOrderObject
  
End Function  

'-------------------------------------------------------------------------------------------------------------------------
'Function: CreateTradeOrder(MyOrder,Qty)
'
'Description
'This is to create a Trade order based on an order that you already have specified. E.g. If your order is a Bid, then this will
'create the an Ask order with the Qty you specify
'-------------------------------------------------------------------------------------------------------------------------
Function CreateTradeOrder(MyOrder,Qty)

  'Set GetNewOrderObject = New MarketViewOrder
  Set NewOrderObject = New MarketViewOrder

  NewOrderObject.ProductType = MyOrder.ProductType
  NewOrderObject.Month = MyOrder.Month
  NewOrderObject.Strike = MyOrder.Strike
  NewOrderObject.OrderRestriction = MyOrder.OrderRestriction
  NewOrderObject.PriceFormula = MyOrder.PriceFormula  
  NewOrderObject.Quantity = Qty
  NewOrderObject.StrategyID = MyOrder.StrategyID
  NewOrderObject.Product = MyOrder.Product
  NewOrderObject.OrderTickSize = MyOrder.OrderTickSize
  NewOrderObject.LastPrice = PriceFormula 
  
  'Pick the opposite of the order you want to trade with
  If MyOrder.BidAsk = "Bid" Then
    NewOrderObject.BidAsk = "Ask"         
  ElseIf MyOrder.BidAsk = "Ask" Then
    NewOrderObject.BidAsk = "Bid"
  End If
  
  Set CreateTradeOrder = NewOrderObject
  
End Function  

'-------------------------------------------------------------------------------------------------------------------------
'Function: CreateNewOrderTemplate
'-------------------------------------------------------------------------------------------------------------------------
Function CreateNewOrderTemplate(OrderName)

  
  Set NewOrderObject = New MarketViewOrder

  NewOrderObject.OrderName = OrderName
    
  Set CreateNewOrderTemplate = NewOrderObject
  
End Function  





'-----------------------------------------------------------------------------------------------
'Sub TradeOrder(MyOrder,Qty)                                                                      
'
'-----------------------------------------------------------------------------------------------
Sub TradeOrder(MyOrder,PriceFormula,Qty) 

  Dim NewTradeOrder, TradePriceFormula, TradeLast
  
  If PriceFormula = "" Then
    TradePriceFormula = MyOrder.Price
  Else
    TradePriceFormula = PriceFormula
  End If
  
      
  If MyOrder.BidAsk = "Bid" Then
    Set NewTradeOrder = CreateNewOrder(MyOrder.ProductType,MyOrder.Month,MyOrder.Strike,"Ask","IOC",TradePriceFormula,Qty,MyOrder.StrategyID)
  ElseIf MyOrder.BidAsk = "Ask" Then 
    Set NewTradeOrder = CreateNewOrder(MyOrder.ProductType,MyOrder.Month,MyOrder.Strike,"Bid","IOC",TradePriceFormula,Qty,MyOrder.StrategyID)   
  End If

  NewTradeOrder.Product = MyOrder.Product
  
  TradeLast = "Try"
      
  Call GrabTheLast(TradeLast, NewTradeOrder)
  
  Call Order_MS(NewTradeOrder, False)
  
  
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Sub PullAtPrice(ProductID,Price)
'Description: Pulling an order from MarketView, has to be pulled at Price
'-------------------------------------------------------------------------------------------------------------------------
Sub PullAtPrice(OrderObject)

  'Adding a delay for orders to appear in MarketView!
  Delay(TestConfig.fl_delay_MV_updates)
 
  Dim ProductRow, PriceColumn, Instance, OrderPrice

  ' Work out the instance of the column
  Select Case OrderObject.ProductType
  Case "Call", "Future", "Strategy", "Equity"
    Instance = 1
  Case "Put"
    Instance = 2
  Case Else
    Log.Error("SubmitOrder : invalid value for CallPut - "&CallPut)
  End Select
    
  Call MarketView.OpenDepth(OrderObject.ProductID, Instance)
  
  'read the price from the price column and compare against OrderObject.Price
  'if price is correct, then pull order
  
  'Check the first row price
  ProductRow = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",OrderObject.ProductID,Instance)
  PriceColumn = GridControl.GetCellColumn(OptionViewGrid.Handle,OrderObject.BidAsk,Instance)
  OrderPrice = TestUtilities.GetFloatFromRow(OptionViewGrid,ProductRow,OrderObject.BidAsk,Instance)

  
  If OrderPrice = OrderObject.Price Then
    Call Pull(OptionViewGrid,ProductRow,PriceColumn)
    'Log.Message("Pulling Order [Order ID] " & OrderObject.OrderID)
         
  Else
    'Check the rest of the depth
    Dim ProductDepth, DepthStackColumn, DepthStackPosition, PriceFound
    
    DepthStackColumn = GridControl.GetCellColumn(OptionViewGrid.Handle,"Depth Stack Position",Instance) 
    PriceFound = False
    ProductDepth = ProductRow + 1 'The first row below the Product
        
    Do
      DepthStackPosition = GridControl.GetCellText(OptionViewGrid.Handle,ProductDepth,DepthStackColumn)
      If DepthStackPosition = "" Or DepthStackPosition = "0" Then
          If PriceFound = False Then
            Log.Message("[Error] Price " & OrderObject.Price & "not found, order does not appear to be in MarketView")
          End If
        Exit Do
      Else
        OrderPrice = TestUtilities.GetFLoatFromRow(OptionViewGrid,ProductDepth,OrderObject.BidAsk,Instance)
          If OrderPrice = OrderObject.Price Then
            Call Pull(OptionViewGrid,ProductDepth,PriceColumn)
            Log.Message("Pulling Order [Order ID] " & OrderObject.OrderID)
            PriceFound = True
            Exit Do
          Else
            ProductDepth = ProductDepth + 1
            PriceFound = False
          End If
      End If
    Loop
  End If 

  'Print Order details to Log  
  LogOrderDetails(OrderObject)
    
End Sub


'-------------------------------------------------------------------------------------------------------------------------
'Function AmendOrderView(OrderID,Field,Direction,Value)
'Thinking ahead here, but when it comes to writing tests, I might make it so testers can specify where to verify the order
'e.g. Verify in OptionView, or verify in OrderView and have seperate functions that verify
'-------------------------------------------------------------------------------------------------------------------------
Function AmendOrderView(OrderID,Field,Value)                          

    Dim OrderFilled             
    OrderFilled = False

    If Field = "Price" Then
      Call OrderView.Amend(OrderID, OrderFilled)   
      'Check for Amend Error Dialog box
            
      If Aliases.OrderView.WaitAliasChild("dlgAmendError", 200).Exists Then        
        Log.Message("Amend Failed for Order: " & OrderID)                       
        Call Aliases.OrderView.dlgAmendError.btnOK.Click

      
      
      ElseIf OrderFilled = True Then
        Log.Warning("Amend Failed for Filled Order: " &OrderID)   
        AmendOrderView = False
              
      Else
        Call OrderView.SetPrice(Value,"Keys")
        AmendOrderView = True        
      End If
      'Call OrderView.PressAmend      
    ElseIf Field = "Qty" Then
      Call OrderView.Amend(OrderID, OrderFilled)
      'Check for Amend Error Dialog box
      If Aliases.OrderView.WaitAliasChild("dlgAmendError", 200).Exists Then        
        Log.Message("Amend Failed for Order: " & OrderID)
        Call Aliases.OrderView.dlgAmendError.btnOK.Click
      
      ElseIf OrderFilled = True Then
        Log.Warning("Amend Failed for Filled Order: " &OrderID)        
        AmendOrderView = False
      
      Else
        Call OrderView.SetQuantity(Value,"Keys")
        AmendOrderView = True
      End If
    Else
      Log.Error("Invalid Field specified: Must be either Price, Quantity or Qty")
    End If
  
End Function 


'-----------------------------------------------------------------------------------------------
'Function VerifyOrderView(OrderID, ColumnName, ExpectedResult)
'
'-----------------------------------------------------------------------------------------------
Function VerifyOrderView(MyOrder, CurrentAction, ColumnName, ExpectedResult, OrderDict, MultiOrders)
  
 'There's a slight delay in the update of Residual Volume when pulling an order through MarketView
  'In future need to replace the Delays with a better way of compensating
  'Delay(150)
  'KIS
  Delay(TestConfig.fl_delay_OrderView_updates)
  
    Dim EventLogGrid, EventMessage, Result, OrderInfo, Orders
  'Error checking
  If MyOrder.OrderID = "" Then
    Log.Error("Cannot verify this order as it does not exist: No OrderID available")
    Exit Function
  End If
    
  If CurrentAction = "MultiOrderPull" Then
    If ColumnName = "Order Status" or ColumnName = "Residual Volume" Then
      Result = True
      For each Orders in MultiOrders
        Set OrderInfo = OrderDict.Item(Orders)
        ActualResult = OrderView.GetColumnValue(OrderInfo.OrderID, ColumnName)
        If CStr(ActualResult) <> CStr(ExpectedResult) Then 
          Result = False
          Log.Error("[Fail] " & Orders & " " &ColumnName&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")  
        End If
      Next 
      If  Result = True Then
        Log.Checkpoint("[Pass] Multiple Order's " &ColumnName&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")
      End If
    End If   
    Exit Function

  
   ElseIf CurrentAction = "PullAll" Then
    If ColumnName = "Order Status" Then
      Result = True
      For each Orders in MultiOrders
        Set OrderInfo = OrderDict.Item(Orders)
        ActualResult = OrderView.GetColumnValue(OrderInfo.OrderID, ColumnName)
        'If CStr(ActualResult) <> "Confirmed" then ActualResult = "Pulled" End If 
        If CStr(ActualResult) <> CStr(ExpectedResult) Then 
          Result = False
          Log.Error("[Fail] " & Orders & " " &ColumnName&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")  
        End If
      Next 
      If  Result = True Then
        Log.Checkpoint("[Pass] All the Orders are pulled.")
      End If
    End If   
    Exit Function
 
  
  ElseIf CurrentAction = "MultiOrderUnhold" Then
    If ColumnName = "Order Status" Then 'or ColumnName = "Residual Volume" Then
      Result = True
      For each Orders in MultiOrders
        Set OrderInfo = OrderDict.Item(Orders)
        ActualResult = OrderView.GetColumnValue(OrderInfo.OrderID, ColumnName)
        If CStr(ActualResult) <> CStr(ExpectedResult) Then 
          Result = False
          Log.Error("[Fail] " & Orders & " " &ColumnName&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")  
        End If
      Next 
      If  Result = True Then
        Log.Checkpoint("[Pass] Multiple Order's " &ColumnName&" - Expected Result: """&ExpectedResult&""" Actual Result: """&ActualResult&"""")
      End If
    End If   
    Exit Function
  End If 
    
  If ColumnName = "Price" Then
    'Here I'm actually 'recalculating' the Price that is entered in the Order ticket as the user is required to
    'pass the ExpectedResult (Price Formula) as VBScript does not support optional paramaters/arguments 
    
    'Dim OrderViewTheo, ExpectedPrice, ActualPrice
    
    If MyOrder.Theo = "" Then
    
        Dim OpQtyCell(1) '(row,col) location of the Bid/Ask Qty cell for the product. Submitting new orders will be using this                      Cell
        Dim OpInstance 'This determines whether the QtyCell should be taken from the Call or Put side of the OptionViewSheet
        Dim OpTheoCell(1) '(Row,Col) location of the Theo cell for this product
        Dim OpLastCell '(Row,Col) localtion of the Last price cell for the test product
        Dim OpLast
        Dim OpTheo
        Dim OpSearchSTring
  
  ' Work out the instance of the column
        Select Case MyOrder.ProductType
          Case "Future" 
            OpInstance = 1
        	  OpSearchString = MyOrder.Product & "|" & MyOrder.Month & "|FUTURE"
            OpQtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Product Type",OpSearchString,OpInstance)
          Case "Strategy","TMCStrategy"
            OpInstance = 1
        	  OpQtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",MyOrder.ProductID,OpInstance)
          Case "Call"
            OpInstance = 1
            OpSearchString = MyOrder.Product & "|" & MyOrder.Month & "|" & MyOrder.Strike
            OpQtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Strike",OpSearchString,OpInstance)  
          Case "Put"
            OpInstance = 2
            OpSearchString = MyOrder.Product & "|" & MyOrder.Month & "|" & MyOrder.Strike
            OpQtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Strike",OpSearchString,OpInstance)
          Case "Equity" 
            OpInstance = 1
            OpSearchString = MyOrder.Product
            OpQtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product",OpSearchString,OpInstance)      
            'QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",NewOrder.ProductID,Instance)  
          Case Else
            Log.Error("SubmitOrder : invalid value for CallPut - "&CallPut)
        End Select
  
        'Determine the row of the product and also the TheoCell location
        'QtyCell(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",NewOrder.ProductID,Instance)
  
        OpTheoCell(0) = OpQtyCell(0)
        If OpTheoCell(0) = -1 then 
         	Log.Error("[Fail]:  could not find the Product "& MyOrder.ProductID) 
          Exit Function
        End IF  
        OpTheoCell(1) = GridControl.GetCellColumn(OptionViewGrid.Handle,"Theo",OpInstance)
        OpLastCell = GridControl.GetCellColumn(OptionViewGrid.Handle, "Last", OpInstance) 
        'Current Theo of the Product - handle the error if the theo is not available
  
        Dim OpTheoValue, OpLastValue
        OpTheoValue = GridControl.GetCellText(OptionViewGrid.Handle,OpTheoCell(0),OpTheoCell(1))
        OpLastValue = GridControl.GetCellText(OptionViewGrid.Handle,OpTheoCell(0),OpLastCell)
  
        If OpTheoValue <> "" Then
          OpTheo = StrToFloat(OpTheoValue)
          MyOrder.Theo = OpTheo
        Else 
          log.Error ("The theo for product " & MyOrder.ProductID & " is not available.")
          Exit Function
        End If
  
        If OpLastValue <> "" Then
          OpLast = StrToFloat(OpLastValue)
          MyOrder.LastPrice = OpLast
        Else
          MyOrder.LastPrice = OpTheo
          'Log.Message("No Last Price available - using Theo instead")            
        End If 
    
        'Log.Message("MyOrder.Theo on OrderView =" & MyOrder.Theo) 
    
    End If
    'Debug
    'Log.Message("OverViewTheo: " & OrderViewTheo)
    
    If Left(ExpectedResult,1) = "T" Or Left(ExpectedResult,1) = "t" Then
      'ExpectedPrice = Order.CalculatePriceAmend(ExpectedResult,OrderViewTheo,MyOrder.OrderTickSize)
      'ExpectedPrice = Order.CalculatePriceAmend(LASTPRICE_1,ExpectedResult,OrderViewTheo,MyOrder.OrderTickSize)
      ExpectedPrice = Order.CalculatePriceAmend(MyOrder, ExpectedResult)
    ElseIf Left(ExpectedResult,1) = "P" Or Left(ExpectedResult,1) = "p" Then
      'ExpectedPrice = Order.CalculatePriceAmend(LASTPRICE_1, ExpectedResult,MyOrder.PriceBefore,MyOrder.OrderTickSize)
      ExpectedPrice = Order.CalculatePriceAmend(MyOrder, ExpectedResult)
    ELseIf Left(ExpectedResult,1) = "L" Or Left(ExpectedResult,1) = "l" Then
      Log.Message("ExpectedResult = " & ExpectedResult)
      'ExpectedPrice = Order.CalculatePriceAmend(LASTPRICE_1, ExpectedResult,MyOrder.PriceBefore,MyOrder.OrderTickSize)
      ExpectedPrice = Order.CalculatePriceAmend(MyOrder, ExpectedResult)      
    Else
      ExpectedPrice = ExpectedResult      
    End If

    '19/06/12    
    'ExpectedPrice = Order.CalculatePriceAmend(MyOrder, ExpectedResult)
    
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

'-----------------------------------------------------------------------------------------------
'Sub AmendOrder(MyOrder,Field,Formula)                                                                  
'
'-----------------------------------------------------------------------------------------------
Function AmendOrder(MyOrder,Field,Formula)
  
  Dim NewValue  
  
  If Field = "Price" Then
      'Store the current price before it is amended
      MyOrder.PriceBefore = MyOrder.Price
      '19/06/12
      NewValue = Order.CalculatePriceAmend(MyOrder, Formula)  
      If AmendOrderView(MyOrder.OrderID,Field,NewValue) = True Then
        AmendOrder = True
      End If  
  ElseIf Field = "Qty" Then
    Dim CurrentResidualVolume
    On Error Resume Next
    CurrentResidualVolume = StrToFloat(OrderView.GetColumnValue(MyOrder.OrderID, "Residual Volume"))
      'MS - need convert this to a check to see whether or not Formula is a -+Integer 
      If Formula > 0 Then
        NewValue = CurrentResidualVolume + Formula
      ElseIf Formula < 0 Then
        NewValue = CurrentResidualVolume + Formula     
      Else
        Log.Error("Invalid Qty Amendment: Must be a positive or negative number")
      End If    
    If AmendOrderView(MyOrder.OrderID,Field,NewValue) = True Then
      AmendOrder = True
    End If
  End If

End Function

'-----------------------------------------------------------------------------------------------
'Sub SubmitAmend(MyOrder)                                                                
'
'-----------------------------------------------------------------------------------------------
Sub SubmitAmend(MyOrder)

  Dim SeqNrA, SeqNrB
  
  On Error Resume Next
  SeqNrA = StrToFloat(OrderView.GetColumnValue(MyOrder.OrderID, "SeqNr"))
  
  Call OrderView.PressAmend
  
  'Need to add in a check here as any Invalid Volume amendments cause an error dialog to pop
  'Not sure if this is the best way, but TestComplete allows you to check and wait for a specified time
  'before moving on
  If Aliases.OrderView.WaitAliasChild("dlgSubmitOrderInvalidQuantityError", 200).Exists Then
         'If Aliases.OrderView.dlgSubmitOrderInvalidQuantityError.Exists Then         
          Log.Message("Quantity Amend Failed for Order: " & MyOrder.OrderID)
          Call Aliases.OrderView.dlgSubmitOrderInvalidQuantityError.btnOK.Click
          Call PressCancel
  Else
    'An amendment is expected so check sequence number is updated
    
    If Err = 0 Then
            
      Dim Timeout
      Timeout = DateAdd("s",5,Now)
     
      Do 
        SeqNrB = StrToFloat(OrderView.GetColumnValue(MyOrder.OrderID, "SeqNr"))
        If TimeOut = Now Then
          Log.Message("Timeout Error: It's been 5 seconds and the order has not updated it's sequence number")
          Log.Message("Order has not been amended")
        Exit Do     
        End If            
      Loop Until SeqNrA <> SeqNrB
  
    Else
      Log.Message("The Order """&MyOrder&""" is not exist, Amend cannot be submitted!")  
      
    End If

  End If  
  
  'Once OrderView is updated, update MyOrder.Price so you can pull at price
  MyOrder.Price = StrToFloat(OrderView.GetColumnValue(MyOrder.OrderID,"Price"))
  
  'Debug
  'Log.Message("SubmitAmend: PriceBefore=" & MyOrder.PriceBefore & " Price=" & MyOrder.Price)
  
  'Print Order details to Log  
  LogOrderDetails(MyOrder)
  
  
End Sub

'-----------------------------------------------------------------------------------------------
'Sub PullOrder(MyOrder,GUI_Name)                                                                      
'
'-----------------------------------------------------------------------------------------------

Sub PullOrder(MyOrder,GUI_Name)

  If GUI_Name = "MarketView" Then
    Call PullMarketView(MyOrder)
    Log.Message("Pulling " &MyOrder.OrderName &" from " &GUI_Name)
  ElseIf GUI_Name = "OrderView" Then
    Call OrderView.Pull(MyOrder.OrderID)
    Log.Message("Pulling " &MyOrder.OrderName &" from " &GUI_Name)                            
  Else
    Log.Message("Invalid GUI specified for Pull action, defaulting to OrderView")
    Call OrderView.Pull(MyOrder.OrderID)                                           
  End If

End Sub

'-----------------------------------------------------------------------------------------------
'Sub PullMarketView(MyOrder)                                                                      
'
'-----------------------------------------------------------------------------------------------
Sub PullMarketView(MyOrder)
       
    Call Order.PullAtPrice(MyOrder)

End Sub

Sub MultiUnhold(MultiOrderIDs)
  
  Dim OrderID, FindOrder

  Call OrderView.HighlightOrder(MultiOrderIDs(0), FindOrder, "Left") 
  
  For i =1 to Ubound(MultiOrderIDs) 
     Call OrderView.HighlightOrder(MultiOrderIDs(i), FindOrder, "Ctrl") 
  Next       
  
  Call OrderView.ClickToolbar("Unhold")

  Dim dlgCannotHoldUnhold
  Set dlgCannotHoldUnhold = Aliases.OrderView.dlgCannotHoldUnhold
  
  If dlgCannotHoldUnhold.Exists then 
    dlgCannotHoldUnhold.btnOK.ClickButton
  End If   
    
End Sub


Sub MultiPull(MultiOrderIDs)
  
  Dim OrderID, FindOrder, i

  Call OrderView.HighlightOrder(MultiOrderIDs(0), FindOrder, "Left") 
  
  For i =1 to Ubound(MultiOrderIDs)
    Call OrderView.HighlightOrder(MultiOrderIDs(i), FindOrder, "Ctrl") 
  Next       
  
  Call OrderView.ClickToolbar("Pull")
    
End Sub


Sub UnholdOrder(MyOrder)    

  Dim FindOrder

  FindOrder = True
  
  Call OrderView.HighlightOrder(MyOrder.OrderID, FindOrder, "Left") 
  
  If Not FindOrder Then
    Log.Message("Cannot find the highlighted order on OrderView and exit the "&"Unhold"&" action")
    Exit Sub
  Else  
    Call OrderView.ClickToolbar("Unhold")
  End If
End Sub

Sub LogOrderDetails(MyOrder)

  Dim OrderID
  OrderID = MyOrder.OrderID
  
   'Print Order details to Log  
      Dim AdditionalInfo
    AdditionalInfo = ("Name: " & MyOrder.OrderName & VBNewLine _ 
    & "Order ID: " & OrderID & " " & VBNewLine _
    & "PriceFormula: " & MyOrder.PriceFormula & " " & VBNewLine _
    & "OrderTickSize: " & MyOrder.OrderTickSize & " " & VBNewLine _    
    & "Price: " & MyOrder.Price & VBNewLine _   
    & "Theo (MarketView): " & MyOrder.Theo & VBNewLine _
    & "Theo (OrderView): " & GetColumnValue(OrderID,"Theo") & VBNewLine _
    & "Buy/Sell: " & GetColumnValue(OrderID,"Buy/Sell") & VBNewLine _
    & "Order Status: " & GetColumnValue(OrderID,"Order Status") & VBNewLine _
    & "Volume: " & GetColumnValue(OrderID,"Volume") & VBNewLine _
    & "Residual Volume: " & GetColumnValue(OrderID,"Residual Volume") & VBNewLine _
    & "Executed Volume: " & GetColumnValue(OrderID,"Executed Volume") & VBNewLine _  
    & "Product Name: " & GetColumnValue(OrderID,"Product Name") & VBNewLine _
    & "Price (OrderView): " & GetColumnValue(OrderID,"Price") & VBNewLine _  
    & "Time: " & GetColumnValue(OrderID,"Time") & VBNewLine _ 
    & "SeqNr: " & GetColumnValue(OrderID,"SeqNr") & VBNewLine _ 
    & "Restriction: " & GetColumnValue(OrderID,"Order Restriction") & VBNewLine _ 
    & "Last: " & MyOrder.LastPrice)
    
    Call Log.Message("Order Details: " & MyOrder.OrderName & " (Click Here for Order Details)", AdditionalInfo)

End Sub                       

'--------------------------------------------------------------------------------------------------
'define the Pull Held Order action
'--------------------------------------------------------------------------------------------------

Sub PullHeldOrder(MyOrder)

  Dim FindOrder
  
  FindOrder = True
  
  Call OrderView.HighlightOrder(MyOrder.OrderID, FindOrder, "Left")
  
  If Not FindOrder Then
    Log.Message ("Cannot find the highlighted order on OrderView and exti the Pull action")
    Exit Sub
  Else
    Call OrderView.ClickToolbar("Pull")
  End If
  
End Sub

Function GetOrderNames (OrderNames)

  Dim OrderArray, i, Order, OrderNumber
  OrderArray = split (OrderNames, ",")
  OrderNumber = Ubound(OrderArray)
  
  For i = 0 to OrderNumber
    OrderArray(i) = Trim(OrderArray(i))
  Next  
  
  GetOrderNames = OrderArray

End Function



Sub TradeOut(MyOrder)

  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
  Dim Row, ColBid, ColAsk, Instance
  Dim QuanlityBid, QuanlityAsk, PriceBid, PriceAsk, ColAskQty, BidQty  
  
  Select Case MyOrder.ProductType
      Case "Future", "Strategy", "Call", "Equity", "TMCStrategy" 
       Instance = 1 
        
      Case "Put"
         Instance = 2
          
      Case Else
        Log.Error("Invalid ProductType Specified from Worksheet")   
        Exit Sub
                 
    End Select

  Dim NewTradeOrder, TradePrice, Count
  
  Row = GridControl.GetCellRow(OptionViewGrid.Handle, "ProductID", MyOrder.ProductID, Instance)  
  ColBid = GridControl.GetCellColumn(OptionViewGrid.Handle,"Bid",Instance)
  PriceBid = GridControl.GetCellText(OptionViewGrid.Handle,Row,ColBid)
  ColAsk = GridControl.GetCellColumn(OptionViewGrid.Handle,"Ask",Instance)
  PriceAsk = GridControl.GetCellText(OptionViewGrid.Handle,Row,ColAsk)
  
  Dim OptionClick,FutureClick
  If  PriceBid <> "" or PriceAsk <> "" Then 
    If MyOrder.ProductType = "Equity" or MyOrder.ProductType = "Future" Then
       Set FutureClick = Aliases.MarketView.wndAfx.BCGPDockBar.UnderlyingQuantityToolbar.Edit
        If FutureClick.wText <> "200" Then 
          Call ProductDefaults.ClickQuantity(MyOrder.ProductPath,"200")
        End If  
          
     Else     
       Set OptionClick = Aliases.MarketView.wndAfx.BCGPDockBar.OptionsQuantityToolbar.Edit
       If OptionClick.wText <> "200" Then 
          Call ProductDefaults.ClickQuantity(MyOrder.ProductPath,"200")
       End If
    End If 

  End If  
  
  Count = 1
  Do While (PriceBid <> "" or PriceAsk <> "")  and Count < 50

      If PriceBid <> ""  Then Call ClickGrid(OptionViewGrid, Row, ColBid, "Left") End If
      Delay (100) 
      If PriceAsk <> ""  Then Call ClickGrid(OptionViewGrid, Row, ColAsk, "Left") End If
      Delay (100)
      Count =  Count + 1
      PriceBid = GridControl.GetCellText(OptionViewGrid.Handle,Row,ColBid)  
      PriceAsk = GridControl.GetCellText(OptionViewGrid.Handle,Row,ColAsk)
      
  Loop 

'  Count = 1
'  Do While PriceAsk <> ""  and Count < 50
'
'      Call ClickGrid(OptionViewGrid, Row, ColAsk, "Left") 
'      Delay (100) 
'       
'      Count =  Count + 1
'   btnCancel.ClickButton
'      
'  Loop  
  
  
End Sub



'--------------------------------------------------------------------------------------------------------------------------
'Public Function for get the Last column's value
'--------------------------------------------------------------------------------------------------------------------------

Public Function GrabTheLast(lastprice, NewOrderForLast)
  
  Dim LastCell_1 '(Row,Col) localtion of the Last price cell for the test product
  Dim Last_1
  Dim QtyCellForLast(0)
  Dim Instance_Last 
  
  Select Case NewOrderForLast.ProductType
    Case "Future" 
      Instance_Last = 1
  	  SearchString = TestConfig.FutureProduct & "|" & NewOrderForLast.Month & "|FUTURE"
      QtyCellForLast(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Product Type",SearchString,Instance_Last)
    Case "Strategy","TMCStrategy"
      Instance_Last = 1
  	  QtyCellForLast(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",NewOrderForLast.ProductID,Instance_Last)
    Case "Call"
      Instance_Last = 1
      SearchString = TestConfig.OptionProduct & "|" & NewOrderForLast.Month & "|" & NewOrderForLast.Strike
      QtyCellForLast(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Strike",SearchString,Instance_Last)  
    Case "Put"
      Instance_Last = 2
      SearchString = TestConfig.OptionProduct & "|" & NewOrderForLast.Month & "|" & NewOrderForLast.Strike
      QtyCellForLast(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"Product|Series|Strike",SearchString,Instance_Last)
    Case "Equity" 
      Instance_Last = 1
      QtyCellForLast(0) = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID",NewOrderForLast.ProductID,Instance_Last)  
    Case Else
      Log.Error("SubmitedOrder : invalid value for CallPut - "&CallPut)
  End Select


  LastCell_1 = GridControl.GetCellColumn(OptionViewGrid.Handle, "Last", Instance_Last) 
    'Current Theo of the Product - handle the error if the theo is not available
  
  Dim Last_Value
    
  Last_Value = GridControl.GetCellText(OptionViewGrid.Handle,QtyCellForLast(0),LastCell_1)
  

  If Last_Value <> "" Then
    Last_1 = StrToFloat(Last_Value)
  ElseIf Last_Value = "" Then
    Last_1 = "Blank"        
  End If

  lastprice = Last_1  
  'Log.Message("lastprice = "" Next     


  'Log.Message("lastprice = """&lastprice&"""") 
   
End Function




