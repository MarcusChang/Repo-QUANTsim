'USEUNIT TestUtilities
Option Explicit

Dim TestConfig
Set TestConfig = ProjectSuite.Variables

Private GridControl
Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl

Sub Open
  Dim MenuToolbar
  Dim MenuToolbarPopup
  Set MenuToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.MenuBar  
  Dim OptionViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
    
  Call OptionViewGrid.Click(480, 217)
  Call OptionViewGrid.Keys("~vg")

  
  Call TestUtilities.WaitUntilAliasVisible(Aliases.MarketView,"dlgGlobalDefaults",10000)
End Sub

Sub SetStrategyMakerLaunch'(Launch)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Strategies")
  Call dlgGlobalDefaults.TabControl.Strategies.Launch.ClickItem("Right Click")' (Launch)
End Sub
  
Sub SetStrategyMakerNewLegTimeout(Timeout)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Strategies")
  Call dlgGlobalDefaults.TabControl.Strategies.StrategyNewLegTimeout.ClickItem(""&Timeout&"")
End Sub
  
Sub SetTmcMakerNewLegTimeout(Timeout)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Strategies")
  Call dlgGlobalDefaults.TabControl.Strategies.TmcNewLegTimeout.ClickItem(""&Timeout&"")
End Sub

Sub SetSingleClickOrderButtonClick(Button)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Orders")
  Call dlgGlobalDefaults.TabControl.Orders.SingleClickOrder.ClickItem(""&Button&"")
End Sub
  
Sub SetPopupOrderTicketButtonClick(Button)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Orders")
  Call dlgGlobalDefaults.TabControl.Orders.PopupOrderTicket.ClickItem(""&Button&"")
End Sub
  
Sub SetJoinAtPriceButtonClick(Button)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Orders")
  Call dlgGlobalDefaults.TabControl.Orders.JoinAtPrice.ClickItem(""&Button&"")
End Sub
  
Sub SetDimeMarketTicketButtonClick(Button)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Orders")
  Call dlgGlobalDefaults.TabControl.Orders.DimeMarket.ClickItem(""&Button&"")
End Sub
 
Sub SetTickOrderBetterButtonClick(Button)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Orders")
  Call dlgGlobalDefaults.TabControl.Orders.TickOrderBetter.ClickItem(""&Button&"")
End Sub
  
Sub SetTickOrderWorseButtonClick(Button)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Orders")
  Call dlgGlobalDefaults.TabControl.Orders.TickOrderWorse.ClickItem(""&Button&"")
End Sub
 
Sub SetPullOrdersAtPriceButtonClick(Button)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Orders")
  Call dlgGlobalDefaults.TabControl.Orders.PullOrdersAtPrice.ClickItem(""&Button&"")
End Sub
  
Sub OK
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  dlgGlobalDefaults.btnOK.Click
End Sub

Sub SetNotificationsCategories(Value)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Notifications")
  Call dlgGlobalDefaults.TabControl.Notifications.Categories.ClickItem(Value)
End Sub


Sub SetNotificationsActive(Name, Value)
  Dim dlgGlobalDefaults
  Set dlgGlobalDefaults = Aliases.MarketView.dlgGlobalDefaults
  Call dlgGlobalDefaults.TabControl.ClickTab("Notifications")
  
  Dim NotificationGrid, CoordArray, Col, Row
  Set NotificationGrid = Aliases.MarketView.dlgGlobalDefaults.TabControl.Notifications.NotificationGrid
  
  Row = GridControl.GetCellRow(NotificationGrid.Handle,"Name",Name,1)
  Col = GridControl.GetCellColumn(NotificationGrid.Handle,"Active",1)
  CoordArray = Split(GridControl.GetCellCoordinates(NotificationGrid.Handle,Row,Col),"?")
  
  If GetNotificationsActive(Name) <> Value Then
    Log.Message("SetNotificationsActive : Clicking active check box")
    
    Call TestUtilities.ClickGrid(NotificationGrid,Row,Col,"Left") 
    
    If GetNotificationsActive(Name) <> Value Then
      Log.Error("SetNotificationsActive : problem setting active")
    End If
  End If
End Sub

Function GetNotificationsActive(Name)
  Dim PictTicked, PictUnticked, PictCurrent, PictResult, NotificationGrid, i, count, CoordArray, Col, Row
  Set NotificationGrid = Aliases.MarketView.dlgGlobalDefaults.TabControl.Notifications.NotificationGrid
  
  Set PictTicked = Utils.Picture
  Set PictUnticked = Utils.Picture
  Call PictTicked.LoadFromFile("..\BCGticked.bmp")
  Call PictUnticked.LoadFromFile("..\BCGunticked.bmp")
  
  Row = GridControl.GetCellRow(NotificationGrid.Handle,"Name",Name,1)
  If Row = -1 Then
   Log.Error("Did not get a valid row number, column=Name, Value=" &Name)
   Exit Function  
  End If
  
  Col = GridControl.GetCellColumn(NotificationGrid.Handle,"Active",1)
  CoordArray = Split(GridControl.GetCellCoordinates(NotificationGrid.Handle,Row,Col),"?")
  
  Sys.Desktop.MouseX = CoordArray(0)
  Sys.Desktop.MouseY = CoordArray(1)
  
  Set PictCurrent = Sys.Desktop.PictureUnderMouse(9, 8, False)
  
  Call Log.Picture(PictCurrent, "Current "&Name&" Notification state")
  
  Set PictResult = PictTicked.Difference(PictCurrent)
  If PictResult Is Nothing Then
    GetNotificationsActive = True
    Exit Function   
  Else
    Set PictResult = PictUnticked.Difference(PictCurrent)  
    If PictResult Is Nothing Then
      GetNotificationsActive = False
      Exit Function 
    Else
      Log.Error("GetNotificationsActive : state not recognised")
      Exit Function
    End If
  End If
End Function

Sub SetLinkBuyTo(Value)
  Dim TabControl
  Set TabControl=Aliases.MarketView.dlgGlobalDefaults.TabControl
  TabControl.ClickTab("Orders")
  TabControl.Orders.LinkBuyTo.ClickItem(Value)
End Sub

Sub SetUseQtyToolbarForOrderTicket(Value)
  Dim TabControl
  Set TabControl = Aliases.MarketView.dlgGlobalDefaults.TabControl
  
  If TabControl.wTabCaption(TabControl.wFocusedTab) <> "Orders" Then
    TabControl.ClickTab("Orders")
  End If 
  
  Dim btnUseQtyToolbarForOrderTicket 
  Set btnUseQtyToolbarForOrderTicket = Aliases.MarketView.dlgGlobalDefaults.TabControl.Orders.btnUseQtyToolbarForOrderTicket
                           
  If Value = True Then
    Log.Message("Set Use Qty Toolbar For Order Ticket to True")
  ElseIf Value = False Then
    Log.Message("Set Use Qty Toolbar For Order Ticket to False")
  Else
    Log.Message("SetUseQtyToolbarForOrderTicket : value should be True or False, value = "&Value)
  End If
  
  Dim ButtonState
  ButtonState = IsBCGButtonTicked(btnUseQtyToolbarForOrderTicket)
    
  If Value = False And ButtonState = True Then
    Log.Message("Set Use Qty Toolbar For Order Ticket is currently True, clicking the button to set to false")
    Call btnUseQtyToolbarForOrderTicket.ClickButton
  ElseIf Value = True And ButtonState = False Then
    Log.Message("Set Use Qty Toolbar For Order Ticket is currently False, clicking the button to set to true")
    Call btnUseQtyToolbarForOrderTicket.ClickButton
  Else
    Log.Message("Button is already set to correct state")
  End If
End Sub

Sub SetKeepOrderEntryTicketOpen(Value)
  

  Dim KeepOrderEntryTicketOpen
  Set KeepOrderEntryTicketOpen = Aliases.MarketView.dlgGlobalDefaults.TabControl.Orders.btnKeepOrderEntryTicketOpen
  
  'MS - replaced original code
  Dim ButtonState, ButtonOn
  Set ButtonState = KeepOrderEntryTicketOpen.Picture   
  Set ButtonOn = Regions.Items("KeepOrderBtnOn")
  
    If ButtonOff.Check(ButtonState) Then 'Button is OFF
      If Value = "On" Then
        Log.Message("[Keep Order Entry Ticket Open] Button is set to Off... now enabling")
        Call KeepOrderEntryTicketOpen.ClickButton
      End If  
    Else 'Button is ON
      If Value = "Off" Then
        Log.Message("[Keep Order Entry Ticket Open] Button is set to On... now disabling")
        Call KeepOrderEntryTicketOpen.ClickButton
      End If
    
    End If
      
  'Dim ButtonState
  'ButtonState=TestUtilities.IsBCGButtonTicked(KeepOrderEntryTicketOpen)
  'If ButtonState=Value Then
   'Log.Message("Do not need to click button")
   'Else Log.Message("Need to click button")
    'Call KeepOrderEntryTicketOpen.ClickButton
  
End Sub

 Sub SetKeepOrderEntryTicketOpen2
  Dim KeepOrderEntryTicketOpen
  Set KeepOrderEntryTicketOpen=Aliases.MarketView.dlgGlobalDefaults.TabControl.Orders.btnKeepOrderEntryTicketOpen
  
  'MS - replaced original code
  Dim ButtonState, ButtonOn, Value
  Value = "Off"
  Set ButtonState = KeepOrderEntryTicketOpen.Picture   
  Set ButtonOn = Regions.Items("KeepOrderBtnOn")
  
    If Not ButtonOn.Check(ButtonState) Then 'Button is OFF
      If Value = "On" Then
        Log.Message("[Keep Order Entry Ticket Open] Button is set to Off... now enabling")
        Call KeepOrderEntryTicketOpen.ClickButton
      End If  
    Else 'Button is ON
      If Value = "Off" Then
        Log.Message("[Keep Order Entry Ticket Open] Button is set to On... now disabling")
        Call KeepOrderEntryTicketOpen.ClickButton
      End If
    
    End If
 End Sub    
 


 
 
 Function IsBCGButtonTicked' (ButtonObject)
 
  Dim KeepOrderEntryTicketOpen, ButtonObject
  Set ButtonObject = Aliases.MarketView.dlgGlobalDefaults.TabControl.Orders.btnKeepOrderEntryTicketOpen
  
  Dim PictTicked, PictUnticked, PictCurrent, PictResult
  Set PictTicked   = Utils.Picture
  Set PictUnticked = Utils.Picture
  
  PictTicked.LoadFromFile("..\BCGButtonTicked.bmp")
  PictUnticked.LoadFromFile("..\BCGButtonCross.bmp")
  Set PictCurrent = Sys.Desktop.Picture(ButtonObject.ScreenLeft + 6,ButtonObject.ScreenTop + 5, 20, ButtonObject.Height-8, False)
  
  Call Log.Picture(PictCurrent,"Current image for button "&ButtonObject.MappedName)
  Call Log.Picture(PictTicked)
  
  Set PictResult = PictTicked.Compare(PictCurrent)
  log.message(PictResult)
  
  If PictResult then 
    Call KeepOrderEntryTicketOpen.ClickButton 
  End If
  
End Function


Sub CheckAmendError
  Dim AmendErrorPop 
  Set AmendErrorPop = Aliases.OrderView.dlgAmendError
  
  Dim AmendErrorState, AmendError
  Set AmendErrorState = AmendErrorPop.Picture   
  Set AmendError = Regions.Items("dlgAmendError")
  

    
  If AmendError.Check(AmendErrorState) Then 'The AmendError is popped-up
    Call Aliases.OrderView.dlgAmendError.btnOK.Click
    Log.Message("[The Amend Error Occurred] The Held Order Cannot be Amended!!")

  
  Else 'everything is ok
    Log.Message("[There is no Amend Error Occurred]")
  
  End If
  
End Sub

