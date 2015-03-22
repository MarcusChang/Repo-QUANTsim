'USEUNIT TestUtilities
Option Explicit

Dim TestConfig
Set TestConfig = ProjectSuite.Variables

Private GridControl
Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl

Sub Initialize
  Set GridControl = TestConfig.QuantCOREControl 
End Sub

Sub Open
  Dim MenuToolbar
  Dim MenuToolbarPopup
  Set MenuToolbar = Aliases.MarketView1.wndAfx.BCGPDockBar.MenuBar
  Set MenuToolbarPopup = Aliases.MarketView1.wndAfx3.BCGPToolBar40000081000710    '[1.4]%%%%%%%%%%%%%%%%%%%%%%%%%%%change wndAfx1.BCGToolBar40000081001510 to wndAfx3.BCGPToolBar40000081000710 
  
  Dim ViewButtonId
  If MenuToolBar.wButtonText(0) = "" Then
    ViewButtonId = 3
  ElseIf MenuToolBar.wButtonText(0) = "&File" Then
    ViewButtonId = 2
  End If
    
  Call MenuToolbar.ClickItem(ViewButtonId, True)
    
  Call MenuToolbarPopup.ClickItem("P&roduct Defaults", False)
  
  Call TestUtilities.WaitUntilAliasVisible(Aliases.MarketView1,"dlgProductDefaultsEditor",2000)
End Sub

Sub OK
  Dim dlgProductDefaultsEditor
  Set dlgProductDefaultsEditor = Aliases.MarketView1.dlgProductDefaultsEditor
  dlgProductDefaultsEditor.btnOK.Click
End Sub

Sub ClickProduct(Product)
  Dim dlgProductDefaultsEditor
  Set dlgProductDefaultsEditor = Aliases.MarketView1.dlgProductDefaultsEditor
  Call dlgProductDefaultsEditor.ProductList.ClickItem(Product, 0)
End Sub

Sub ClickTab(TabName)
  Dim dlgProductDefaultsEditor
  Set dlgProductDefaultsEditor = Aliases.MarketView1.dlgProductDefaultsEditor
  Call dlgProductDefaultsEditor.TabControl.ClickTab(TabName)
End Sub

Sub ClickMachineTab(TabName)
  Dim MachineTab
  Set MachineTab = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab
  Call MachineTab.ClickTab(TabName)
End Sub

Sub ACRestrictionsClickAC(ClickType)
End Sub

Sub ACRestrictionSetLocation(ClickType,Location)
  Dim AOMGrid, i, count, coordarray, Col
  Set AOMGrid = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab.Restrictions.AOMGrid
  
  count = 0
  For i = 1 To GridControl.GetRowCount(AOMGrid.Handle)
    If TestUtilities.GetTextFromRow(AOMGrid,i,"Click Type",1) = ClickType Then
      Col = GridControl.GetCellColumn(AOMGrid.Handle,"Location",1)
      While TestUtilities.GetTextFromRow(AOMGrid,i,"Location",1) <> Location And count < 10
        Call TestUtilities.ClickGrid(AOMGrid,i,Col,"Double")
        count = count + 1
        Delay(100)   
      Wend
      Exit For
    End If
  Next
  
  If TestUtilities.GetTextFromRow(AOMGrid,i,"Location",1) <> Location Then
    Log.Error("ACRestrictionSetLocation : unable to set Location for "&ClickType&" to "&Location)
  End If
End Sub

Function ACRestrictionsGetAC(ClickType)
  Dim PictTicked, PictUnticked, PictCurrent, PictResult, AOMGrid, i, count, CoordArray, Col, Row
  Set AOMGrid = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab.Restrictions.AOMGrid
  
  Set PictTicked = Utils.Picture
  Set PictUnticked = Utils.Picture
  Call PictTicked.LoadFromFile("..\BCGticked.bmp")
  Call PictUnticked.LoadFromFile("..\BCGunticked.bmp")
  
  Row = GridControl.GetCellRow(AOMGrid.Handle,"Click Type",ClickType,1)
  Col = GridControl.GetCellColumn(AOMGrid.Handle,"AC",1)
  CoordArray = Split(GridControl.GetCellCoordinates(AOMGrid.Handle,Row,Col),"?")
  
  Sys.Desktop.MouseX = CoordArray(0)
  Sys.Desktop.MouseY = CoordArray(1)
  
  Set PictCurrent = Sys.Desktop.PictureUnderMouse(9, 8, False)
  
  Call Log.Picture(PictCurrent, "Current AC state")
  
  Set PictResult = PictTicked.Difference(PictCurrent)
  If PictResult Is Nothing Then
    ACRestrictionsGetAC = True
    Exit Function   
  Else
    Set PictResult = PictUnticked.Difference(PictCurrent)  
    If PictResult Is Nothing Then
      ACRestrictionsGetAC = False
      Exit Function 
    Else
      Log.Error("ACRestrictionsGetAC : AC state not recognised")
      Exit Function
    End If
  End If
End Function

Sub ACRestrictionsSetAC(ClickType, ACState)
  Dim AOMGrid, CoordArray, Col, Row
  Set AOMGrid = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab.Restrictions.AOMGrid
  
  Row = GridControl.GetCellRow(AOMGrid.Handle,"Click Type",ClickType,1)
  Col = GridControl.GetCellColumn(AOMGrid.Handle,"AC",1)
  CoordArray = Split(GridControl.GetCellCoordinates(AOMGrid.Handle,Row,Col),"?")
  
  If ACRestrictionsGetAC(ClickType) <> ACState Then
    Log.Message("Clicking AC enabled check box")
    
    Call TestUtilities.ClickGrid(AOMGrid,Row,Col,"Left") 
    
    If ACRestrictionsGetAC(ClickType) <> ACState Then
      Log.Error("ACRestrictionsSetAC : problem setting AC state")
    End If
  End If
End Sub

Sub ACRestrictionsSetCustomValue(ClickType, Value)
  Dim AOMGrid, Col, Row, Count
  Set AOMGrid = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab.Restrictions.AOMGrid
  
  Row = GridControl.GetCellRow(AOMGrid.Handle,"Click Type",ClickType,1)
  Col = GridControl.GetCellColumn(AOMGrid.Handle,"Custom Value",1)

  Call TestUtilities.ClickGrid(AOMGrid,Row,Col,"Left") 
  
  Call AOMGrid.Keys("[Home]")
  
  Count = 0
  While GridControl.GetCellText(AOMGrid.Handle, Row, Col) <> "" And Count < 20
    Call AOMGrid.Keys("[Del]")
    Count = Count + 1
  Wend
  
  Call AOMGrid.Keys(Value&"[Enter]")
End Sub

Function ACRestrictionsGetCustomValue(ClickType)
  Dim AOMGrid, Col, Row
  Set AOMGrid = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab.Restrictions.AOMGrid
  
  Row = GridControl.GetCellRow(AOMGrid.Handle,"Click Type",ClickType,1)
  Col = GridControl.GetCellColumn(AOMGrid.Handle,"Custom Value",1)

  ACRestrictionsGetCustomValue = GridControl.GetCellText(AOMGrid.Handle, Row, Col)
End Function

Sub SetAOM(SpecName)
  Dim  AOMComboBox
  Set  AOMComboBox = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab.General.AOM
  Call AOMComboBox.ClickItem(SpecName)
End Sub

Sub SetMQ(SpecName)
  Dim  MQComboBox
  Set  MQComboBox = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab.General.MQ
  Call MQComboBox.ClickItem(SpecName)
End Sub

Sub SetTOM(SpecName)
  Dim  TOMComboBox
  Set  TOMComboBox = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.MachineTab.General.TOM
  Call TOMComboBox.ClickItem(SpecName)
End Sub

Sub SetTheoCheckType(Value)
  Dim  TheoCheckType
  Set  TheoCheckType = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.TheoCheckType
  Call TheoCheckType.ClickItem(Value)
End Sub

Sub SetTheoCheckTicks(Value)
  Dim  TheoCheckTicks
  Set  TheoCheckTicks = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.TheoCheckTicks
  Call TheoCheckTicks.Keys("[Home]![End][Del]"&Value) 
End Sub

Sub SetTheoCheckPercent(Value)
  Dim  TheoCheckPercent
  Set  TheoCheckPercent = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.TheoCheckPercent
  Call TheoCheckPercent.Keys("[Home]![End][Del]"&Value)
End Sub

Sub SetTheoCheckWarning(Value)
  Dim  btnWarnings
  Set  btnWarnings = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.btnWarnings
  Call btnWarnings.ClickButton
End Sub


Sub SetClickQuantity(Value)    '******************************[1.0.1] added this sub to set the click quantity value
  Dim QuantityDefaults 
  Set QuantityDefaults = Aliases.MarketView1.dlgProductDefaultsEditor.TabControl.General.BCGPGridCtrl40000081000710
  Call Open
  Call QuantityDefaults.Click(126, 30)
  Call QuantityDefaults.Edit.SetText(Value)
  Call OK
End Sub                         '******************************[1.0.1] added this sub to set the click quantity value 