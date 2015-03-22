'USEUNIT TestUtilities

Sub Login(Username, Password)
  Call TestedApps.VolsCharting.Run(1, True)
  
  Call WaitUntilAliasVisible(Aliases.VolsCharting, "dlgLogin", 20000)

  Dim dlgLogin
  Set dlgLogin = Aliases.VolsCharting.dlgLogin
  
  Call dlgLogin.Username.Click
  Call dlgLogin.Username.Keys("[Home]![End][Del]"&Username)
  Call dlgLogin.Password.Keys("[Home]![End][Del]"&Password)
  
  dlgLogin.btnOK.ClickButton
  
  Call WaitUntilAliasVisible(Aliases.VolsCharting, "wndAfx", 20000)
End Sub

Sub SelectPricingSpec(PricingSpecName)
  Dim VolsChartingWindow
  Set VolsChartingWindow = Aliases.VolsCharting.wndAfx
  
  Call VolsChartingWindow.Main.PricingSpecName.ClickItem(PricingSpecName)
  Delay(2500)
End Sub

Sub OpenUnderlyingSpec
  Dim VolsChartingWindow
  Set VolsChartingWindow = Aliases.VolsCharting.wndAfx
  
  Call VolsChartingWindow.Main.btnUnderlyingSpec.ClickButton
 
  Call TestUtilities.WaitUntilAliasVisible(Aliases.VolsCharting,"dlgUnderlyingEditor",10000)
  
  Call TestUtilities.MakeWindowVisible(Aliases.VolsCharting.dlgUnderlyingEditor)
End Sub

Sub DeleteAllUnderlyingProducts
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
  
  For i = 0 To 50
    Call dlgUnderlyingEditor.ProductGrid.Click(30,33)
    Call dlgUnderlyingEditor.btnDelete.ClickButton
  Next
End Sub

Sub AddUnderlyingProduct(OutrightType,OutrightProduct,OutrightMonth,UnderlyingType,UnderlyingProduct,UnderlyingMonth,PriceOffset)
 Dim dlgUnderlyingEditor
 Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
 
 Dim dlgProducts
 Set dlgProducts = Aliases.VolsCharting.dlgProducts
 
 Call dlgUnderlyingEditor.btnOutrightID.ClickButton
 Call SelectProducts(dlgProducts, OutrightType, OutrightProduct, OutrightMonth) 
 
 Call dlgUnderlyingEditor.btnUnderlyingID.ClickButton
 Call SelectProducts(dlgProducts, UnderlyingType, UnderlyingProduct, UnderlyingMonth)
 
 ' Set the price offset
 Call dlgUnderlyingEditor.PriceOffset.Keys(""&PriceOffset&"") 
End Sub

Sub SetRootUnderlyingID(ProductType,Product)
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
 
  Dim dlgProducts
  Set dlgProducts = Aliases.VolsCharting.dlgProducts
 
  Call dlgUnderlyingEditor.btnRootUnderlyingID.ClickButton
  
  Call SelectProductSet(dlgProducts,ProductType,Product)
End Sub

Sub SetFrontMonthUnderlyingID(ProductType,Product,Month)
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
 
  Dim dlgProducts
  Set dlgProducts = Aliases.VolsCharting.dlgProducts
 
  Call dlgUnderlyingEditor.btnFrontMonthUnderlying.ClickButton
  Call SelectProducts(dlgProducts, ProductType, Product, Month) 
End Sub

Sub SetRootDeltaScalingID(ProductType,Product)
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
 
  Dim dlgProducts
  Set dlgProducts = Aliases.VolsCharting.dlgProducts
 
  Call dlgUnderlyingEditor.btnRootDeltaScalingID.ClickButton
  
  Call SelectProductSet(dlgProducts,ProductType,Product)
End Sub

Sub CheckAdvanced
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
  
  Call dlgUnderlyingEditor.checkAdvanced.ClickButton(cbChecked)   
End Sub

Sub UncheckAdvanced
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
  
  Call dlgUnderlyingEditor.check.ClickButton(cbUnChecked)
End Sub

Sub SetRollMethod(RollMethod)
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor  

  Call dlgUnderlyingEditor.RollMethod.ClickItem(RollMethod)
End Sub

Sub SetPricingRuleToHighestPriority(PriceType)
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
  
  Call dlgUnderlyingEditor.ListBox.ClickItem(PriceType) 
  
  For i = 0 To 5
    Call dlgUnderlyingEditor.btnUp.DblClick()
    If dlgUnderlyingEditor.ListBox.wItem(0) = PriceType Then
      Exit For
    End If
  Next
End Sub

Sub UnderlyingSpecOK
  Dim dlgUnderlyingEditor
  Set dlgUnderlyingEditor = Aliases.VolsCharting.dlgUnderlyingEditor
  
  Call dlgUnderlyingEditor.btnOK.ClickButton
End Sub

Sub ClickApply
  Aliases.VolsCharting.wndAfx.Main.btnApply.ClickButton
End Sub

Sub SetTimeToExpiryAlgorithm(TimeToExpiryType, Date, Time)
  Call Aliases.VolsCharting.wndAfx.Main.ComboBox.ClickItem(TimeToExpiryType)
  
  If TimeToExpiryType = "Fixed Time" Then
    Aliases.VolsCharting.wndAfx.Main.DatePick.wDate = Date
    Aliases.VolsCharting.wndAfx.Main.TimePick.wTime = Time
  End If
End Sub

Sub Minimize
  Aliases.VolsCharting.wndAfx.Minimize
End Sub