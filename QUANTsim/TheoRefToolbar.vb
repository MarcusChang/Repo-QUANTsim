Option Explicit

Public Sub Open
  Dim TheoRefToolbar
  Set TheoRefToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar
  Call TheoRefToolbar.ClickItem(34890, False) 
  Delay(100)
End Sub

Public Sub SetUnderlying(ProductTickRatioNumerator, ProductTickRatioDenominator, ValueType)
  Dim RefModifiers
  Set RefModifiers = Aliases.MarketView.dlgConfigureReferenceModifiers.TabControl.RefModifiers       
  Call RefModifiers.ProductRatioNum.Keys("[Home]![End][Del]"&ProductTickRatioNumerator)
  Call RefModifiers.ProductRatioDenom.Keys("[Home]![End][Del]"&ProductTickRatioDenominator)
  Call RefModifiers.UnderlyingValueType.ClickItem(ValueType)
End Sub
  
Public Sub SetInterestRate(StepSize, ValueType, UnitType)
  Dim dlgConfigureReferenceModifiers
  Set dlgConfigureReferenceModifiers = Aliases.MarketView.dlgConfigureReferenceModifiers.TabControl.RefModifiers          
  dlgConfigureReferenceModifiers.InterestRateStepSize.Keys("[Home]![End][Del]"&StepSize)
  Call dlgConfigureReferenceModifiers.InterestRateValueType.ClickItem(ValueType)
  If dlgConfigureReferenceModifiers.InterestRateUnit.Enabled Then
    Call dlgConfigureReferenceModifiers.InterestRateUnit.ClickItem(UnitType)
  End If
End Sub
  
Public Sub SetVolatility(StepSize, ValueType, UnitType)
  Dim dlgConfigureReferenceModifiers
  Set dlgConfigureReferenceModifiers = Aliases.MarketView.dlgConfigureReferenceModifiers.TabControl.RefModifiers              
  dlgConfigureReferenceModifiers.VolatilityStepSize.Keys("[Home]![End][Del]"&StepSize)
  Call dlgConfigureReferenceModifiers.VolatilityValueType.ClickItem(ValueType)
  Call dlgConfigureReferenceModifiers.VolatilityUnit.ClickItem(UnitType)
End Sub
  
Public Sub SetTimeToExpiry(StepSize, ValueType, UnitType)
  Dim dlgConfigureReferenceModifiers
  Set dlgConfigureReferenceModifiers = Aliases.MarketView.dlgConfigureReferenceModifiers.TabControl.RefModifiers              
  dlgConfigureReferenceModifiers.TimeToExpiryStepSize.Keys("[Home]![End][Del]"&StepSize)
  Call dlgConfigureReferenceModifiers.TimeToExpiryValueType.ClickItem(ValueType)
  Call dlgConfigureReferenceModifiers.TimeToExpiryUnit.ClickItem(UnitType)
End Sub
  
Public Sub SetRefType(RefType) 
  Dim TheoRefToolbar                                                                                          
  Set TheoRefToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar  
  Call TheoRefToolbar.ClickItem(34885, False)    
  TheoRefToolbar.Keys("[Enter]") ' This is click twice to close the box before you use the cursor
  ' to select a value.  A bug in the BCG control means you cannot move the cursor and click Enter to select something
   
  Dim SelectedIndex
  Dim ComboBoxList
  Dim ComboBoxListArray
  SelectedIndex     = TheoRefToolbar.ComboBox.wSelectedItem
  ComboBoxList      = TheoRefToolbar.ComboBox.wItemList
  ComboBoxListArray = Split(ComboBoxList, vbCrLf)
    
  ' Work out where in the list the option we want is
  Dim RequiredIndex
  RequiredIndex = 0
  For Each c In ComboBoxListArray
    If c = RefType Then
      Exit For
    End If
    RequiredIndex=RequiredIndex+1         
  Next
    
  ' Use the cursor keys to select the value we want
  If RequiredIndex-SelectedIndex > 0 Then
    For i = 0 To (RequiredIndex-SelectedIndex)-1
      Call TheoRefToolbar.Keys("[Down]")
    Next
  ElseIf  SelectedIndex-RequiredIndex > 0 Then
    For i = 0 To (SelectedIndex-RequiredIndex)-1
      Call TheoRefToolbar.Keys("[Up]")
    Next
  End If
End Sub

  ' Set ref value, making sure nothing is in field before editing  
Public Sub SetRefValue(RefValue)
  Dim TheoRefToolbar
  Set TheoRefToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar
  TheoRefToolbar.RefValue.Keys("[Home]![End][Del]"&RefValue&"[Enter]")
End Sub
  
' This gets the reference value from the reference toolbar.  It also does a check that it is in the correct format
Public Function GetRefValue(UnitType)
  Dim StrRefValue
  StrRefValue = Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar.RefValue.wText
  Set RegExpObj = New RegExp
  Select Case UnitType
  Case "Fixed"
    RegExpObj.Pattern = "^(-?[0-9]+\.?[0-9]*)$"
  Case "Percentage"
    RegExpObj.Pattern = "^(-?[0-9]+\.?[0-9]*)( %%)$"
  Case "Percentage"
    RegExpObj.Pattern = "^(-?[0-9]+\.?[0-9]*)( days)$"
  Case Else
    Log.Error("GetRefValue(UnitType) - invalid value for UnitType : "&UnitType)
  End Select
  If RegExpObj.Test(StrRefValue) Then
    Log.Message("Format of Theo Reference value is OK")
    GetRefValue = RegExpObj.Execute(StrRefValue)(0).SubMatches(0)
  Else
    Log.Error("Test against pattern """&RegExpObj.Pattern&""" StrRefValue = """&StrRefValue&"""   Matches="&RegExpObj.Test(StrRefValue))
  End If  
End Function  
  
Public Sub ScrollRefValue(Direction, NumberOfSteps)
  Call Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar.RefValue.Click    
  Select Case Direction
  Case "Up"
    Call Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar.RefValue.MouseWheel(NumberOfSteps)
  Case "Down"
    Call Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar.RefValue.MouseWheel(0-NumberOfSteps)
  Case Else
    Log.Error("ScrollRefValue(Direction) - Invalid value : "&Direction)
  End Select
End Sub
  
Public Sub Calculate
  Dim TheoRefToolbar
  Set TheoRefToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar     
  Call TheoRefToolbar.ClickItem(34887, False)
End Sub
  
Public Sub Reset
  Dim  TheoRefToolbar
  Set  TheoRefToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar     
  Call TheoRefToolbar.ClickItem(34888, False)
End Sub
  
Public Sub Scenarios(Value)
  Dim TheoRefToolbar
  Set TheoRefToolbar = Aliases.MarketView.wndAfx.BCGPDockBar.TheoRefToolbar
  Select Case Value
  Case True
  Case False     
    Call TheoRefToolbar.CheckItem(34889, Value, False)
  Case Else
    Log.Error("TheoRefConfig.Scenarios Invalid Value : "&Value)
  End Select
End Sub
  
Public Sub SelectProduct(ProductID)
  Dim CoordString, CoordArray
  CoordString = StingrayGrid.GetGridCoordinates("ProductID",ProductID,"Strike","Call")
  CoordArray  = Split(CoordString, ",")
  Call StingrayGrid.Click("Left", CoordArray(0), CoordArray(1))
  Delay(100)
End Sub
  
Public Sub OK
  Dim dlgConfigureReferenceModifiers
  Set dlgConfigureReferenceModifiers = Aliases.MarketView.dlgConfigureReferenceModifiers
  Call dlgConfigureReferenceModifiers.btnOK.ClickButton
  Delay(100)
End Sub

Public Sub Cancel
  Dim dlgConfigureReferenceModifiers
  Set dlgConfigureReferenceModifiers = Aliases.MarketView.dlgConfigureReferenceModifiers
  Call dlgConfigureReferenceModifiers.btnCancel.ClickButton
  Delay(100)
End Sub