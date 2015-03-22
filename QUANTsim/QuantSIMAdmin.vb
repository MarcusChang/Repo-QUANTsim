'USEUNIT TestUtilities
Option Explicit
Sub Login(Username, Password)
  Call TestedApps.QuantSIMAdmin.Run(1, True)
  
  Call WaitUntilAliasVisible(Aliases.QuantSIMAdmin, "dlgLogin", 15000)
  
  Dim dlgLogin
  Set dlgLogin = Aliases.QuantSIMAdmin.dlgLogin
  
  Call dlgLogin.Username.Click
  Call dlgLogin.Username.Keys("[Home]![End][Del]"&Username)
  Call dlgLogin.Password.Keys("[Home]![End][Del]"&Password)
  
  dlgLogin.btnOK.ClickButton
  
  Call WaitUntilAliasVisible(Aliases.QuantSIMAdmin, "wndAfx", 20000)
End Sub

Function AttSpecExists(AttSpecName)
  Dim TreeControl
  Set TreeControl = Aliases.QuantSIMAdmin.wndAfx.QuantSIMAdmin.TreeControl

  Call TreeControl.Keys("[BS]")
  Call TreeControl.ClickItem("|QuantSIM Administration|ATT Control|New ATT Spec")
  Call TreeControl.Keys(AttSpecName)
  If TreeControl.wSelection = "|QuantSIM Administration|ATT Control|"&AttSpecName Then
    AttSpecExists = True
  Else
    AttSpecExists = False
  End If
End Function

Sub NewATTSpec
  Dim TreeControl
  Set TreeControl = Aliases.QuantSIMAdmin.wndAfx.QuantSIMAdmin.TreeControl

  Call TreeControl.DblClickItem("|QuantSIM Administration|ATT Control|New ATT Spec")
  
  Call WaitUntilAliasVisible(Aliases.QuantSIMAdmin,"dlgEditSpec",5000)
End Sub

Sub SetATTName(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set ATT Spec Name = "&Value)
  
  dlgEditSpec.ATTSpecName.Keys("[Home]![End][Del]"&Value)
End Sub

Sub ExpandATTControl
  Dim TreeControl
  Set TreeControl = Aliases.QuantSIMAdmin.wndAfx.QuantSIMAdmin.TreeControl

  If TreeControl.wSelection <> "|QuantSIM Administration|ATT Control" Then
    Call TreeControl.ExpandItem("|QuantSIM Administration|ATT Control")
  End If
  
  Delay(7000)
End Sub

Sub OpenAttSpec(AttSpecName)
  Dim TreeControl
  Set TreeControl = Aliases.QuantSIMAdmin.wndAfx.QuantSIMAdmin.TreeControl
  
  Call TreeControl.DblClickItem("|QuantSIM Administration|ATT Control|"&AttSpecName)
End Sub
  
Sub Save
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  dlgEditSpec.btnSave.Click
End Sub

Sub Cancel
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
 
  dlgEditSpec.btnCancel.Click
End Sub
  
Sub SetMachine(MachineName)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  dlgEditSpec.Machine.ClickItem(MachineName)
End Sub
  
Sub SetEdgeSpec(EdgeSpec)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set EdgeSpec = "&EdgeSpec)
  
  dlgEditSpec.EdgeSpec.ClickItem(EdgeSpec)
End Sub
  
Sub SetPricingSpec(PricingSpec)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set PricingSpec = "&PricingSpec)  
  
  dlgEditSpec.PricingSpec.ClickItem(PricingSpec)
End Sub
  
Sub SetConnectionType(ConnectionType)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set ConnectionType = "&ConnectionType)
  
  dlgEditSpec.ConnectionType.ClickItem(ConnectionType)
End Sub
  
Sub SetAcMode(AcMode)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set AC Mode = "&AcMode)
  
  dlgEditSpec.AcMode.ClickItem(AcMode)
End Sub
  
Sub SetUserGroup(UserGroup)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set UserGroup = "&UserGroup)    
  
  dlgEditSpec.UserGroup.ClickItem(UserGroup)
End Sub
  
Sub SetUser(User)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set User = "&User)
  
  dlgEditSpec.User.ClickItem(User)
End Sub
 
Sub SetCubeDepth(CubeDepth)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Cube Depth = "&CubeDepth)
  
  dlgEditSpec.CubeDepth.Keys("[Home]![End][Del]"&CubeDepth)
End Sub
  
Sub SetUnderlyingRefType(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Underlying Ref Type = "&Value)
  
  dlgEditSpec.UnderlyingRefType.ClickItem(Value)
End Sub
  
Sub SetSwoModificationPullType(SwoModificationPullType)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  dlgEditSpec.SwoModificationPullType.ClickItem(SwoModificationPullType)
End Sub
  
Sub SetPriceDriver(PriceDriver)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Price Driver= "&PriceDriver)
  
  dlgEditSpec.PriceDriver.ClickItem(PriceDriver)
End Sub
  
Sub SetAsMode(AsMode)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set AS Mode = "&AsMode)
  
  dlgEditSpec.AsMode.ClickItem(AsMode)
End Sub

Sub SetTimedRecalc(State)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  If State = True Then
    Log.Message("Set Timed Recalc = True")
    dlgEditSpec.checkTimedRecalc.ClickButton(cbChecked)
  ElseIf State = False Then
    Log.Message("Set Timed Recalc = False")
    dlgEditSpec.checkTimedRecalc.ClickButton(cbUnChecked)
  Else
    Log.Error("SetTimedRecalc : State value not recognised")
  End If
End Sub

Sub SetTimedRecalcMinutes(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Timed Recalc = "&Value&" minutes")
  
  dlgEditSpec.TimedRecalc.Keys("[Home]![End][Del]"&Value)
End Sub

Sub SetASCancellation(State)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  If State = True Then
    Log.Message("Set AS Cancellation = True")
    dlgEditSpec.checkASCancellation.ClickButton(cbChecked)
  ElseIf State = False Then
    Log.Message("Set AS Cancellation = False")
    dlgEditSpec.checkASCancellation.ClickButton(cbUnChecked)
  Else
    Log.Error("SetASCancellation : State value not recognised")
  End If
End Sub

Sub SetASExpiration(State)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  If State = True Then
    Log.Message("Set AS Expiration = True")
    dlgEditSpec.checkASExpirationDepth.ClickButton(cbChecked)
  ElseIf State = False Then
    Log.Message("Set AS Expiration = False")
    dlgEditSpec.checkASExpirationDepth.ClickButton(cbUnChecked)
  Else
    Log.Error("SetASExpiration : State value not recognised")
  End If
End Sub

Sub SetACEvaluationStep(State)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  If State = True Then
    Log.Message("Set AC Evaluation Step = True")
    dlgEditSpec.checkACEvaluationStep.ClickButton(cbChecked)
  ElseIf State = False Then
    Log.Message("Set AC Evaluation Step = False")
    dlgEditSpec.checkACEvaluationStep.ClickButton(cbUnChecked)
  Else
    Log.Error("SetACEvaluationStep : State value not recognised")
  End If
End Sub

Sub SetAutoHedge(State)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  If State = True Then
    Log.Message("Set Auto Hedge = True")
    dlgEditSpec.checkAutoHedge.ClickButton(cbChecked)
  ElseIf State = False Then
    Log.Message("Set Auto Hedge = False")
    dlgEditSpec.checkAutoHedge.ClickButton(cbUnChecked)
  Else
    Log.Error("SetAutoHedge : State value not recognised")
  End If
End Sub

Sub SetFSAAmendQty(State)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  If State = True Then
    Log.Message("Set FSA Amend Qty = True")
    dlgEditSpec.checkFSAAmendQty.ClickButton(cbChecked)
  ElseIf State = False Then
    Log.Message("Set FSA Amend Qty = False")
    dlgEditSpec.checkFSAAmendQty.ClickButton(cbUnChecked)
  Else
    Log.Error("SetFSAAmendQty : State value not recognised")
  End If
End Sub

Sub SetHedgeName(ProductType,Product,ProductMonth)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
    
  Call dlgEditSpec.btnHedgeName.ClickButton

  Call SelectProducts(Aliases.QuantSIMAdmin.dlgProducts,ProductType,Product,ProductMonth)
End Sub

Sub SetHedgeExchangeAccount(Value)
  Dim  dlgEditSpec
  Set  dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Hedge Exchange Account = "&Value)
  
  Call dlgEditSpec.HedgeExchangeAccount.Keys("[Home]![End][Del]"&Value)
End Sub

Sub SetHedgeExchange(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Hedge Exchange = "&Value)
  
  dlgEditSpec.HedgeExchange.ClickItem(Value)
End Sub

Sub SetHedgeExchangeUser(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Hedge Exchange User = "&Value)
  
  dlgEditSpec.HedgeExchangeUser.ClickItem(Value)
End Sub

Sub SetHedgeAccount(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Hedge Account = "&Value)
  
  dlgEditSpec.HedgeAccount.ClickItem(Value)
End Sub

Sub SetHedgeForceQuantCOREExchangeRelationship(State)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  If State = True Then
    Log.Message("Set Hedge Force QuantCORE Exchange User Relationship = True")
    dlgEditSpec.checkHedgeForceRelationship.ClickButton(cbChecked)
  ElseIf State = False Then
    Log.Message("Set Hedge Force QuantCORE Exchange User Relationship = False")
    dlgEditSpec.checkHedgeForceRelationship.ClickButton(cbUnChecked)
  Else
    Log.Error("SetHedgeForceQuantCOREExchangeRelationship : State value not recognised")
  End If
End Sub

Sub SetTargetName(ProductType,Product)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Call dlgEditSpec.btnTargetName.ClickButton

  Call TestUtilities.SelectProductSet(Aliases.QuantSIMAdmin.dlgProducts,ProductType,Product)
End Sub

Sub SetTargetExchangeAccount(Value)
   Dim  dlgEditSpec
  Set  dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Target Exchange Account = "&Value)
  
  Call dlgEditSpec.TargetExchangeAccount.Keys("[Home]![End][Del]"&Value)
End Sub

Sub SetTargetExchange(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Target Exchange = "&Value)
  
  dlgEditSpec.TargetExchange.ClickItem(Value)
End Sub

Sub SetTargetExchangeUser(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Target Exchange User = "&Value)
  
  dlgEditSpec.TargetExchangeUser.ClickItem(Value)
End Sub

Sub SetTargetAccount(Value)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  Log.Message("Set Target Account = "&Value)
  
  dlgEditSpec.TargetAccount.ClickItem(Value)
End Sub

Sub SetTargetForceQuantCOREExchangeRelationship(State)
  Dim dlgEditSpec
  Set dlgEditSpec = Aliases.QuantSIMAdmin.dlgEditSpec
  
  If State = True Then
    Log.Message("Set Target Force QuantCORE Exchange User Relationship = True")
    dlgEditSpec.checkTargetForceRelationship.ClickButton(cbChecked)
  ElseIf State = False Then
    Log.Message("Set Target Force QuantCORE Exchange User Relationship = False")
    dlgEditSpec.checkTargetForceRelationship.ClickButton(cbUnChecked)
  Else
    Log.Error("SetTargetForceQuantCOREExchangeRelationship : State value not recognised")
  End If
End Sub
