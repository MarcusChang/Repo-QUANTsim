' Class used for controlling TradeView
'USEUNIT TestUtilities
'USEUNIT MarketView

Private dlgFilters, TradeView
Set dlgFilters = Aliases.TradeView.dlgFilters
Set TradeView = Aliases.TradeView

Public Sub Initialize

End Sub

Sub Login(Username, Password)

  'Run TradeView
  Call TestedApps.TradeView.Run(1, True)
  
  'Wait until Login appears
  Call WaitUntilAliasVisible(TradeView, "dlgLogin", 20000) 
  
  'Click on the username box
  Call TradeView.dlgLogin.Username.Click
  Call TradeView.dlgLogin.Username.Keys("[Home]![End][Del]"&Username)
  Call TradeView.dlgLogin.Password.Keys("[Home]![End][Del]"&Password)
  
  TradeView.dlgLogin.btnOK.ClickButton
  
  Call WaitUntilAliasVisible(TradeView, "wndAfx", 10000)  

End Sub

Public Sub OpenFilters

  TradeView.wndAfx.MainMenu.Click("View|Filters...")

End Sub

Public Sub AddFilters_User(UserGroup,Username)
    
  Dim i, Found
  Found = False
  
  Log.Message("dlgFilters.Users.wItemCount=" & dlgFilters.Users.wItemCount)
    
  For i = 0 To dlgFilters.Users.wItemCount - 1
    If dlgFilters.Users.wItem(i) = UserGroup Then
      Found = True
      Log.Message("Found user group: " & UserGroup)
      dlgFilters.Users.DblClickItem(UserGroup)
      Exit For
    End if
  Next
  
  If Found = False Then
    Log.Error("Could not find user group")
    Exit Sub
  End If  
  
  Found = False
  For i = 0 To dlgFilters.Users.wItemCount - 1
    If dlgFilters.Users.wItem(i) = Username Then
      Found = True
      Log.Message("Found username: " & Username)
      dlgFilters.Users.ClickItem(Username)
      dlgFilters.Users.Keys(" ")
      Exit For
    End if
  Next

End Sub

Public Sub AddFilters_Accounts(AccountName)
      
  Dim i, Found
  Found = False
  
  For i = 0 To dlgFilters.Accounts.wItemCount - 1
    If dlgFilters.Accounts.wItem(i) = AccountName Then
      Found = True
      dlgFilters.Accounts.ClickItem(AccountName)
      Exit For
    End If
  Next
  
  If Found = False Then
    Log.Error("Could not find the specified Account name in TradeView")
    Exit Sub
  End If
  
End Sub

Public Sub OkFilters

       dlgFilters.btnOK.Click

End Sub

