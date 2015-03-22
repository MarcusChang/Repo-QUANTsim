'USEUNIT TestUtilities
Option Explicit

Dim TestConfig
Set TestConfig = ProjectSuite.Variables

Private BTN_START_APPLICATION
Private BTN_STOP_APPLICATION
Private GridControl

BTN_START_APPLICATION = 32784
BTN_STOP_APPLICATION = 32785
Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl

  Public Sub Initialize
    BTN_START_APPLICATION = 32784
    BTN_STOP_APPLICATION = 32785
    Set GridControl = TestConfig.QuantCOREControl
  End Sub

  Sub Login(Username, Password)
    Call TestedApps.QuantSIMControl.Run(1, True)
    
    Call WaitUntilAliasVisible(Aliases.QuantSIMControl, "dlgLogin", 15000)
  
    Dim dlgLogin
    Set dlgLogin = Aliases.QuantSIMControl.dlgLogin
    
    Call dlgLogin.Username.Click
    Call dlgLogin.Username.Keys("[Home]![End][Del]"&Username)
    Call dlgLogin.Password.Keys("[Home]![End][Del]"&Password)
    
    dlgLogin.ComboBox.ClickItem("New Workspace")
    
    dlgLogin.btnOK.ClickButton
    
    Call WaitUntilAliasVisible(Aliases.QuantSIMControl, "wndAfx", 20000)
  End Sub
  
  Sub NewView(ViewType)
    ' This is not an ideal solution... however, when you open market view the toolbar first looks like this:
    ' File, Edit, View, Window, Help
    ' When you open the first view, they all shift right one, and you get
    ' DOC, File, Edit, View, Window, Help
    ' When you open a second view, the DOC icon disappears again and you get
    ' File, Edit, View, Window, Help
    ' There is not an elegant solution as TestComplete cannot get the text of the buttons in the menu bar
    Dim ViewButtonId
    If Aliases.QuantSIMControl.wndAfx.BCGPDockBar.MenuBar.wButtonText(0) = "" Then
      ViewButtonId = 1
    ElseIf Aliases.QuantSIMControl.wndAfx.BCGPDockBar.MenuBar.wButtonText(0) = "&File" Then
      ViewButtonId = 0
    End If
    
    Call Aliases.QuantSIMControl.wndAfx.BCGPDockBar.MenuBar.ClickItem(ViewButtonId, True)
    
     ' Open the New View control
    Call Aliases.QuantSIMControl.wndAfx1.BCGPToolBar40000081001510.ClickItem("&New"&Chr(9)&"Ctrl+N", False)
    
    Call WaitUntilAliasVisible(Aliases.QuantSIMControl, "dlgNew", 1000)
    
    ' Check if ViewType exists in the list of view types
    Dim ListBox
    Set ListBox = Aliases.QuantSIMControl.dlgNew.ListBox
    If Not ItemInList(ListBox.wItemList,ListBox.wListSeparator,ViewType) Then
      Log.Error("NewView : ViewType is not recognised : "&ViewType)
      Aliases.QuantSIMControl.dlgNew.btnCancel.ClickButton  
      Exit Sub
    End If
    
    Call Aliases.QuantSIMControl.dlgNew.ListBox.ClickItem(ViewType)
    
    Aliases.QuantSIMControl.dlgNew.btnOK.ClickButton
    
    Call WaitUntilAliasVisible(Aliases.QuantSIMControl.wndAfx.MDIClient, ViewType, 2000)
  End Sub
  
  Public Sub ClickTab(TabName)
    Dim QuantSIMControlTab
    Set QuantSIMControlTab = Aliases.QuantSIMControl.wndAfx
    
    Dim ViewNames, CoordArray
    ViewNames = GridControl.GetViewNames(QuantSIMControlTab.Handle)
    
    Dim RegExpObj
    Set RegExpObj = New RegExp
    
    RegExpObj.Pattern = "^"&TabName&"(\*( \([0-9]+ New Msg[s]?\))?)?$"
  
    Dim Item
    Dim Found
    Found = False
    For Each Item In Split(ViewNames,"?")
      If RegExpObj.Test(Item) Then
        Call Log.Message("Clicking on the tab """&Item&"""")
        
        CoordArray = Split(GridControl.GetViewTabCoordinates(QuantSIMControlTab.Handle, Item),"?")

        If UBound(CoordArray) = 1 Then 
          Call QuantSIMControlTab.Click(CoordArray(0)-QuantSIMControlTab.ScreenLeft, CoordArray(1)-QuantSIMControlTab.ScreenTop)
        End If
        Found = True
        Exit For
      End If
    Next
    
    If Not Found Then
      Log.Error("ClickTab : There is not a tab called "&TabName&" available to click")
    End If
  End Sub
  
  Sub StartApplication(ServerName, Host, Username, Password, SpecName)
    Call Aliases.QuantSIMControl.wndAfx.BCGPDockBar.Standard.ClickItem(BTN_START_APPLICATION, False)
    
    Dim dlgStartNewApplication
    Set dlgStartNewApplication = Aliases.QuantSIMControl.dlgStartNewApplication
    
    Call WaitUntilAliasVisible(Aliases.QuantSIMControl, "dlgStartNewApplication", 1000)
    
    Call dlgStartNewApplication.Server.ClickItem(ServerName)

    Call dlgStartNewApplication.Host.Keys("[Home]![End][Del]"&Host)
    
    If ServerName = "attd" Or ServerName = "takeoutd" Then
      Call dlgStartNewApplication.Username.Keys("[Home]![End][Del]"&Username)
      Call dlgStartNewApplication.Password.Keys("[Home]![End][Del]"&Password)
    End If
    
    Call dlgStartNewApplication.ServerSpec.ClickItem(SpecName)
    
    Call dlgStartNewApplication.btnStartApp.ClickButton
  End Sub
  
  Sub StartTheod(ServerName, Host, SpecName)
    Call Aliases.QuantSIMControl.wndAfx.BCGPDockBar.Standard.ClickItem(BTN_START_APPLICATION, False)
    
    Dim dlgStartNewApplication
    Set dlgStartNewApplication = Aliases.QuantSIMControl.dlgStartNewApplication
    
    Call WaitUntilAliasVisible(Aliases.QuantSIMControl, "dlgStartNewApplication", 1000)
    
    Call dlgStartNewApplication.Server.ClickItem(ServerName)
    Call dlgStartNewApplication.Host.ClickItem(Host)
    Call dlgStartNewApplication.ServerSpec.ClickItem(SpecName)
    Call dlgStartNewApplication.btnStartApp.ClickButton
  End Sub
  
  Private Function readTillChar(WshShellExecObj, endChar)
    Dim out, curChar
  
    Do While Not WshShellExecObj.StdOut.AtEndOfStream
      curChar = WshShellExecObj.StdOut.Read(1)
      out = out + curChar
      If (curChar = endChar) Then
        readTillChar = out
        Exit Function
      End If
    Loop
  End Function 
  
  Sub StopApplication(ApplicationType, ApplicationInstance)
    Dim QuantCOREViewGrid
    Set QuantCOREViewGrid = Aliases.QuantSIMControl.wndAfx.MDIClient.QuantCOREView.AfxFrameOrView90

    Dim Row, Col
    Row = GridControl.GetCellRow(QuantCOREViewGrid.Handle,"Application|Instance",ApplicationType&"|"&ApplicationInstance,1)
    Col = GridControl.GetCellColumn(QuantCOREViewGrid.Handle,"Application",1)
    
    If Row = -1 or Col = -1 Then
      Log.Message("Stop Application : "&ApplicationType&" "&ApplicationInstance&" was not found")
      Exit Sub
    End If
    
    Call ClickGrid(QuantCOREViewGrid, Row, Col, "Left")
    
    Call Aliases.QuantSIMControl.wndAfx.BCGPDockBar.Standard.ClickItem(BTN_STOP_APPLICATION, False)
    
    Call WaitUntilAliasVisible(Aliases.QuantSIMControl, "dlgStopApplications", 1000)
    
    Call Aliases.QuantSIMControl.dlgStopApplications.btnOK.ClickButton
  End Sub
  
  Function GetApplicationStatus(ApplicationType, ApplicationInstance)
    Dim QuantCOREViewGrid
    Set QuantCOREViewGrid = Aliases.QuantSIMControl.wndAfx.MDIClient.QuantCOREView.AfxFrameOrView90
    
    Dim Row, Col
    Row = GridControl.GetCellRow(QuantCOREViewGrid.Handle,"Application|Instance",ApplicationType&"|"&ApplicationInstance,1)
    If Row > -1 Then
      Col = GridControl.GetCellColumn(QuantCOREViewGrid.Handle,"Status",1)
      GetApplicationStatus = GridControl.GetCellText(QuantCOREViewGrid.Handle,Row,Col)
    Else
      GetApplicationStatus = "N/A"
    End If
  End Function
  
  Function GetApplicationFTStatus(ApplicationType, ApplicationInstance)
    Dim QuantCOREViewGrid
    Set QuantCOREViewGrid = Aliases.QuantSIMControl.wndAfx.MDIClient.QuantCOREView.AfxFrameOrView90
    
    Dim Row, Col
    Row = GridControl.GetCellRow(QuantCOREViewGrid.Handle,"Application|Instance",ApplicationType&"|"&ApplicationInstance,1)
    Col = GridControl.GetCellColumn(QuantCOREViewGrid.Handle,"FT Status",1)
    GetApplicationFTStatus = GridControl.GetCellText(QuantCOREViewGrid.Handle,Row,Col)
  End Function
