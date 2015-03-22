'USEUNIT ColumnProperties
'USEUNIT GlobalDefaults
'USEUNIT MarketView
'USEUNIT Order
'USEUNIT OrderView
'USEUNIT ProductDefaults
'USEUNIT TestUtilities
'USEUNIT TheoRefToolbar
'USEUNIT VolsCharting
'USEUNIT ExcelDriver
'USEUNIT QuantSIMAdmin
'USEUNIT TradeView
'USEUNIT QuantSIMControl
'USEUNIT VolsCharting
'USEUNIT ColumnProperties1
'USEUNIT GlobalDefaults1
'USEUNIT MarketView1
'USEUNIT ProductDefaults1
                
Option Explicit

Dim TestConfig, GridControl
Set TestConfig = ProjectSuite.Variables
Set TestConfig.QuantCOREControl = QuantCOREControl
Set GridControl = TestConfig.QuantCOREControl

'-------------------------------------------------------------------------------------------------------------------------
'Name: InitialiseGUI(AppName)
'Arguments: AppName[Name of the GUI application to run] 
'Description: 
'This routine calls the Initialisation process for specified GUI
'-------------------------------------------------------------------------------------------------------------------------
Sub InitialiseGUI(AppName)

  Select Case AppName
    Case "All" 
      MarketView_Initialise
      OrderView_Initialise
      TradeView_Initialise
      VolsCharting_Initialise
      QuantSIMControl_Initialise
      QuantSIMAdmin_Intiialise     
    Case "MarketView" 
      MarketView_Initialise
    Case "OrderView" 
      OrderView_Initialise
    Case "TradeView" 
     TradeView_Initialise
    Case "VolsCharting" 
      VolsCharting_Initialise
    Case "QuantSIMControl" 
      QuantSIMControl_Initialise
    Case "QuantSIMAdmin" 
      QuantSIMAdmin_Intiialise
    Case Else 
        Log.Error(AppName & " is not a valid GUI name")    
  End Select 
  
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Name: ConfigureGUI(AppName)
'Arguments: AppName[Name of the GUI application to configure] 
'Description: 
'This routine calls the Configuration routine for specified GUI
'-------------------------------------------------------------------------------------------------------------------------
Sub ConfigureGUI(AppName)

  Select Case AppName
    Case "All" 
      MarketView_Configure
      OrderView_Configure
      TradeView_Configure
      VolsCharting_Configure
      QuantSIMControl_Configure
      QuantSIMAdmin_Intiialise     
    Case "MarketView" 
      MarketView_Configure
    Case "OrderView" 
      OrderView_Configure
    Case "TradeView" 
     TradeView_Configure
    Case "VolsCharting" 
      VolsCharting_Configure
    Case "QuantSIMControl" 
      QuantSIMControl_Configure
    Case "QuantSIMAdmin" 
      QuantSIMAdmin_Intiialise
    Case Else 
        Log.Error(AppName & " is not a valid GUI name")    
  End Select 
  
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Name: LoadGUISheet(AppName,Sheetname)
'Arguments:  
'Description: 
'Load
'-------------------------------------------------------------------------------------------------------------------------
Sub LoadGUISheet(AppName,Sheetname)

  Dim attr
  Set attr = Log.CreateNewAttributes
  attr.Bold = True
  attr.Italic = False 
 

  Select Case AppName
      Case "MarketView" 

        If Aliases.MarketView.wndAfx.MDIClient.OptionView1.WaitAliasChild("OptionViewGrid").Exists = True Then
          ColumnProperties_Config
          Call Log.Checkpoint("The MarketView is already open.", , , attr)
          Exit sub
        End If

        
        Call OpenApplication(AppName)
        
        Call MarketView.OpenSheet(AppName, Sheetname)
        
        If Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid.Exists Then
          Call Log.Checkpoint("The MarketView is loaded and ready for testing.", , , attr)
          Aliases.MarketView.wndAfx.Maximize
          ColumnProperties_Config
        Else 
          Log.Error("Fail to load OptionView file.")  
        End If
        Delay(1000)
      
      Case "MarketView1" 

'       
'        If Sys.Process("MarketView", 2).Exists = True Then 
'      
'          Call Log.Checkpoint("The MarketView1 is already open.", , , attr)
'          Exit sub
'          
'        End If
      
        Call OpenApplication(AppName)   
        
        Call MarketView.OpenSheet(AppName, Sheetname)
        If Aliases.MarketView1.wndAfx.MDIClient.OptionView1.OptionViewGrid.Exists Then
          Call Log.Checkpoint("The second MarketView is loaded and ready for testing.", , , attr)
          Aliases.MarketView1.wndAfx.Maximize
        Else 
          Log.Error("Fail to open the second MarketView, check configuration please!") 
        End If
        Delay(1000) 
      
      
      Case "OrderView"
        'Call OrderView.OpenSheet(Sheetname)
      Case "QuantSIMControl"
        'Call QuantSIMControl.OpenSheet(Sheetname)      
  End Select
    
    
End Sub




'-------------------------------------------------------------------------------------------------------------------------
'Name: SetTestedApps
'Arguments: 
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub SetTestedApps(TestedAppsTable)
  
  'Dim TestedAppsTable
  
  On Error Resume Next
  'Set TestedAppsTable = GetDataTable(xlFilename, xlWorksheetName, xlTableName)
  Log.Message("Name: " & TestedAppsTable.Name)
  
  If Err = 0 Then
  
    'Loop through the TestedApps
    Dim TestApp, i
  
    For i = 0 To TestedApps.Count - 1
      Set TestApp = TestedApps.Items(i)
       Log.Message(TestApp.ItemName)
      
      'Assumption is that the table in Excel is organised as Name, Filename, FilePath, Enabled    
      Dim RowNumber, ColNumber, ColName, CellValue
      For RowNumber = 2 To TestedAppsTable.ListRows.Count + 1
           
        If TestApp.ItemName = TestedAppsTable.Range.Cells(RowNumber,1).Value Then
            
          Set TestApp.Filename = TestedAppsTable.Range.Cells(RowNumber,2).Value
          TestApp.Path = TestedAppsTable.Range.Cells(RowNumber,3).Value         
        
        End If     
      Next 
    Next
      
  Else
    Log.Error("Error setting up TestedApps... check the Excel worksheet")
  End If

End Sub




'*************************************************************************************************************************
'Name: SetTestedAppsForMuT
'Added for IntegrationM[1.1]
'*************************************************************************************************************************

Sub SetTestedAppsForMuT(TestedAppsTableForMuT)
  On Error Resume Next
  Log.Message("Name: " & TestedAppsTableForMuT.Name)
  If Err = 0 Then
    Dim TestAppForMuT, i
    
    For i = 0 To TestedApps.Count - 1
      Set TestAppForMuT = TestedApps.Items(i)
      Log.Message(TestedAppForMuT.ItemName)
      Dim RowNumberForMuT, ColNumberForMuT, ColNameForMuT, CellValueForMuT
      
      For RowNumberForMuT = 2 To TestedAppsTableForMuT.ListRows.Count + 1
      
        If TestAppForMuT.ItemName = TestedAppsTableForMuT.Range.Cells(RowNumberForMuT, 2).Value Then
        
          Set TestAppForMuT.Filename = TestedAppsTableForMuT.Range.Cells(RowNumberForMuT, 3).Value
          TestAppForMuT.Path =  TestedAppsTableForMuT.Range.Cells(RowNumberForMuT, 4).Value
          
        End If
      Next
    Next
    
  Else
    Log.Error("Error when setting up the TestedApps... check the Excel worksheet MultiUserTest table! ")        
  End If
  
End Sub


'*************************************************************************************************************************
'Name: OrderView_InitialiseForMuT
'Added for IntegrationM[1.1]
'*************************************************************************************************************************
Sub OrderView_InitialiseForMuT
  
  Call OrderView.OpenFilters
  Call OrderView.SetFiltersUsers(TestConfig.Username2)
  Call OrderView.FiltersOK

End Sub




'-------------------------------------------------------------------------------------------------------------------------
'Name: RunTestedApps
'Arguments: 
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub RunTestedApps(TestedAppsTable)
   
'  On Error Resume Next
  'Set TestedAppsTable = GetDataTable(xlFilename, xlWorksheetName, xlTableName)
    
'  If Err = 0 Then 
    'Loop thru the table and Initiliase the App        
    'Assumption is that the table in Excel is organised as Name, Filename, FilePath, Enabled    
      Dim RowNumber, ColNumber, ColName, CellValue, App
      
      For RowNumber = 2 To TestedAppsTable.ListRows.Count + 1 
   
        If TestedAppsTable.Range.Cells(RowNumber,4).Value = "Yes" Then
          If TestedAppsTable.Range.Cells(RowNumber,5).Value = "Yes" Then
            App = TestedAppsTable.Range.Cells(RowNumber,1).Value
            Call InitialiseGUI(App)

          ElseIf TestedAppsTable.Range.Cells(RowNumber,5).Value = "No" Then
            Dim Sheetname 
            App = TestedAppsTable.Range.Cells(RowNumber,1).Value
            TestConfig.OptionViewSheetName = TestedAppsTable.Range.Cells(RowNumber,7).Value 
            If TestConfig.OptionViewSheetName <> "N/A" Then
              Sheetname = TestedAppsTable.Range.Cells(RowNumber,8).Value + TestedAppsTable.Range.Cells(RowNumber,7).Value   
              Call LoadGUISheet(App,Sheetname)
            Else  
              Log.Error("Error setting up TestedApps... check the Excel worksheet")
           End If          
        End If 
       End If   
       
       Set TestedAppsTable = ExcelDriver.GetDataTable(TestConfig.xlConfigFile, TestConfig.xlConfigSheet, "TestedApps")   
         
      Next 
               
'  Else
'    Log.Error("Error setting up TestedApps... check the Excel worksheet")
'  End If

End Sub


'-------------------------------------------------------------------------------------------------------------------------
'Name: SetExchangeConfigDetails
'Arguments: 
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub SetExchangeConfigDetails(TestCfgDataTbl)
  
  'Dim TestConfigDataTbl
  
  On Error Resume Next
  'Set TestConfigDataTbl = GetDataTable(xlFilename, xlWorksheetName, xlTableName)
  
  If Err = 0 Then
    
    Dim RowNumber, ColNumber, ColName, CellValue
      'Assumption is that the table in Excel is organised as Item, Filename, Worksheet, Table   
      
      For RowNumber = 2 To TestCfgDataTbl.ListRows.Count + 1
        If TestCfgDataTbl.Range.Cells(RowNumber,1).Value = "Exchange" Then
          For ColNumber = 1 To TestCfgDataTbl.ListColumns.Count               
          
            ColName = TestCfgDataTbl.ListColumns.Item(ColNumber).Name
            
            Select Case ColName
              Case "Filename"
                TestConfig.xlExchangeFile = TestCfgDataTbl.Range.Cells(RowNumber,ColNumber).Value
              Case "Worksheet"
                TestConfig.xlExchangeSheet = TestCfgDataTbl.Range.Cells(RowNumber,ColNumber).Value
              Case "Table"
                TestConfig.xlExchangeTable = TestCfgDataTbl.Range.Cells(RowNumber,ColNumber).Value
            End Select                
      
          Next
        End If 
      Next
      
  Else
    Log.Error("Error setting up TestedApps... check the Excel worksheet")
  End If

End Sub



'-------------------------------------------------------------------------------------------------------------------------
'Name: SetProjectVariables 
'Arguments: 
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub SetProjectVariables(VariablesTable)

  'Dim VariablesTable
    
  On Error Resume Next
  'Set VariablesTable = GetDataTable(xlFilename, xlWorksheetName, xlTableName)
  
  If Err = 0 Then
    Dim RowNumber, ColNumber, VariableName, VariableValue
    For RowNumber = 2 to VariablesTable.ListRows.Count + 1
      'Check if variable exists, if it does, then assign it to a value, if not add it to the collection
      VariableName = VariablesTable.Range.Cells(RowNumber, 1).Value
      VariableValue = VariablesTable.Range.Cells(RowNumber, 2).Value
      
      If TestConfig.VariableExists(VariableName) Then 
        'Variable exists, figure out what type it is defined as and convert it accordingly
        'Debug
        'Log.Message("Project Variable already exists: " & VariableName)
        
        Dim VariableType
        VariableType = TestConfig.GetVariableType(VariableName)
        'Log.Message("Project Variable Type is: " & VariableType)
        
        Select Case VariableType
          Case "Boolean"
            TestConfig.VariableByName(VariableName) = VartoBool(VariableValue)    
          Case "Double"
            TestConfig.VariableByName(VariableName) = VartoFloat(VariableValue)
          Case "Integer"
            TestConfig.VariableByName(VariableName) = VartoInteger(VariableValue)
          Case "String"
            TestConfig.VariableByName(VariableName) = VartoString(VariableValue)
          Case "Object"
            Set TestConfig.VariableByName(VariableName) = Eval(VariableValue)         
        End Select
                          
      Else
        Log.Message("Project Variable does not exist: " & VariableName)
        Log.Message("Adding this as a String for now")
        TestConfig.AddVariable VariableName, "String"
        TestConfig.VariableByName(VariableName) = VartoString(VariableValue)
             
      End If  
    Next
  
    ' String which represents the underlying product name
    'TestConfig.UnderlyingProductName = TestConfig.UnderlyingProduct&" "&TestConfig.UnderlyingShortMonth
    'Added support for Equities 19/01/11
    If TestConfig.UnderlyingProductType = "EQUITY" Or "Equity" Then
      TestConfig.UnderlyingProductID = "SIM.E."&TestConfig.UnderlyingProduct
    Else    
      TestConfig.UnderlyingProductID = "SIM.F."&TestConfig.UnderlyingProduct&"."&TestConfig.UnderlyingMonth
    End If
    
  Else
      Log.Error("Project Variables not set as table could not be found")
  End If

End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Name: MarketView_Initialise
'Arguments: None
'Description: 
'The initialisation process for MarketView e.g. Login, basic setup, non-exchange specific configuration
'-------------------------------------------------------------------------------------------------------------------------
Sub MarketView_Initialise

  'If a MarketView is already open then by default we will try to use that one for test
  
  If Aliases.MarketView.wndAfx.Exists Then
    Log.Message("The MarketView is already open.")
    Exit sub
  End If 
  
  Call MarketView.Login(TestConfig.Username,TestConfig.Password)
  
  Aliases.MarketView.wndAfx.Maximize

  Call MarketView.NewView("OptionView")
  
  Delay(1000)
  
  Call GlobalDefaults_Config
  
  Call ColumnProperties_Config
  
  Call ProductSeletion_Config
  
  Dim attr
  Set attr = Log.CreateNewAttributes
  attr.Bold = True
  attr.Italic = False 
 

  If Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid.Exists Then
    Call Log.Checkpoint("The MarketView is Configured and ready for testing.", , , attr)
  End If  

End Sub 



'-------------------------------------------------------------------------------------------------------------------------
'Name: MarketView_Configure
'Arguments: None
'Description: 
'The Configuration process for MarketView. This is where exchange specific stuff is configured such as Products, Months
'-------------------------------------------------------------------------------------------------------------------------
Sub MarketView_Configure
  'Need to add a check in case object doesn't exist
  
  'Closes any existing OptionView sheets, note that it only works if there's only 1 open, need to update in future for
  'multiple sheets opened.
  If Aliases.MarketView.wndAfx.MDIClient.OptionView1.WaitAliasChild("OptionViewGrid").Visible = True Then
    Call MarketView.CloseSheet
  End If
  
  Call MarketView.NewView("OptionView")
  'Call MarketView.NewView("FloorView")
  Delay(1000)

  MarketView.ClickTab("OptionView1")
  
  'Add the ProductID column into the OptionView sheet
  '1/07/2011 - In future it is probably more convenient to combine this function
  'and only use e.g. "OptionView.EnableColumn("Column name", "Side")
  Call ColumnProperties.Open
  Call ColumnProperties.EnableColumn("Product Definition", "ProductID")
  Call ColumnProperties.EnableColumn("Product Definition", "Product State")
  Call ColumnProperties.EnableColumn("Product Definition", "Product")
  Call ColumnProperties.EnableColumn("Product Definition", "Depth Stack Position")  
  Call ColumnProperties.EnableColumn("Product Definition", "Product Type")
  Call ColumnProperties.SetDecimals("Theo",3)
  
  Call ColumnProperties.SelectTab("Puts")
  Call ColumnProperties.EnableColumn("Product Definition", "ProductID")
  Call ColumnProperties.EnableColumn("Product Definition", "Product State")
  Call ColumnProperties.EnableColumn("Product Definition", "Product")
  Call ColumnProperties.EnableColumn("Product Definition", "Strike")
  Call ColumnProperties.EnableColumn("Product Definition", "Depth Stack Position")
  Call ColumnProperties.EnableColumn("Product Definition", "Series")  
  Call ColumnProperties.EnableColumn("Product Definition", "Product Type")
  Call ColumnProperties.SetDecimals("Theo",3)
  Call ColumnProperties.SetDecimals("Strike",0)
  Call ColumnProperties.Close
  
  
  'Iterate through the Months Table
  Dim x, ProductType, LongMonth
  TestConfig.EquityEnabled = False
  
  Call MarketView.RemoveAllProducts()  
  
  Call MarketView.ProductSelection("Open")
  
  For x = 0 To TestConfig.Months.RowCount - 1
    ProductType = TestConfig.Months.Item(0,x)
    LongMonth =  TestConfig.Months.Item(1,x)  
    
    Dim MyString
    MyString = ProductType & "Product"
    'Call MarketView.AddProduct(ProductType,TestConfig.FutureProduct,LongMonth)
         
      Call MarketView.AddProductMultiples(UCase(ProductType),TestConfig.VariableByName(MyString),LongMonth)
      
      'Check if an Equity is being added
      If ProductType = "Equity" Then
        TestConfig.EquityEnabled = True
      End If
      
    
        
    'If ProductType = "OPTION" Then 
      
      'MarketView.GetOptionStrikes(TestConfig.OptionProduct, TestConfig.LongMonth)
      
      'Dim MyString
      'MyString = LongMonth & "Strikes"
      'Log.Message(MyString)
    'TestConfig.AddVariable VariableName, "String"
    
    'End If
    
  Next

  Call MarketView.ProductSelection("Close")
    
  'Call MarketView.AddProduct("FUTURE",TestConfig.FutureProduct,TestConfig.FutureLongMonth1)
  'Call MarketView.AddProduct("FUTURE",TestConfig.FutureProduct,TestConfig.FutureLongMonth2)      
  'Call MarketView.AddProduct("FUTURE",TestConfig.UnderlyingProduct,TestConfig.UnderlyingLongMonth) 

  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth1)              
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth2)             
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,"STRATEGIES")  
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth1)              
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth2)  
  
    'Add strategies            
  Call MarketView.AddProduct("FUTURE",TestConfig.FutureProduct,"STRATEGIES") 
  Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,"STRATEGIES") 
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth1)              
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth2)             
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,"STRATEGIES")  
  
  'Call ProductDefaults.Open
  
  'Call ProductDefaults.ClickProduct("SIM.F."&TestConfig.FutureProduct&".>")
  
  'Call ProductDefaults.SetAOM(TestConfig.AOMSpec)
  'Call ProductDefaults.SetMQ(TestConfig.MQSpec)
  'Call ProductDefaults.SetTOM(TestConfig.TOMSpec)
  'Call ProductDefaults.SetTheoCheckType("None")
  
  'Call ProductDefaults.ClickProduct("SIM.O."&TestConfig.OptionProduct&".>")
  
  'Call ProductDefaults.SetAOM(TestConfig.AOMSpec)
  'Call ProductDefaults.SetMQ(TestConfig.MQSpec)
  'Call ProductDefaults.SetTOM(TestConfig.TOMSpec)
  'Call ProductDefaults.SetTheoCheckType("None")
  'Call ProductDefaults.OK

  'MarketView.ClickTab("FloorView1")
  
  'Call MarketView.AddProductSet("OPTION",TestConfig.OptionProduct)

  MarketView.ClickTab("OptionView1")
  Delay(1000)

  ' Gets list of strikes for each month
  TestConfig.Month1Strikes = MarketView.GetOptionStrikes(TestConfig.OptionProduct, TestConfig.LongMonth1)
  'TestConfig.Month2Strikes = MarketView.GetOptionStrikes(TestConfig.OptionProduct, TestConfig.LongMonth2)

  'Loop through Months table
  'Get strikes if Product = Option
  'Set it to Strikes table
  'GetStrikes
  
  
  ' Gets the tick size for each of the products
  '-------------------------------------------------------------------------------------------------------------------------
  'TestConfig.FutureTickSize =     GetTickSize("SIM.F."&TestConfig.FutureProduct&"."&TestConfig.LongMonthF)
  'TestConfig.UnderlyingTickSize = GetTickSize(TestConfig.UnderlyingProductID)
  
  'Dim OpTickSz_0, OpTickSz_1, i
  
  'OpTickSz_0 = "^SIM\.O\."&TestConfig.OptionProduct&"\."&TestConfig.LongMonth1&"."&Split(TestConfig.Month1Strikes,"|")(0)&".*C\.0$"
  'OpTickSz_1 = "^SIM\.O\."&TestConfig.OptionProduct&"\."&TestConfig.LongMonth1&"."&Split(TestConfig.Month1Strikes,"|")(0)&".*C\.1$"
  
   'For i = 1 To GridControl.GetRowCount(OptionViewGrid.Handle)
    'If RegExpMatch(OpTickSz_0, GetTextFromRow(OptionViewGrid,i,"ProductID",1)) Or RegExpMatch(OpTickSz_0, GetTextFromRow(OptionViewGrid,i,"ProductID",2)) = True Then
      'TestConfig.OptionTickSize =     GetTickSize("SIM.O."&TestConfig.OptionProduct&"."&TestConfig.LongMonth1&"."&Split(TestConfig.Month1Strikes,"|")(0)&".C.0")
      'Exit For  
   ' ElseIf RegExpMatch(OpTickSz_1, GetTextFromRow(OptionViewGrid, i, "ProductID", 1)) Or RegExpMatch(OpTickSz_0, GetTextFromRow(OptionViewGrid,i,"ProductID",2)) = True Then
      'TestConfig.OptionTickSize =     GetTickSize("SIM.O."&TestConfig.OptionProduct&"."&TestConfig.LongMonth1&"."&Split(TestConfig.Month1Strikes,"|")(0)&".C.1")
      'Exit For
    'End If
  'Next
 '------------------------------------------------------------------------------------------------------------------ 
  'If an equity is being added, get it's tick size. Currently assumes only one equity is being used at a time.
  If TestConfig.EquityEnabled = True Then
    TestConfig.EquityTickSize = GetTickSize("SIM.E."&TestConfig.EquityProduct)
  End If
  

End Sub





'-------------------------------------------------------------------------------------------------------------------------
'Name: OrderView_Initialise
'Arguments: None
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub OrderView_Initialise 

  If Aliases.OrderView.wndAfx.Exists Then
    Log.Message("The OrderView is already open.")
    Exit sub
  End If


  Call OrderView.Login(TestConfig.Username,TestConfig.Password)

  Call OrderView.OpenFilters
  Call OrderView.SetFiltersWorkingCheckbox(True)
  Call OrderView.SetFiltersFilledCheckbox(True)
  Call OrderView.SetFiltersHeldCheckbox(True)
  Call OrderView.SetFiltersOthersCheckbox(True)
  Call OrderView.SetFiltersMassQuoteCheckbox(False)
  Call OrderView.SetFiltersUsers(TestConfig.Username)
  Call OrderView.FiltersOK

End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Name: OrderView_Configure 
'Arguments: 
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub OrderView_Configure
  
  'Set up filters for user and product under test
  'Call OrderView.OpenFilters
  
  'Call OrderView.SetFiltersProduct(TestConfig.UnderlyingProductType,TestConfig.UnderlyingProduct)
  'Call OrderView.SetFiltersProduct("FUTURE",TestConfig.FutureProduct,TestConfig.FutureLongMonth2)
  'Call OrderView.SetFiltersProduct("OPTION",TestConfig.OptionProduct)
  'Call OrderView.SetFiltersProduct("OPTION",TestConfig.OptionProduct,Variables.LongMonth2)  
  'Call OrderView.SetFiltersProduct("OPTION",TestConfig.OptionProduct,Variables.LongMonth2) 
  
  'Allowing strategies
  'Dim dlgFilters,ProductsList
  'Set dlgFilters = Aliases.OrderView.dlgFilters
  'Set ProductsList = Aliases.OrderView.dlgProducts
  'Call dlgFilters.btnAdd.Click
  'Call SelectProducts(ProductsList,"FUTURE",TestConfig.FutureProduct,"STRATEGIES") 
  'Call dlgFilters.btnAdd.Click
  'Call SelectProducts(ProductsList,"OPTION",TestConfig.OptionProduct,"STRATEGIES")
  'Call OrderView.SetFiltersProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth2)  
    
  'Call OrderView.FiltersOK

  'Set up columns
  Dim ColumnNames
  ColumnNames = Array("Product Name","Exchange User","Account","Product","Date","Time","Order Restriction","Product Type"                      ,"Exchange","Order ID","Order Status","Buy/Sell","Price","Volume","Residual Volume","Executed Volume"                     ,"Theo","SeqNr")
  
  Call OrderView.OpenColumnProperties
  Call OrderView.EnableColumns(ColumnNames)
  Call OrderView.SetColumnDecimals("Theo",3) 
  Call OrderView.OKColumnProperties
  
  'Set up sorting
  Call OrderView.SetSorting("Time")
    
  'Add grid lines and stripey 
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Name: TradeView_Initialise
'Arguments: None
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub TradeView_Initialise
    
    Call TradeView.Login(TestConfig.Username,TestConfig.Password)
    
    Call TradeView.OpenFilters
    Call TradeView.AddFilters_User(TestConfig.UserGroup,TestConfig.Username)
    Call TradeView.AddFilters_Accounts(TestConfig.Username)
    Call TradeView.AddFilters_Accounts("xchang2011")
    Call TradeView.OkFilters
    
End Sub 

'-------------------------------------------------------------------------------------------------------------------------
'Name: VolsCharting_Initialise
'Arguments: None
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub VolsCharting_Initialise 
    
    Call VolsCharting.Login(TestConfig.Username, TestConfig.Password)
    Call VolsCharting.SelectPricingSpec(TestConfig.PricingSpec)
    Call VolsCharting.OpenUnderlyingSpec
   
    Call VolsCharting.CheckAdvanced
    Call VolsCharting.SetRollMethod("Auto Rolls")
    Call VolsCharting.SetPricingRuleToHighestPriority("Average") 
    'Need to set Front Month
    Call VolsCharting.SetFrontMonthUnderlyingID("FUTURE",TestConfig.FutureProduct,TestConfig.FutureLongMonth1)
    
    Call VolsCharting.UnderlyingSpecOK
    Call VolsCharting.SetTimeToExpiryAlgorithm("Fixed Time","01/01/2011","10:00:00")
    Call VolsCharting.Minimize
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Name: QuantSIMControl_Initialise
'Arguments: None
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub QuantSIMControl_Initialise
    
  Call QuantSIMControl.Initialize                                  
    
  Call QuantSIMControl.Login("ADMIN","ADMIN")
 
  Call QuantSIMControl.NewView("QuantCOREView")
  Call QuantSIMControl.NewView("LogView")
  Call QuantSIMControl.ClickTab("QuantCOREView")
  Delay(2500)
  
    If QuantSIMControl.GetApplicationStatus("theod",TestConfig.PricingSpec) = "N/A" Then
    Call QuantSIMControl.StartTheod("theod",TestConfig.PricingServer,TestConfig.PricingSpec)
    Delay(1000)
  End If
  
End Sub
  
'-------------------------------------------------------------------------------------------------------------------------
'Name: QuantSIMAdmin_Initialise
'Arguments: None
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub QuantSIMAdmin_Initialise
    
   Call QuantSIMAdmin.Login("ADMIN","ADMIN")  
   Call QuantSIMAdmin.ExpandATTControl 
    
   'Dim TOMSpec
   'Set TOMSpec = GetAttdSpecObject
   'Call GetATTDSpecs(TOMSpec)
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
'Name: SetTestedMonths(MonthsTable)
'Arguments:  
'Description: 
'
'-------------------------------------------------------------------------------------------------------------------------
Sub SetTestMonths(MonthsTable)
 
 On Error Resume Next
  
  If Err = 0 Then
    Dim RowNumber, ColNumber, ProductType, MonthValue, x
    x = 0
    
    TestConfig.Months.RowCount = MonthsTable.ListRows.Count
              
    For RowNumber = 2 To MonthsTable.ListRows.Count + 1
        
      ProductType = MonthsTable.Range.Cells(RowNumber, 1).Value
      MonthValue = MonthsTable.Range.Cells(RowNumber, 2).Value
          
      TestConfig.Months.Item(0,x) = ProductType
      TestConfig.Months.Item(1,x) = MonthValue
      x = x + 1    
    
    Next
      
  Else
      Log.Error("Project Variables not set as table could not be found")
  End If
 
End Sub 




Sub GetStrikes

  Dim MonthsTbl, RowNumber, TotalRows
  Set MonthsTbl = TestConfig.Months
  
  TotalRows = 0
    
  For RowNumber = 0 To MonthsTbl.RowCount - 1
  
    If MonthsTbl.Item("ProductType", RowNumber) = "Option" Then
      
      TotalRows = TotalRows + 1 
      TestConfig.Strikes.RowCount = TotalRows 
      
      Dim LongMonth
      LongMonth = MonthsTbl.Item("Month", RowNumber)
      TestConfig.Strikes.Item(0,(TotalRows - 1)) = LongMonth
      TestConfig.Strikes.Item(1,(TotalRows - 1)) = MarketView.GetOptionStrikes(TestConfig.OptionProduct, LongMonth)

    End If
      
  Next 
  
  
End Sub

Sub ApplicationsInitialise 

  Dim OptionViewGrid, OrderViewGrid
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid 
  Set OrderViewGrid = Aliases.OrderView.wndAfx.OrderViewGrid
  
  Dim xlFilename, xlWorksheetName, xlTableName

 'Load  the VariablesDataTable, set location of Exchange Variables
  
  TestConfig.MultiUserMode = "No"
  
  Dim VariablesDataTbl 
  Set VariablesDataTbl = GetDataTable(TestConfig.xlConfigFile, TestConfig.xlConfigSheet, "TestConfigData")   
  Call SetExchangeConfigDetails(VariablesDataTbl)
  
  Dim ExchangeVarTbl
  Set ExchangeVarTbl = GetDataTable(TestConfig.xlExchangeFile,TestConfig.xlExchangeSheet,TestConfig.xlExchangeTable)
  Call SetProjectVariables(ExchangeVarTbl)
  
  Dim MonthsTbl
  Set MonthsTbl = GetDataTable(TestConfig.xlExchangeFile,TestConfig.xlExchangeSheet,"Months")  
  Call Configuration.SetTestMonths(MonthsTbl)
  
  'Set the TestedApps file locations
  Dim TestedAppsTbl
  Set TestedAppsTbl = GetDataTable(TestConfig.xlConfigFile, TestConfig.xlConfigSheet, "TestedApps")  
  Call Configuration.SetTestedApps(TestedAppsTbl)
  
  'Initialise and Configure any TestedApps that are marked Enabled=Yes
  
  Call Configuration.RunTestedApps(TestedAppsTbl)
 
   'Check the MultiUser table
  Dim MultiUsersTbl
  Set MultiUsersTbl = GetDataTable(TestConfig.xlConfigFile, TestConfig.xlConfigSheet, "MultiUsersTable")
        
  If MultiUsersTbl.Range.Cells(2,1).Value = "Yes" Then
        
        Delay(1000)    
        
        Log.Message("The Test requires multiple users login into MarketView")
'        Log.Message("Begin the initialize for Multiple MarketView and OrderView")
        TestConfig.MultiUserMode = "Yes"
        
'        Dim VariablesDataTbl_1
'        Set VariablesDataTbl_1 = GetDataTable(TestConfig.xlConfigFile, TestConfig.xlConfigSheet, "MultiUsersTable") 
'        Call SetExchangeConfigDetailsForMuT(VariablesDataTbl_1)
  
'        Dim ExchangeVarTbl_1
'        Set ExchangeVarTbl_1 = GetDataTable(TestConfig.xlExchangeFile, TestConfig.xlExchangeSheet, TestConfig.xlExchangeTable)
'        Call SetProjectVariablesForMuT(ExchangeVarTbl_1)
'  
'        Dim MonthsTbl_1
'        Set MonthsTbl_1 = GetDataTable(TestConfig.xlExchangeFile, TestConfig.xlExchangeSheet, "Months")
'        Call Configuration.SetTestMonths(MonthsTbl_1)    
'  
        If Aliases.MarketView1.WaitAliasChild("wndAfx",200).Exists  Then     
         
            Dim attr
            Set attr = Log.CreateNewAttributes  

            attr.Bold = True
            attr.Italic = False 
      
            Call Log.Checkpoint("The second MarketView is already open.", , , attr)
            Exit Sub
            
        End If
        
        Dim TestedAppsTbl_1
        Set TestedAppsTbl_1 = GetDataTable(TestConfig.xlConfigFile, TestConfig.xlConfigSheet, "MultiUsersTable")
        Call Configuration.SetTestedAppsForMuT(TestedAppsTbl_1)
  
  
        Delay(200)
        
           Dim Sheetname, App, UserName, Password
            App = TestedAppsTbl_1.Range.Cells(2,2).Value
            Sheetname = TestedAppsTbl_1.Range.Cells(2,6).Value 
            TestConfig.UserName2 = TestedAppsTbl_1.Range.Cells(2,7).Value
            TestConfig.Password2 = TestedAppsTbl_1.Range.Cells(2,8).Value  
            Call LoadGUISheet(App,Sheetname)
            Call OrderView_InitialiseForMuT
        
  End If

End Sub


Sub GlobalDefaults_Config
  
  GlobalDefaults.Open
  
  ' First reset to defaults
  GlobalDefaults.SetSingleClickOrderButtonClick("None")
  GlobalDefaults.SetPopupOrderTicketButtonClick("None")
  GlobalDefaults.SetJoinAtPriceButtonClick("None")
  GlobalDefaults.SetDimeMarketTicketButtonClick("None")
  GlobalDefaults.SetTickOrderBetterButtonClick("None")
  GlobalDefaults.SetTickOrderWorseButtonClick("None")
  GlobalDefaults.SetPullOrdersAtPriceButtonClick("None")
  
  ' Now set to the values we want
  GlobalDefaults.SetSingleClickOrderButtonClick("Left Click")
  GlobalDefaults.SetPopupOrderTicketButtonClick("Right Click")
  GlobalDefaults.SetTickOrderBetterButtonClick("Fourth Click")
  GlobalDefaults.SetTickOrderWorseButtonClick("Fifth Click")
  GlobalDefaults.SetPullOrdersAtPriceButtonClick("Middle Click")
  GlobalDefaults.SetLinkBuyTo("Bid")
  'GlobalDefaults.SetKeepOrderEntryTicketOpen("Off")
   
  ' Set up strategy launch
  Dim StrategiesLaunch
  Set StrategiesLaunch = Aliases.MarketView.dlgGlobalDefaults.TabControl
  Call StrategiesLaunch.ClickTab("Strategies")                    
  Call StrategiesLaunch.Strategies.Launch.ClickItem("Left Click") 
  Call StrategiesLaunch.Strategies.StrategyNewLegTimeout.ClickItem("Off")
  Call StrategiesLaunch.Strategies.TMCLaunch.ClickItem("Right Click") 
  Call StrategiesLaunch.Strategies.TmcNewLegTimeout.ClickItem("Off")
  Call GlobalDefaults.Ok   
    
  If Aliases.OrderView.WaitAliasChild("dlgStartingOrderViewWorkspace").Exists Then  
    If Aliases.OrderView.WaitAliasChild("dlgStartingOrderViewWorkspace").Visible Then
      Dim dlgStartingOrderViewWorkspace
      Set dlgStartingOrderViewWorkspace = Aliases.OrderView.dlgStartingOrderViewWorkspace
      dlgStartingOrderViewWorkspace.Close
    End If
  End If

  'Call MarketView.CloseSheet
  
End Sub
 
   
Sub ColumnProperties_Config
  'Need to add a check in case object doesn't exist
  
  'Closes any existing OptionView sheets, note that it only works if there's only 1 open, need to update in future for
  'multiple sheets opened.
  
  Call ColumnProperties.Open
  
  Dim MVColumns, Row
  Set MVColumns = ExcelDriver.GetDataTable(TestConfig.xlConfigFile,"MarketView","MarketView_Columns")
  
  For Row = 2 to MVColumns.ListRows.Count + 1 
    Dim Name, Decimals
    Name = MVColumns.Range.Cells(Row, 1).Value
    Decimals = MVColumns.Range.Cells(Row, 2).Value
  
    Select Case Name
    
      case  "Call"
        Call ColumnProperties.SelectTab("Calls")
        
      case  "Put"
        Call ColumnProperties.SelectTab("Puts")  
        
      case  Else
        Call ColumnProperties.EnableColumn("Product Definition", Name)
      
        If Decimals <> "" Then
           Call ColumnProperties.SetDecimals(Name, Decimals)
        End If   
    
    End Select
  
  Next

  Call ColumnProperties.Close
  
  
End Sub


   
Sub ProductSeletion_Config 

'  Dim MonthsTbl
'  Set MonthsTbl = GetDataTable(TestConfig.xlExchangeFile,TestConfig.xlExchangeSheet,"Months")  
'  Call Configuration.SetTestMonths(MonthsTbl)
  
  'Iterate through the Months Table
  Dim x, ProductType, LongMonth
  TestConfig.EquityEnabled = False
  
'  Call MarketView.RemoveAllProducts()  
  
  Call MarketView.ProductSelection("Open")
  
  For x = 0 To TestConfig.Months.RowCount - 1
    ProductType = TestConfig.Months.Item(0,x)
    LongMonth =  TestConfig.Months.Item(1,x)  
    
    Dim MyString
    MyString = ProductType & "Product"
    'Call MarketView.AddProduct(ProductType,TestConfig.FutureProduct,LongMonth)
         
      Call MarketView.AddProductMultiples(UCase(ProductType),TestConfig.VariableByName(MyString),LongMonth)
      
      'Check if an Equity is being added
      If ProductType = "Equity" Then
        TestConfig.EquityEnabled = True
      End If
      
        
    'If ProductType = "OPTION" Then 
      
      'MarketView.GetOptionStrikes(TestConfig.OptionProduct, TestConfig.LongMonth)
      
      'Dim MyString
      'MyString = LongMonth & "Strikes"
      'Log.Message(MyString)
    'TestConfig.AddVariable VariableName, "String"
    
    'End If
    
  Next

  Call MarketView.ProductSelection("Close")
    
  'Call MarketView.AddProduct("FUTURE",TestConfig.FutureProduct,TestConfig.FutureLongMonth1)
  'Call MarketView.AddProduct("FUTURE",TestConfig.FutureProduct,TestConfig.FutureLongMonth2)      
  'Call MarketView.AddProduct("FUTURE",TestConfig.UnderlyingProduct,TestConfig.UnderlyingLongMonth) 

  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth1)              
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth2)             
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,"STRATEGIES")  
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth1)              
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth2)  
  
    'Add strategies            
  Call MarketView.AddProduct("FUTURE",TestConfig.FutureProduct,"STRATEGIES") 
  Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,"STRATEGIES") 
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth1)              
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,TestConfig.LongMonth2)             
  'Call MarketView.AddProduct("OPTION",TestConfig.OptionProduct,"STRATEGIES")  
  
  'Call ProductDefaults.Open
  
  'Call ProductDefaults.ClickProduct("SIM.F."&TestConfig.FutureProduct&".>")
  
  'Call ProductDefaults.SetAOM(TestConfig.AOMSpec)
  'Call ProductDefaults.SetMQ(TestConfig.MQSpec)
  'Call ProductDefaults.SetTOM(TestConfig.TOMSpec)
  'Call ProductDefaults.SetTheoCheckType("None")
  
  'Call ProductDefaults.ClickProduct("SIM.O."&TestConfig.OptionProduct&".>")
  
  'Call ProductDefaults.SetAOM(TestConfig.AOMSpec)
  'Call ProductDefaults.SetMQ(TestConfig.MQSpec)
  'Call ProductDefaults.SetTOM(TestConfig.TOMSpec)
  'Call ProductDefaults.SetTheoCheckType("None")
  'Call ProductDefaults.OK

  'MarketView.ClickTab("FloorView1")
  
  'Call MarketView.AddProductSet("OPTION",TestConfig.OptionProduct)

'  MarketView.ClickTab("OptionView1")
  Delay(1000)
  
End Sub


Sub OpenApplication(AppName)

  Select Case AppName
    
    Case "MarketView" 
      Call MarketView.Login(TestConfig.Username,TestConfig.Password) 
    Case "MarketView1" 
      Call MarketView1.Login(TestConfig.Username2,TestConfig.Password2)   
    Case "OrderView" 
      Call OrderView.Login(TestConfig.Username,TestConfig.Password)   
    Case "TradeView" 
     Call TradeView.Login(TestConfig.Username,TestConfig.Password) 
    Case "VolsCharting" 
      Call VolsCharting.Login(TestConfig.Username,TestConfig.Password) 
    Case "QuantSIMControl" 
      Call QuantSIMControl.Login(TestConfig.Username,TestConfig.Password) 
    Case "QuantSIMAdmin" 
      Call QuantSIMAdmin.Login(TestConfig.Username,TestConfig.Password) 
    Case Else 
        Log.Error(AppName & " is not a valid GUI name")    
        
  End Select 
  
End Sub 



Sub ProductsTickSize

  ' Gets list of strikes for each month
  TestConfig.Month1Strikes = MarketView.GetOptionStrikes(TestConfig.OptionProduct, TestConfig.LongMonth1)
  'TestConfig.Month2Strikes = MarketView.GetOptionStrikes(TestConfig.OptionProduct, TestConfig.LongMonth2)

  'Loop through Months table
  'Get strikes if Product = Option
  'Set it to Strikes table
  'GetStrikes
  
  
  ' Gets the tick size for each of the products
  TestConfig.FutureTickSize =     GetTickSize("SIM.F."&TestConfig.FutureProduct&"."&TestConfig.LongMonthF)
'  TestConfig.UnderlyingTickSize = GetTickSize(TestConfig.UnderlyingProductID)
  
  Dim OptionViewGrid, OpTickSz_0, OpTickSz_1, i
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid
  'Dim ProductID, Row, Col, Strike
  'Strike = Split(TestConfig.Month1Strikes,"|")(0)
  'Row = GridControl.GetCellRow(OptionViewGrid.Handle,"Strike",Strike,1)
  'Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"ProductID",1)
 'ProductID = GridControl.GetCellText(OptionViewGrid.Handle,Row,Col)
  
  
  OpTickSz_0 = "^SIM\.O\."&TestConfig.OptionProduct&"\."&TestConfig.LongMonth1&"."&Split(TestConfig.Month1Strikes,"|")(0)&".*C\.0$"
  OpTickSz_1 = "^SIM\.O\."&TestConfig.OptionProduct&"\."&TestConfig.LongMonth1&"."&Split(TestConfig.Month1Strikes,"|")(0)&".*C\.1$"

  
   For i = 1 To GridControl.GetRowCount(OptionViewGrid.Handle)
    
    If RegExpMatch(OpTickSz_0, GetTextFromRow(OptionViewGrid,i,"ProductID",1)) Or RegExpMatch(OpTickSz_0, GetTextFromRow(OptionViewGrid,i,"ProductID",2)) = True Then
      
      TestConfig.OptionTickSize =     GetTickSize("SIM.O."&TestConfig.OptionProduct&"."&TestConfig.LongMonth1&"."&Split(TestConfig.Month1Strikes,"|")(0)&".C.0")
      
      Exit For  
    ElseIf RegExpMatch(OpTickSz_1, GetTextFromRow(OptionViewGrid, i, "ProductID", 1)) Or RegExpMatch(OpTickSz_0, GetTextFromRow(OptionViewGrid,i,"ProductID",2)) = True Then
      TestConfig.OptionTickSize =     GetTickSize("SIM.O."&TestConfig.OptionProduct&"."&TestConfig.LongMonth1&"."&Split(TestConfig.Month1Strikes,"|")(0)&".C.1")
      Exit For
    End If
  Next
  
  'If an equity is being added, get it's tick size. Currently assumes only one equity is being used at a time.
  If Ucase(TestConfig.UnderlyingProductType) = "EQUITY" Then
    TestConfig.EquityTickSize = GetTickSize("SIM.E."&TestConfig.EquityProduct)
  End If
  

End Sub





