''USEUNIT ColumnProperties1
'USEUNIT GlobalDefaults1
'USEUNIT MarketView1
'USEUNIT ProductDefaults1  
'USEUNIT ColumnProperties
'USEUNIT GlobalDefaults
'USEUNIT MarketView
'USEUNIT Order
'USEUNIT MarketView1
'USEUNIT Order1
'USEUNIT OrderView
'USEUNIT ProductDefaults
'USEUNIT TestUtilities
'USEUNIT TheoRefToolbar
'USEUNIT VolsCharting
'USEUNIT ExcelDriver
'USEUNIT QuantSIMAdmin
'USEUNIT Configuration
'USEUNIT Strategy
''USEUNIT MultiPullOrUnhold


Option Explicit


Dim TestConfig
Set TestConfig = ProjectSuite.Variables
Set TestConfig.QuantCOREControl = QuantCOREControl

Dim Vars
Set Vars = Project.Variables

'19/06/12
'Exeuction Flags
TestConfig.fl_PullPreviousOrders = True
TestConfig.fl_delay_OrderView_updates = 10
TestConfig.fl_delay_MV_updates = 50

'Logging formatting
Dim attr
Set attr = Log.CreateNewAttributes  
 
Sub Initialise
  'Vars.xlTestScriptsFile = "c:\Automation\TC_TestScripts_Strategy_XJO.xlsm" 

  'Log.Message("Vars.xlTestScriptsFile = """&Vars.xlTestScriptsFile&"""")
  
  Dim OptionViewGrid, OrderViewGrid
  
  Dim Strategy_Table
  Set Strategy_Table = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, "Strategies", "Strategies")
  
  
  
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid 
  Set OrderViewGrid = Aliases.OrderView.wndAfx.OrderViewGrid
  
  ' If OptionView has been open then assume configuration has been done, though will check all the columns needed are shown in MV
  Call Configuration.ApplicationsInitialise
  
  If Not OptionViewGrid.Exists or Not OrderViewGrid.Exists Then
    Log.Error("There is no valid OptionView or OrderView available to test, please check!")
    Exit Sub
  End If  

  TestConfig.EquityEnabled = False
  If Ucase(TestConfig.UnderlyingProductType) = "EQUITY" Then
    TestConfig.EquityEnabled = True
  End If    
  
  Call Configuration.ProductsTickSize

  Dim Strategies_Array()    
  Call Strategy.ConfigureStrategies(Vars.xlTestScriptsFile, "Strategies", "Strategies", Strategies_Array)

  
  'SFE
  'TestConfig.OptionTickSize = 1
  'TestConfig.FutureTickSize = 0.05
  
  Dim TestRunTbl
  Set TestRunTbl = GetDataTable(TestConfig.xlConfigFile,TestConfig.xlConfigSheet,"TestRun")
   
  Dim TestWorksheetName, OrderTblName, StepsTblName, ExpectedTblName, RowNumber, ColNumber, MultiUnholdTblName, MultiPullTblName
      
  For RowNumber = 2 to TestRunTbl.ListRows.Count + 1
    'Get the Filename, currently don't need Test Script name, also saving this is a Prj var for now
                                                         
    If TestRunTbl.Range.Cells(RowNumber,4).Value = "Yes" Then
      Vars.xlTestScriptsFile = TestRunTbl.Range.Cells(RowNumber,2).Value
      TestWorksheetName = TestRunTbl.Range.Cells(RowNumber,3).Value
      OrderTblName = TestWorksheetName + "_Orders"   
      StepsTblName = TestWorksheetName + "_Steps"
      ExpectedTblName = TestWorksheetName + "_Expected"
     
      Call Log.Checkpoint("|" &TestWorksheetName &"| "&TestRunTbl.Range.Cells(RowNumber,1).Value &" (Execution started)", , , attr)
      
      'Copy the strategy information from the Strategies spread sheet if needed.
      Log.Message("Get the strategy information from the Strategies spread sheet.")
      Call Strategy.CopyStrategy(Strategies_Array, TestWorksheetName, OrderTblName)
      
      'Call MarketView.PullAllOrders
      'Call MarketView1.PullAllOrders     '********************************************added for IntegrationM[1.1]   
     
      Call StartTest(TestWorksheetName,OrderTblName,StepsTblName,ExpectedTblName)   
      
    ' Reset Strategy Table to default value - StrategyID as "N/A" and StrategyName as empty  
      'Call Strategy.ResetStrategyTable

      Call Log.Checkpoint("|Test Script| " & TestRunTbl.Range.Cells(RowNumber,1) & " (Execution complete)", , , attr)
      Log.Checkpoint("------------------------------------------------------------------------------- ")
    
    End If
  Next     

End Sub
 
'-----------------------------------------------------------------------------------------------
'Sub StartTest(xlWorkSheetname,OrderTableName,TestStepsTableName)                                                                
'
'-----------------------------------------------------------------------------------------------
Sub StartTest(TestWorkSheetname,OrderTblName,StepsTblName,ExpectedTblName) 
  
  'Get orders from excel spreadsheet
  'Find order table
  
          
 'Current Implementation: Creating a Dictionary of MarketViewOrder objects
  'Alternatives to consider: A Dictionary of dictionary objects     
  Dim OrderObjDict
  'Set OrderTable = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, TestWorkSheetname, OrderTblName)  
  Set OrderObjDict = CreateObject("Scripting.Dictionary")
  
  Dim RowNumber, ColNumber, CurrentOrderName, ColHeader, ColValue
  Dim TempObjReference, TempObjProperty, OrderTable
  Set OrderTable = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, TestWorkSheetname, OrderTblName)
  'Loop thru the rows
    For RowNumber = 2 to OrderTable.ListRows.Count + 1
      'Get the OrderName in each row
      CurrentOrderName = OrderTable.Range.Cells(RowNumber, 1).Value  
      'Add each Ordername to the Dictionary with a MarketViewOrder object
      OrderObjDict.Add CurrentOrderName, CreateNewOrderTemplate(CurrentOrderName)     
      Set TempObjReference = OrderObjDict.Item(CurrentOrderName)      
      'Debug
      'Log.Message("Processing the following Order: " & CurrentOrderName)
      'Log.Message("TempObjReference Name is: " & TempObjReference.OrderName)
      
      'Loops thru the columns
      For ColNumber = 2 to OrderTable.ListColumns.Count
        ColHeader = OrderTable.ListColumns(ColNumber).Name
        ColValue = OrderTable.Range.Cells(RowNumber,ColNumber).Value 
                     
        'Set the properties of each order
        'Pre-requisite: Class MarketViewOrder variables must match up with the column headers in the spreadsheet table
        'Note: I'm using "" around ColValue because if you pass the ProductType:Call it will screw up as CALL is a native
        'VBScript statement!
        Execute "TempObjReference." & ColHeader & " = " & "ColValue"      
      Next 
    Next
        
  Dim TestCaseTable
  Set TestCaseTable = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, TestWorkSheetname, StepsTblName)
    
  'Iterate through each row of the table and create an internal table
  'Current implementation, decided to go with creating a Dictionary object for each TestCase 
  Dim TestCaseDict
  Set TestCaseDict = CreateObject("Scripting.Dictionary")
 
  Dim TestCaseNumber, CurrentAction, OrderName
  
  For RowNumber = 2 To TestCaseTable.ListRows.Count + 1 
    'Get the TestCaseNumber from each row, this must be unique and also assuming that this is the first column
    TestCaseNumber = TestCaseTable.Range.Cells(RowNumber,1).Value
    TestCaseDict.Add TestCaseNumber, CreateObject("Scripting.Dictionary")   
    
    For ColNumber = 2 To TestCaseTable.ListColumns.Count
      ColHeader = TestCaseTable.ListColumns(ColNumber).Name
      ColValue = TestCaseTable.Range.Cells(RowNumber,ColNumber).Value
      TestCaseDict(TestCaseNumber).Add ColHeader, ColValue 
    Next
  Next 
  
  'Dim ExpectedOutcomesTable
  'Set ExpectedOutcomesTable = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, TestWorkSheetname, ExpectedTblName) 
  
  'Exceute the tests with the Order Dictionary and TestCaseTable
  Call ExecuteTestCases(OrderObjDict,TestCaseDict,ExpectedTblName, TestWorkSheetname, TestCaseTable)
          
End Sub


'-----------------------------------------------------------------------------------------------
'Sub ExecuteTestCases(OrderDict, TestCaseTable)
'
'Description
'
'
'-----------------------------------------------------------------------------------------------
Sub ExecuteTestCases(OrderDict,TestCaseDict,ExpectedTblName, TestWorkSheetname, TestCaseTable)      

  Dim RowNumber, ColNumber, ColHeader, ColValue, MultiOrders, i, Orders, OrderInfo
 
  'Iterate thru the TestCase dictionary using the Keys, which are Test Case Numbers
  Dim TestCaseNumberArray, TestOrderArray, s
  TestCaseNumberArray = TestCaseDict.Keys
  TestOrderArray = OrderDict.Keys
  
  For s = 0 To TestCaseDict.Count - 1
    'Debug print the Test Case Number and the Action required
    'Log.Message("[TestCaseNumber] " & TestCaseNumberArray(s) & " [Action] " & TestCaseDict(TestCaseNumberArray(s)).Item("Action"))
    
    'Process the Current Action
    Dim CurrentTestCase, CurrentOrder, CurrentAction, OrderName, CurrentUser, CurrentTestStep           
    Set CurrentTestCase = TestCaseDict(TestCaseNumberArray(s))   
    CurrentTestStep =  TestCaseNumberArray(s)                  
       
    'Get the Action for the current Test_Case 
    CurrentAction = CurrentTestCase.Item("Action")
    
    'Get the OrderName used for the specific Test_Case and assign the OrderObject to CurrentOrder 
    'If PullPreviousOrders flag is enabled - pulls previous order whenever a new order is to be tested.
    If TestConfig.fl_PullPreviousOrders = True Then
      If Not OrderName = CurrentTestCase.Item("OrderName") Then
            Call MarketView.PullAllOrders
            OrderName = CurrentTestCase.Item("OrderName")
            Delay(100)
      Else
            OrderName = CurrentTestCase.Item("OrderName")
      End If
    Else
      OrderName = CurrentTestCase.Item("OrderName")      
    End If
      
    'Avoiding MultiUser when set to NO
    If TestConfig.MultiUserMode = "No" Then
      CurrentUser = "1"
    ElseIf TestConfig.MultiUserMode = "Yes" Then
      'Get the test user for current Test_Case for test on which MarketView or OrderView
      CurrentUser = CurrentTestCase.Item("User")
    End If
            
    attr.Bold = True
    attr.Italic = True
    Call Log.Checkpoint("Test Step: " & TestCaseNumberArray(s) & " (" & CurrentAction & ")" & " (" &CurrentUser & ")", , , attr)
 
    ' Error handling - muliple orders on the actions other than "multiPull", "PallAll" and "MultiUnhold"
    If Instr(CurrentAction, "Multi") = 0 And Instr(CurrentAction, "All") = 0 Then
      If Ubound(Order.GetOrderNames(OrderName)) = 0 Then
        Set CurrentOrder = OrderDict.Item(OrderName)
      Else 
        Log.Error("Required function can not be operated on " & OrderName) 
        CurrentAction = "Not Apply"
         CurrentOrder.OrderID = "" 
      End If  
    End If
                                             
    Dim NeedHeld, Need, OrderNumber
    NeedHeld = False                                           
    
    If CurrentUser = "1" Then              
      
      'Process the Action  
      Select Case CurrentAction
      
        '---------------------------------------------------------------------------------------              
        Case "Submit Order"
          'Get the Price Formula, and Quantity for the Order
          CurrentOrder.PriceFormula = CurrentTestCase.Item("PriceFormula")          
          CurrentOrder.Quantity = CurrentTestCase.Item("Quantity")
          
          'Submit the Order
          Call Order.Order_MS(CurrentOrder, NeedHeld)             
          Delay(TestConfig.fl_delay_MV_updates)

        '---------------------------------------------------------------------------------------                 
        Case "Clear Orders" 
          Order.Tradeout(CurrentOrder)       

        '---------------------------------------------------------------------------------------     
        Case "Click Trade"                 
      
          Dim CurrentQty
          CurrentQty = CurrentTestCase.Item("Quantity")
          If Not CurrentQty = "" Or CurrentQty < 0 Then
            Call ProductDefaults.ClickQuantity(CurrentOrder.ProductPath,CurrentQty) 
            Call Order.ClickOrderForTrade(CurrentOrder)                   
          End If                            
        
        '---------------------------------------------------------------------------------------               
        Case "Amend"
          'Check the PriceFormula and Quantity Columns to see if the user specified what to amend
          Dim AmendPriceValue, AmendQuantityValue, OrderAmended
          AmendPriceValue = CurrentTestCase.Item("PriceFormula")
          AmendQuantityValue = CurrentTestCase.Item("Quantity")
          
          OrderAmended = False
                        
          'Check if the PriceFormula column is empty, if it isnt, then process the Price Amendment
          If Not AmendPriceValue = "" Then
            If Order.AmendOrder(CurrentOrder,"Price",AmendPriceValue) = True Then
              OrderAmended = True
            End If
          End If       
          
          'Check if the Quantity column is empty, if it isnt, then process the Quantity amendment
          If Not AmendQuantityValue = "" Then         
            If Order.AmendOrder(CurrentOrder,"Qty",AmendQuantityValue) = True Then
              OrderAmended = True
            End If
          End If             

          If OrderAmended = False Then       
            Log.Message("Could not open Amend Dialog. Order Status: " & CurrentOrder.OrderStatus)
          Else
            Call Order.SubmitAmend(CurrentOrder)
            Delay(1000)
          End If

        '---------------------------------------------------------------------------------------     
        Case "Trade"    
          'Get the Price and Qty for the Trade
          Dim TradeQuantity, TradePrice
          TradePrice = CurrentTestCase.Item("PriceFormula")    
          TradeQuantity = CurrentTestCase.Item("Quantity")    
          
          Call Order.TradeOrder(CurrentOrder,TradePrice,TradeQuantity)
          Delay(2000)
        '---------------------------------------------------------------------------------------
        Case "Submit Trade"         
          Dim ExpectedTable_1
          Set ExpectedTable_1 = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, TestWorkSheetname, ExpectedTblName)
        
          Dim ExistingOrderForTradeQuantity
          '19/06/12
          Dim Executed_Price, Executed_Vol
          'Get the Qty for the existing order by either
          '1) Getting the value from the Executed Volume column 
          '2) Subtracting expected Residual Volume from Quantity column if Executed Volume is not available   
           For RowNumber = 2 To ExpectedTable_1.ListRows.Count +1
            If TestCaseNumberArray(s) = ExpectedTable_1.Range.Cells(RowNumber,1).Value Then
              For ColNumber = 2 To ExpectedTable_1.ListColumns.Count
                If ExpectedTable_1.ListColumns(ColNumber).Name = "Residual Volume" Then
                ColValue = ExpectedTable_1.Range.Cells(RowNumber,ColNumber).Value          
                End If
                
                If ExpectedTable_1.ListColumns(ColNumber).Name = "Price" Then
                Executed_Price = ExpectedTable_1.Range.Cells(RowNumber,ColNumber).Value          
                End If
                
                If ExpectedTable_1.ListColumns(ColNumber).Name = "Executed Volume" Then
                Executed_Vol = ExpectedTable_1.Range.Cells(RowNumber,ColNumber).Value          
                End If
                
              Next   
            End if    
          Next
        
          If Executed_Vol > 0 Then
             ExistingOrderForTradeQuantity = CInt(Executed_Vol)
          Else
             ExistingOrderForTradeQuantity = CInt(CurrentTestCase.Item("Quantity")) - CInt(ColValue)
          End If
            
          'Get the Price Formula, and Quantity for the Order
          CurrentOrder.PriceFormula = CurrentTestCase.Item("PriceFormula")                     
          CurrentOrder.Quantity = CurrentTestCase.Item("Quantity")
        
          'Create the existing Market Order        
          Dim ExistingOrderForTrade
          Set ExistingOrderForTrade = CreateTradeOrder(CurrentOrder,ExistingOrderForTradeQuantity)
          '19/06/12
          ExistingOrderForTrade.OrderRestriction = "GFD"
                  
          'Submit both orders        
          Call Order.Order_MS(ExistingOrderForTrade, NeedHeld)          
          CurrentOrder.UnderlyingTheo = ExistingOrderForTrade.UnderlyingTheo          
          
           '19/06/12
           If Executed_Price = "" Then
              CurrentOrder.Price = ExistingOrderForTrade.Price
               
           Else
            CurrentOrder.PriceFormula = Executed_Price
            'CurrentOrder.PriceFormula = CurrentTestCase.Item("PriceFormula")
           End If
            
          Call Order.Order_MS(CurrentOrder, NeedHeld)        
        
        '---------------------------------------------------------------------------------------
        Case "Pull"
        
          If CurrentOrder.Held = "Yes" Then
            Call Order.PullHeldOrder(CurrentOrder)
            Delay(2000)
          Else
            Call PullOrder(CurrentOrder, CurrentTestCase.Item("GUI_Component"))
			
            Delay(2000) 
          End If 
          
        '---------------------------------------------------------------------------------------
        Case "MultiOrderPull"   
          
          MultiOrders = Order.GetOrderNames(CurrentTestCase.Item("OrderName"))
          
          OrderNumber = UBound(MultiOrders)
          Redim OrderIDArray(OrderNumber)
          
          For i = 0 to OrderNumber
           OrderIDArray(i) = OrderDict.Item(MultiOrders(i)).OrderID
          Next 
            
          Call Order.MultiPull(OrderIDArray)
          Delay(2000)
        
        '---------------------------------------------------------------------------------------   
        Case "PullAll"  
'          MultiOrders = TestOrderArray
          MultiOrders = Order.GetOrderNames(CurrentTestCase.Item("OrderName"))
          Select case CurrentTestCase.Item("GUI_Component")
            Case "OrderView"
               Call OrderView.PullAllOrders
            Case "MarketView"
               Call MarketView.PullAllOrders
            Case Else
               Log.Error("Unrecognised Action specified in the Test Steps Table, please check!") 
          End Select    
          Delay(2000)
          
        '---------------------------------------------------------------------------------------
        Case "Unhold"     

          Call Order.UnholdOrder(CurrentOrder)  
          
        '---------------------------------------------------------------------------------------
        Case "MultiOrderUnhold"   
       
          MultiOrders = Order.GetOrderNames(CurrentTestCase.Item("OrderName"))
          OrderNumber = UBound(MultiOrders)
          Redim OrderIDArray(OrderNumber)
          
          For i = 0 to OrderNumber
           OrderIDArray(i) = OrderDict.Item(MultiOrders(i)).OrderID
          Next 
            
          Call Order.MultiUnhold(OrderIDArray)
     
        '---------------------------------------------------------------------------------------
        Case Else
          Log.Message("Unrecognised Action specified in the Test Steps Table, please check that!") 
      End Select
    
      'Verification Process for the Current TestCase
      'Using the CurrentTestCase Number (TestCaseNumberArray(s)), find the row
      
      Dim ExpectedTable
      Set ExpectedTable = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, TestWorkSheetname, ExpectedTblName)
      
      If CurrentOrder.OrderID = "" and CurrentAction <> "Clear Orders" Then
        Log.Error("Cannot verify this order as it does not exist: No OrderID available")
        
      ElseIf  CurrentAction = "Clear Orders"  Then  
        Call Log.Checkpoint("Order Details: " & CurrentOrder.BidAsk & " on the product "& CurrentOrder.ProductID)
        Log.Checkpoint ("The orders on the product has been cleared.")
        
      Else   
        'Log the order details
        Call Log.Checkpoint("Order Details: " & CurrentOrder.BidAsk & " on "& CurrentOrder.ProductID)
      
        For RowNumber = 2 To ExpectedTable.ListRows.Count + 1
          
          If TestCaseNumberArray(s) = ExpectedTable.Range.Cells(RowNumber,1).Value Then
            For ColNumber = 2 To ExpectedTable.ListColumns.Count
              ColHeader = ExpectedTable.ListColumns(ColNumber).Name
              ColValue = ExpectedTable.Range.Cells(RowNumber,ColNumber).Value
              'Log.Message("Expected ColHeader =" & ColHeader)
              'Log.Message("Expected ColValue =" & ColValue)
              
              If Not ColValue = "" Then
                Delay(50)
                Dim FindOrderOnOrderView, HighlightClick
                FindOrderOnOrderView = True
                HighlightClick = "Left"
                Call OrderView.HighlightOrder(CurrentOrder.OrderID, FindOrderOnOrderView, HighlightClick)
                Delay(100)
                
                Call Order.VerifyOrderView(CurrentOrder,CurrentAction,ColHeader,ColValue,OrderDict,MultiOrders)
              End If
            Next   
          End if    
        Next  
  	  End If
    
    ElseIf  CurrentUser = "2" Then
      
      Select Case CurrentAction
        '---------------------------------------------------------------------------------------  
        Case "Submit Order"
          'Get the Price Formula, and Quantity for the Order
          CurrentOrder.PriceFormula = CurrentTestCase.Item("PriceFormula")          
          CurrentOrder.Quantity = CurrentTestCase.Item("Quantity")
          Log.Message("CurrentOrder.ProductType =" & CurrentOrder.ProductType)
          Log.Message("CurrentOrder.Product =" & CurrentOrder.Product)
          Log.Message("CurrentOrder.Month =" & CurrentOrder.Month)
          'Submit the Order
          Call Order1.Order_MS(CurrentOrder, NeedHeld)             
          Delay(TestConfig.fl_delay_MV_updates)                                    
        '---------------------------------------------------------------------------------------
        Case Else
          Log.Message("Unrecognised Action specified in the Test Steps Table, please check that!") 
      End Select
'    
'
'    'Verification Process for the Current TestCase
'    'Using the CurrentTestCase Number (TestCaseNumberArray(s)), find the row
'    
'    Dim ExpectedOutcomesTable
'    Set ExpectedOutcomesTable = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, TestWorkSheetname, ExpectedTblName)
'    
'    If CurrentOrder.OrderID = "" Then
'      Log.Error("Cannot verify this order as it does not exist: No OrderID available")
'    Else   
'    Call LogOrderDetails(CurrentOrder)
'    For RowNumber = 2 To ExpectedOutcomesTable.ListRows.Count + 1
'      If TestCaseNumberArray(s) = ExpectedOutcomesTable.Range.Cells(RowNumber,1).Value Then
'        For ColNumber = 2 To ExpectedOutcomesTable.ListColumns.Count
'          ColHeader = ExpectedOutcomesTable.ListColumns(ColNumber).Name
'          ColValue = ExpectedOutcomesTable.Range.Cells(RowNumber,ColNumber).Value
'    
' 
'
'              If Not ColValue = "" Then
'                Delay(300)
'                Dim FindOrderOnOrderView_1, HighlightClick_1
'                FindOrderOnOrderView_1 = True
'                HighlightClick_1 = "Left"
'                Call OrderView.HighlightOrder(CurrentOrder.OrderID, FindOrderOnOrderView_1, HighlightClick_1)
'                Delay(300)
'                Call Order1.VerifyOrderView(CurrentOrder,CurrentAction,ColHeader,ColValue)
'              End If
'            Next   
'          End if    
'        Next  
'  	  End If
'    
'    Else
'    
'      Log.Error("The User column's value on the testcase table of the TC_TestScripts.xlsx is not 1 or 2 ~!")
'    
    End If
      
  Next 

End Sub


