'USEUNIT TestUtilities 
'USEUNIT ExcelDriver 

Option Explicit
'Logging formatting
Dim attr
Set attr = Log.CreateNewAttributes  

Dim TestConfig
Set TestConfig = ProjectSuite.Variables
Set TestConfig.QuantCOREControl = QuantCOREControl

Dim Vars
Set Vars = Project.Variables


  Dim OptionViewGrid, GridControl,dlgStrategyMaker, StrategyGrid, dlgTmcMaker
  Set OptionViewGrid = Aliases.MarketView.wndAfx.MDIClient.OptionView1.OptionViewGrid 
  Set GridControl = TestConfig.QuantCOREControl
  Set dlgStrategyMaker = Aliases.MarketView.dlgStrategyMaker
  Set StrategyGrid = Aliases.MarketView.dlgStrategyMaker.Strategy 
  Set dlgTmcMaker = Aliases.MarketView.dlgTmcMaker
  
Class StrategyDetails

    Public StrategyNumber
    Public StrategyName
    Public StrategyTable
    Public StrategyID 
    Public StrategyMethod 
    
End Class  

Function ConfigureStrategies(xlFilename, xlWorksheetName, xlTableName, Strategies)

  Dim StrategyTable
  Set StrategyTable = ExcelDriver.GetDataTable(xlFilename, xlWorksheetName, xlTableName)
  
  Dim  StrategyRows, Row, i
  StrategyRows = StrategyTable.ListRows.Count    
 
  Redim Strategies(StrategyRows)
  
  attr.Bold = True
  attr.Italic = False 
      
  Call Log.Checkpoint("Creating the strategies for test suite if needed.", , , attr)
  
  For i = 0 to UBound(Strategies)
      Set  Strategies(i) = New StrategyDetails
     
  Next  
                                                                                                        
  For Row = 2 to StrategyRows + 1
    If StrategyTable.Range.Cells(Row, 5).Value = "" Then
          Strategies(Row -1).StrategyNumber = StrategyTable.Range.Cells(Row, 1).Value
          Strategies(Row -1).StrategyTable = StrategyTable.Range.Cells(Row, 2).Value
'         CreateStrategy.StrategyType(Row -1) =  StrategyTable.Range.Cells(Row, 4).Value
          If StrategyTable.Range.Cells(Row, 3).Value = "TMC Maker" Then
            Strategies(Row -1).StrategyMethod = "TMC Maker"
          Else  Strategies(Row -1).StrategyMethod =  "Strategy" 
          End If
     End If 
  Next

  For i =1 to StrategyRows
    If Strategies(i).StrategyNumber <> "" Then
       Log.Message(Strategies(i).StrategyNumber)
       Log.Message(Strategies(i).StrategyName)
       Log.Message(Strategies(i).StrategyTable)
       Log.Message(Strategies(i).StrategyID)
       Log.Message(Strategies(i).StrategyMethod)
       Call Strategy.CreateStrategy(Strategies(i))   
    End If
  Next

  Set StrategyTable = ExcelDriver.GetDataTable(xlFilename, xlWorksheetName, xlTableName)   
  For  Row = 2 to StrategyRows + 1
        Strategies(Row-1).StrategyNumber = StrategyTable.Range.Cells(Row,1).Value      
        Strategies(Row-1).StrategyName = StrategyTable.Range.Cells(Row,4).Value  
        Strategies(Row-1).StrategyID = StrategyTable.Range.Cells(Row,5).Value 
  Next    
  
   
End Function


'Create the strategy according to the Orders table  
Function CreateStrategy(Strategy)

  Dim Strategylegs(), Row 
  
  Log.Message("Start to creat strategy " & Strategy.StrategyTable)
  Call GetStrategylegs(Strategy.StrategyTable, Strategylegs)

  
  
  If Strategy.StrategyMethod = "Strategy" then 
    Strategy.StrategyID = StrategyMaker(Strategy, Strategylegs)
    
  ElseIf Strategy.StrategyMethod = "TMC Maker" Then 
    Strategy.StrategyID = TMCStrategyMaker(Strategy, Strategylegs)   
  End If
  
  If Strategy.StrategyID <> "" Then
    Call WriteTable(Strategy)
  End If
  
End Function



'Collect the information of strategy legs in the strategy table  
Sub GetStrategylegs(StrategyName,Strategylegs)

  Dim StrategyLegsTable, NumberOfLegs, CallPut
  StrategyLegsTable = "Strategies_" & StrategyName  
  Set StrategyLegsTable = ExcelDriver.GetDataTable(Vars.xlTestScriptsFile, "Strategies", StrategyLegsTable)
     
  NumberOfLegs = StrategyLegsTable.ListRows.Count
  
  Redim Strategylegs(NumberOfLegs)
  Dim RowNumber, ProductType, ProductMonth, ProductStrike 
  
  For RowNumber = 2 to NumberOfLegs + 1
        ProductType = StrategyLegsTable.Range.Cells(RowNumber, 2).Value
        ProductMonth = StrategyLegsTable.Range.Cells(RowNumber, 3).Value    
        If ProductType = "Future" Then 
            Strategylegs(RowNumber-1) = "SIM.F." &TestConfig.FutureProduct &"." &ProductMonth           
        ElseIf ProductType = "Option" Then 
            ProductStrike = StrategyLegsTable.Range.Cells(RowNumber, 4).Value
            Select Case StrategyLegsTable.Range.Cells(RowNumber, 5).Value
              Case "Call" 
                CallPut = ".C.0"             
              Case "Put" 
                CallPut = ".P.0"                 
             End Select 
            Strategylegs(RowNumber-1) = "SIM.O." &TestConfig.OptionProduct &"." &ProductMonth &"." & ProductStrike & CallPut
        Else 
          Log.Error("The information for strategy legs is not correct.")  
          Exit Sub
        End If 
  Next
  
End Sub  


' Collect the stragies already in the MV (before creating new one) to compare later
Function GetOldStrategies(OldStrategies,PreviousNumber)
  
  Dim i, Row,Col1, Col2, NumberRows
  
  NumberRows = GridControl.GetRowCount(OptionViewGrid.Handle) 
  
  i=1
  Col1 = GridControl.GetCellColumn(OptionViewGrid.Handle,"Product Type", 1)
  If Col1 = "-1" Then
    Log.Warning("Could not find column 'Product Type'.")
  End If
  Col2 = GridControl.GetCellColumn(OptionViewGrid.Handle,"ProductID", 1) 
  If Col2 = "-1" Then
    log.Warning("Could not find column 'ProductID'.")
  End If 
  
  For Row = 1 to NumberRows + 1
    If  GridControl.GetCellText(OptionViewGrid.Handle,Row,Col1)= "STRATEGY" Then
        OldStrategies(i) = GridControl.GetCellText(OptionViewGrid.Handle,Row,Col2)
        i = i+1
    End If
   Next 
   
   PreviousNumber = i
   
End Function      
  
 ' Create strategy with StrategyMaker
Function StrategyMaker(Strategy, Strategylegs)  

  Dim ColumnInstance, NumberRows, PreviousNumber  
  Dim Found, ProductName, NumberOfStrategies
  Dim i, Row, Col, NewStrategy, CoordArray, StrategyShort
     
  NumberRows = GridControl.GetRowCount(OptionViewGrid.Handle)   
  
  ReDim OldStrategies(NumberRows)
    
  Call GetOldStrategies(OldStrategies,PreviousNumber)

  For i = 1 to UBound(Strategylegs) 
  
    If InStr(Strategylegs(i), ".F.") or InStr(Strategylegs(i), ".C.") Then
      ColumnInstance = 1
    ElseIf  InStr(Strategylegs(i), ".P.") Then
      ColumnInstance = 2
    Else 
      Log.Error ("Wrong leg for the strategy - " & Strategylegs(i)) 
      Exit Function
    End If 
              
    Row = GridControl.GetCellRow(OptionViewGrid.Handle, "ProductID", Strategylegs(i), ColumnInstance) 
    If Row = -1 Then
      Log.Error("The strategy legs " & Strategylegs(i) & " are not found in the MarketView.")
      If dlgStrategyMaker.Visible Then 
        dlgStrategyMaker.Close  
      End If
      Exit Function
    End If
    Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Theo", ColumnInstance)
    Call MakeCellVisible(OptionViewGrid, Row, Col)
    If i > 1 Then
      CoordArray = Split(GridControl.GetCellCoordinates(OptionViewGrid.Handle, Row, Col),"?")
      Call dlgStrategyMaker.Position(CoordArray(0)+60, dlgStrategyMaker.ScreenTop, dlgStrategyMaker.Width, dlgStrategyMaker.Height)
    End If  
    Call TestUtilities.ClickGrid(OptionViewGrid, Row, Col, "Left")
  Next
    
  Found = False
  Call dlgStrategyMaker.Position(360, dlgStrategyMaker.ScreenTop, dlgStrategyMaker.Width, dlgStrategyMaker.Height)
  NumberOfStrategies = GridControl.GetRowCount(StrategyGrid.Handle)
  StrategyShort = Mid(Strategy.StrategyTable, 1, 3)
  If Mid(StrategyShort, 1, 2) = "TS" Then
    StrategyShort = "TS"
  End If         
         
  For i = 1 to NumberOfStrategies
    ProductName = GridControl.GetCellText(StrategyGrid.Handle, i, 1)  
    If StrategyShort = Left(ProductName, Len(StrategyShort))Then  
      Found = True
      Strategy.StrategyName = ProductName
'      If StrategyType = "Internal" Then
'        dlgStrategyMaker.RadioInternal.ClickButton
'        Delay(100)
'      End If
      Call ClickGrid(StrategyGrid, i, 1, "Left")
      Delay(200)
      dlgStrategyMaker.btnCreate.Click 
      Delay(100)
      Exit for 
    End If
  Next
  
  If Found = False Then
   Log.Error("The required strategy can not be found in the strategy maker with the provided legs." ) 
  End If
  
  Delay(500) 
  
  Dim dlgCreateStrategyError
  Set dlgCreateStrategyError = Aliases.MarketView.dlgCreateStrategyError
  
  'MS - The wait is too long
  'If dlgCreateStrategyError.Exists Then
  '  Log.Error("The request to the server was not replied to within a reasonable time period.")
  '  dlgCreateStrategyError.btnOK.ClickButton
  '  Found = False
  'End If
  
  If Aliases.MarketView.WaitAliasChild("dlgCreateStrategyError",200).Exists Then
    Log.Error("The request to the server was not replied to within a reasonable time period.")
    dlgCreateStrategyError.btnOK.ClickButton
    Found = False
  End If
  
  dlgStrategyMaker.Activate
  dlgStrategyMaker.Close
  
  If Found Then
'    StrategyType = StrategyType & " strategy "
'    Log.Message ("The strategy " & Strategy.StrategyName & " has been created succssfully.")
    StrategyMaker = GetNewStrategy(Strategy, OldStrategies,PreviousNumber)
  End If
  
End Function  


' Create strategy with TMC Maker 
Function TMCStrategyMaker(Strategy, Strategylegs)  

  Dim ColumnInstance, NumberRows, PreviousNumber  
  Dim Found, ProductName, NumberOfStrategies
  Dim i, Row, Col, NewStrategy, CoordArray, StrategyShort
     
  NumberRows = GridControl.GetRowCount(OptionViewGrid.Handle)   
  
  ReDim OldStrategies(NumberRows)
    
  Call GetOldStrategies(OldStrategies,PreviousNumber)

  For i = 1 to UBound(Strategylegs) 
  
    If InStr(Strategylegs(i), ".F.") or InStr(Strategylegs(i), ".C.") Then
      ColumnInstance = 1
    ElseIf  InStr(Strategylegs(i), ".P.") Then
      ColumnInstance = 2
    Else 
      Log.Error ("Wrong leg for the strategy - " & Strategylegs(i)) 
      Exit Function
    End If 
              
    Row = GridControl.GetCellRow(OptionViewGrid.Handle, "ProductID", Strategylegs(i), ColumnInstance) 
    If Row = -1 Then
      Log.Error("At least one of the strategy legs are not found in the MarketView.")
      If dlgTmcMaker.Visible Then
        dlgTmcMaker.Close 
      End If      
      Exit Function
    End If
    Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Theo", ColumnInstance)
    Call MakeCellVisible(OptionViewGrid, Row, Col)
    If i > 1 Then
      CoordArray = Split(GridControl.GetCellCoordinates(OptionViewGrid.Handle, Row, Col),"?")
      Call dlgTmcMaker.Position(CoordArray(0)+60, dlgTmcMaker.ScreenTop, dlgTmcMaker.Width, dlgTmcMaker.Height)
    End If  
    Call TestUtilities.ClickGrid(OptionViewGrid, Row, Col, "Right")
  Next
    
'  If StrategyType = "Internal" Then
'     dlgTmcMaker.RadioInternal.ClickButton
'     Delay(100)
'  End If
  
  Call dlgTmcMaker.Position(360, dlgTmcMaker.ScreenTop, dlgTmcMaker.Width, dlgTmcMaker.Height)
  ProductName = dlgTmcMaker.StrategyDescription.wText
  Strategy.StrategyName = ProductName
  dlgTmcMaker.btnCreate.Click 
  Delay(150)
  
  Dim dlgCreateStrategyFailed
  Set dlgCreateStrategyFailed = Aliases.MarketView.dlgCreateStrategyFailed
   If dlgCreateStrategyFailed.Exists Then
    Log.Error(Strategy.StrategyTable & "  - The request to the server was not replied to within a reasonable time period.")
    dlgCreateStrategyFailed.btnOK.ClickButton
    Found = False
  End If
  
  dlgTmcMaker.Activate
  dlgTmcMaker.Close
  
  Delay(200)
  
'  StrategyType = StrategyType & " TMC strategy "
  TMCStrategyMaker = GetNewStrategy(Strategy,OldStrategies,PreviousNumber)
   
'  Row = GridControl.GetCellRow(OptionViewGrid.Handle, "Product Name", ProductName, 1) 
'    If Row <> -1 Then
'      Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"ProductID", 1)
'      TMCStrategyMaker = GridControl.GetCellText(OptionViewGrid.Handle,Row,Col)
'      Log.Checkpoint(StrategyType & " TMC strategy " & TMCStrategyMaker &" has been created successfully.")  
'    Else
'      log.Error("Strategy " & StrategyName &" can not be created.")
'    End If
  
End Function


'Get the strategy ID if any strategy has been created
Function GetNewStrategy(Strategy,OldStrategies,PreviousNumber)     
   
   Dim i, NumberRows, Row, Col, NewStrategy, ProductName
   
   NumberRows = GridControl.GetRowCount(OptionViewGrid.Handle) 
   Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Product Type", 1) 
   i = 1
   For Row = 1 to NumberRows + 3
      If  GridControl.GetCellText(OptionViewGrid.Handle,Row,Col)= "STRATEGY" Then
        i = i+1
      End If
    Next 
   
    If i = PreviousNumber Then
      Log.Message("No new stragety has been created.")
      Row = GridControl.GetCellRow (OptionViewGrid.Handle,"Product Name", Strategy.StrategyName, 1)  
        If Row > 1 Then
          Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"ProductID", 1)
          Strategy.StrategyID = GridControl.GetCellText(OptionViewGrid.Handle,Row,Col)
          Call WriteTable(Strategy) 
        End If   
    ElseIf i > PreviousNumber + 1 Then
       Log.Message("More than two stragetes have been created.") 
    Else
      NewStrategy = FindNewStrategy (OldStrategies, PreviousNumber)
      If NewStrategy <> "" Then 
        GetNewStrategy = NewStrategy
        Row = GridControl.GetCellRow(OptionViewGrid.Handle,"ProductID", NewStrategy, 1)
        Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Product Name", 1)
        ProductName = GridControl.GetCellText(OptionViewGrid.Handle, Row, Col)
        Log.Checkpoint(ProductName &" has been created successfully.")
      Else log.error("The strategy " & Strategy.StrategyTable & " can not be found.")
      End If
    End If  
  
End Function


'Find the strategy just created
Function FindNewStrategy (OldStrategies, PreviousNumber)
  
  Dim Row, Col, Col2, Col3, i, StrategyID, StrategyName, Found, NewStrategy, NewStrategyName
  Col = GridControl.GetCellColumn(OptionViewGrid.Handle,"Product Type", 1)
  Col2 = GridControl.GetCellColumn(OptionViewGrid.Handle,"ProductID", 1) 

  FindNewStrategy = ""
  
  For Row = 1 to GridControl.GetRowCount(OptionViewGrid.Handle) + 1
    If  GridControl.GetCellText(OptionViewGrid.Handle,Row,Col)= "STRATEGY" Then
       StrategyID = GridControl.GetCellText(OptionViewGrid.Handle,Row,Col2)
       Found = False
       For i = 1 to PreviousNumber
         If StrategyID = OldStrategies(i)Then
          Found = True
          Exit For
        End If
      Next 
      If Not Found Then 
        FindNewStrategy = StrategyID
        Exit For
      End If   
    End If
   Next 

End Function


' Write the StrategyID back to the Orders table
Function WriteTable(Strategy)  
  
  Dim  ExcelObject, ExcelWorkbook, ExcelWorksheet, Table, i

  Set ExcelObject = Sys.OleObject("Excel.Application")
  Set ExcelWorkbook = ExcelObject.Workbooks.Open(Vars.xlTestScriptsFile)
  ExcelWorkbook.Activate
  Set ExcelWorksheet = ExcelObject.Worksheets("Strategies")
  
  For i = 1 to ExcelWorksheet.ListObjects.Count
    If ExcelWorksheet.ListObjects(i).Name = "Strategies" Then
          Set Table = ExcelWorksheet.ListObjects(i)     
      Exit For   
    End If 
  Next

  For i = 1 to Table.ListRows.Count + 1
    If Table.Range.Cells(i, 1).Value = Strategy.StrategyNumber Then
      Table.Range.Cells(i, 4).Value = Strategy.StrategyName     
      Table.Range.Cells(i, 5).Value = Strategy.StrategyID
      ExcelWorkbook.Save 
      Exit For
    End If   
  Next
   
End Function

Sub CopyStrategy (Strategies,TestWorksheetName, OrderTblName)

  Dim StrategyNumber, ExcelObject, ExcelWorkbook, ExcelWorksheet, OrderTable, Row, i
  
      Set ExcelObject = Sys.OleObject("Excel.Application") 
      Set ExcelWorkbook = ExcelObject.Workbooks.Open(Vars.xlTestScriptsFile)
      ExcelWorkbook.Activate
      Set ExcelWorksheet = ExcelObject.Worksheets(TestWorksheetName)
      
        For i = 1 to ExcelWorksheet.ListObjects.Count
          If ExcelWorksheet.ListObjects(i).Name = OrderTblName Then
            Set OrderTable = ExcelWorksheet.ListObjects(i)     
            Exit For   
          End If 
        Next
           
      For Row = 2 to OrderTable.ListRows.Count + 1 
        If OrderTable.Range.Cells(Row,2).Value = "Strategy" or OrderTable.Range.Cells(Row,2).Value = "TMCStrategy" Then 
           StrategyNumber = OrderTable.Range.Cells(Row,7).Value
           For  i = 1 to UBound(Strategies)
              If StrategyNumber = Strategies(i).StrategyNumber Then
                 OrderTable.Range.Cells(Row,8).Value = Strategies(i).StrategyName
                 OrderTable.Range.Cells(Row,9).Value = Strategies(i).StrategyID
                 Exit For
              End If
            Next    
        End If
      Next
                 
  ExcelWorkbook.Save
   
End Sub 


Sub ResetStrategyTable

Dim  ExcelObject, ExcelWorkbook, ExcelWorksheet, Table, i, attr

  Set attr = Log.CreateNewAttributes 
  Set ExcelObject = Sys.OleObject("Excel.Application")
  Set ExcelWorkbook = ExcelObject.Workbooks.Open(Vars.xlTestScriptsFile)
  ExcelWorkbook.Activate
  Set ExcelWorksheet = ExcelObject.Worksheets("Strategies")
  
  attr.Bold = True
  attr.Italic = False 
      
  Call Log.Checkpoint("Reset the strategies table for test suite.", , , attr)
  
  For i = 1 to ExcelWorksheet.ListObjects.Count
    If ExcelWorksheet.ListObjects(i).Name = "Strategies" Then
      Set Table = ExcelWorksheet.ListObjects(i)     
      Exit For   
    End If 
  Next
  
  For i = 2 to Table.ListRows.Count + 1
    Table.Range.Cells(i, 4).Value = ""
    Table.Range.Cells(i, 5).Value = "N/A"
  Next
  
  ExcelWorkbook.Save 
   
End Sub
