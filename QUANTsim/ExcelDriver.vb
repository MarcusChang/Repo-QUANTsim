'-------------------------------------------------------------------------------------------------------------------------
'UNIT Excel Driver
'Description: Functions relating to reading and accessing Excel will be placed here 
'-------------------------------------------------------------------------------------------------------------------------

Dim TestConfig
Set TestConfig = ProjectSuite.Variables

'-------------------------------------------------------------------------------------------------------------------------
'Name: OpenWorkSheet()
'Arguments: 
'Description:
'Opens the specified Excel Worksheet for processing
'-------------------------------------------------------------------------------------------------------------------------
Public Function OpenWorkSheet(xlFilename, xlWorksheetName)

Dim ExcelObject, ExcelWorkbook, ExcelWorksheet 
  
  'Declare the ExcelObject
  Set ExcelObject = Sys.OleObject("Excel.Application")
  ExcelObject.Visible = True
  Delay(100)
  
  'Open the Workbook
  Set ExcelWorkbook = ExcelObject.Workbooks.Open(xlFilename)

  'Open the Worksheet  
  Set ExcelWorksheet = ExcelObject.Worksheets(xlWorksheetName)
  
  Set OpenWorkSheet = ExcelWorksheet

End Function

'-------------------------------------------------------------------------------------------------------------------------
'Name: GetDataTable()
'Arguments: 
'Description:
'Gets the Table from Excel for processing
'-------------------------------------------------------------------------------------------------------------------------
Public Function GetDataTable(xlFilename, xlWorksheetName, xlTableName)

  Dim xlWorksheet 

  Log.Message("xlWorksheetName = " & xlWorksheetName)
  
  Set xlWorksheet = OpenWorkSheet(xlFilename, xlWorksheetName)
     
  'Search through the Worksheet for the specified Table
  Dim i, TableFound
  TableFound = False
  
  For i = 1 to xlWorksheet.ListObjects.Count
    If xlWorksheet.ListObjects(i).Name = xlTableName Then
      TableFound = True
      Set GetDataTable = xlWorksheet.ListObjects(i)     
      Exit For   
    End If 
  Next
  
  'If the table is not found
  If TableFound = False Then
    Log.Error("Table: " & xlTableName & " was not found in Worksheet: " & xlWorkSheetName)
    Exit Function
  End If    
      
End Function


'-------------------------------------------------------------------------------------------------------------------------
'Function Graveyard
'-------------------------------------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------------------------------------------
'26/08/11 Incomplete for now, trying to attempt to create a Class for the ATTD specs
'-------------------------------------------------------------------------------------------------------------------------

Sub GetATTDSpecs(AttdSpecObject)

  Dim ExcelObject, ExcelFileName, ExcelWorkSheet, ExcelWorkbook
  
  ExcelFilename = "C:\Marcus_Chang\Project_TestConfig.xlsx"
  ExcelWorksheet = "TOMSpec"
  
  Set ExcelObject = Sys.OleObject("Excel.Application") 
  ExcelObject.Visible = True
  Delay(1500)
  
  Set ExcelWorkbook = ExcelObject.Workbooks.Open(ExcelFileName)
  
  Dim worksheet, returnValue, range
  
  Set worksheet = ExcelObject.Worksheets(ExcelWorkSheet)
  worksheet.Activate
  
  'I already know theres one table, so just doing it dodgy for now
  'In future do a search through the objects and find table name = ProjectVariables
  Dim VariablesTable
  Set VariablesTable = worksheet.ListObjects(1)
  
  Dim RowNumber, VariableName, VariableValue, i
    
  For RowNumber = 2 to VariablesTable.ListRows.Count + 1
      VariableName = VariablesTable.Range.Cells(RowNumber, 1).Value
      VariableValue = VariablesTable.Range.Cells(RowNumber, 2).Value
      
      Call AttdSpecObject.SetVariable(VariableName,VariableValue)
      
      
      'For i = 0 to (PropNum - 1)
       ' If AttdSpecObject.Properties(i).Name = VariableName Then
        '    AttdSpecObject.Properties(i).Value = VariableValue
         '   Log.Message("Match Found")
        'End if
      'Next      
  Next
  
  
End Sub

Sub GetATTDSpecs_(AttdSpecObject)

  AttdSpecObject.ATT_Spec_Name = "Testing"

End Sub


Class ATTDSpec

  Public ATT_Spec_Name
  Public Machine
  Public Edge_Spec
  Public Pricing_Spec
  Public Connection_Type
  Public TOM_Type
  Public USerGroup
  Public User
  Public Cube_Depth
  Public Underlying_Ref_Type
  Public Price_Driver
  Public Timed_Re_Calc
  Public Spread_Tables
  Public Auto_Cancel
  Public Auto_Hedge
  Public Day
  Public Contingent
  Public Hedge_Name
  Public Hedge_Acct
  Public Hedge_Exchange_Acct
  Public Hedge_Exchange
  Public Hedge_User
  Public Target_Name
  Public Target_Acct
  Public Target_Exchange_Acct
  Public Target_Exchange
  Public Target_User
  Public Note_1
  Public Note_2
  Public Note_3
  Public Note_4
  Public Note_5
  Public Note_6
  
  'Public Property Let
   'End Property
  
   Public Function SetVariable(VariableName,VariableValue)

    If VarType(Eval(VariableName))= 0 Then
        Log.Message("Variable does not exist " & VariableName )
    Else
        Log.Message("Variable does exist " & VariableName)            
    End If    
      
   End Function

End Class

Function GetAttdSpecObject

  Set GetAttdSpecObject = New ATTDSpec

End Function

'-------------------------------------------------------------------------------------------------------------------------
'Name: DDT_Beta(FileName)
'Arguments: Filename (location of xl file)
'Description: Just practicing on DDT usage 
'-------------------------------------------------------------------------------------------------------------------------
Sub DDT_Beta '(FileName,Sheet)

  Dim Filename, Sheet
  FileName = "C:\Marcus_Chang\Project_TestConfig.xlsx"
  Sheet = "Exchange"

  'Create a driver
  Dim Driver, i  
  Set Driver = DDT.ExcelDriver(Filename, Sheet, True)
  
  'Debug - Loop through each column and print it out to the log
  For i = 0 To Driver.ColumnCount - 1
      Log.Message(Driver.ColumnName(i))
      Log.Message(Driver.Value(i))      
  Next

End Sub


'-------------------------------------------------------------------------------------------------------------------------
'Name: SetProjectVariables()
'Arguments: Filename (location of xl file)
'Description: 
'-------------------------------------------------------------------------------------------------------------------------
Sub SetProjectVariables_Beta

  'Declare the Excel Object
  Dim ExcelObject, ExcelFileName, ExcelWorkSheet, ExcelWorkbook
  
  ExcelFilename = "C:\Marcus_Chang\Project_Variables.xlsx"
  ExcelWorkSheet = "Exchange"
  
  
  Set ExcelObject = Sys.OleObject("Excel.Application")
  ExcelObject.Visible = True  
  Delay(2000)
  
  'Open the workbook and assign it to an object
  Set ExcelWorkbook = ExcelObject.Workbooks.Open(ExcelFilename)

  Dim worksheet, returnValue, range
  
  'Open the worksheet and assign it to an object  
  Set worksheet = ExcelObject.Worksheets(ExcelWorkSheet)
  worksheet.Activate
  
  'Search through worksheet and find specified table
  Dim VariablesTable, a, TableName
  TableName = "SGX_"

  For a = 1 to worksheet.ListObjects.Count
     
    If worksheet.ListObjects(a).Name = TableName Then
      Log.Message("Table name found")
      Set VariablesTable = worksheet.ListObjects(a)
      Exit For
    Else
      Log.Message("Cannot find specified Table")
    End If
  Next
  'Log.Message("VariableTable Name: " & VariablesTable)
  'Log.Message("Number of Rows " & VariablesTable.ListRows.Count)
  'Log.Message("Number of Columns " & VariablesTable.ListColumns.Count)
  'Log.Message("Range Count " & VariablesTable.Range.Count)
  
  Dim RowNumber, ColNumber, VariableName, VariableValue
    For RowNumber = 2 to VariablesTable.ListRows.Count + 1
      'Check if variable exists, if it does, then assign it to a value, if not add it to the collection
      VariableName = VariablesTable.Range.Cells(RowNumber, 1).Value
      VariableValue = VariablesTable.Range.Cells(RowNumber, 2).Value
      
      If TestConfig.VariableExists(VariableName) Then 
        'Variable exists, figure out what type it is defined as and convert it accordingly
        Log.Message("Project Variable already exists: " & VariableName)
        
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
  TestConfig.UnderlyingProductName = TestConfig.UnderlyingProduct&" "&TestConfig.UnderlyingShortMonth
  TestConfig.UnderlyingProductID = "SIM.F."&TestConfig.UnderlyingProduct&"."&TestConfig.UnderlyingLongMonth

  ExcelObject.Quit
  
End Sub
