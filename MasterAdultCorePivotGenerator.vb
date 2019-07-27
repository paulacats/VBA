Sub AdultCorePivots()
	'--------------------------------ABOUT THIS SCRIPT--------------------------------------
	'This script generates all of the pivot tables for all of the Adult Core Measures.
	'--------------------------------INSTRUCTIONS FOR USE--------------------------------------
	'To use this file: 
	'1. Load the file as a module into an Excel Workbook. 
	'2. Save the file as macro-enabled (xlsm). 
	'3. Copy the file to the folder location where the Adult Core Original files exist.
	'4. From the Developer window, run the macro AdultCorePivots.
	'Note: Steps 1 and 2 only need to be repeated should you update the code that follows.
	'-----------------------------------------------------------------------------------------
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	FileExt = "xls*"
	''This loops through every file in folder to pull over the tab we want
	directory = ThisWorkbook.Path & "\"
	filesname = Dir(directory & "*" & FileExt)
	Do While filesname <> ""
		''Specify name of file in filename
		If InStr(1, filesname, "Prev 1") > 0 And InStr(1, UCase(filesname), "ANNUAL") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createAWVPivotTables
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 2") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createBMIPivotTables
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 3") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createMammographyPivotTables
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 4") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createColonCancerPivotTables
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 5") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createDepressionPivotTables
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 6") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createFallRiskPivotTable
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 7") > 0 And InStr(1, filesname, "Influenza") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createInfluenzaPivotTablePivotTable
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 8") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createPneumoPivotTable
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 9") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createTobaccoPivotTables
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "Prev 10") > 0 Then
			Set wbOpen = Workbooks.Open(directory & filesname)
			With wbOpen
				.Activate
				Call createCervicalCancerPivotTables
				.Close savechanges:=True
			End With
		ElseIf InStr(1, filesname, "DM Combined") > 0 Then
				Set wbOpen = Workbooks.Open(directory & filesname)
				With wbOpen
					.Activate
					Call createDMCombinedPivotTables
					.Close savechanges:=True
				End With
		ElseIf InStr(1, filesname, "Cardiac 18") > 0 Then
				Set wbOpen = Workbooks.Open(directory & filesname)
				With wbOpen
					.Activate
					Call createCardiacPivotTables
					.Close savechanges:=True
				End With
		ElseIf InStr(1, filesname, "SDOH") > 0 Then
				Set wbOpen = Workbooks.Open(directory & filesname)
				With wbOpen
					.Activate
					Call createSDOHPivotTables
					.Close savechanges:=True
				End With
		End If
		filesname = Dir
	Loop
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub
Sub createAWVPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Page1").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = core & "PivotTables" & .Worksheets.Count
		Set myDestinationWorksheet2 = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet2.Name = core & "AWVPivot" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("A35").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("G9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet2.Range("A9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet2.Range("A35").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange6 = myDestinationWorksheet2.Range("A66").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 18+ at Year End")
			.Orientation = xlPageField
			.Position = 3
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With myDestinationWorksheet.Range("A1")
		.Value = "Coastal Core"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	With myDestinationWorksheet.Range("C7")
		.Value = "1 for unified; (all) for DNS"
	End With
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Pedi Numerator'/ 'Denominator'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("Bald Hill Pedi").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "Bald Hill Pedi" Or Pi.Value = "Narragansett Bay Pedi" Or Pi.Value = "Toll Gate Pedi" Or Pi.Value = "Waterman Pedi" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Narragansett Family Medicine" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Pedi Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 12-21 at Year End")
			.Orientation = xlPageField
			.Position = 3
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With myDestinationWorksheet.Range("A27")
		.Value = "Pedi Core"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	With myDestinationWorksheet.Range("C33")
		.Value = "1 for unified; (all) for DNS"
	End With
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated fields
		.CalculatedFields.Add "UHC Elig Numerator Rate", "= 'UHC Elig Numerator'/ 'Denominator'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("UHC Elig Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("UHC Elig Numerator Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("UHC MA Eligibility Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC Elig Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With myDestinationWorksheet.Range("G1")
		.Value = "UHC MA-Eligibility"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'clean up
	With myDestinationWorksheet.Rows("10")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("31")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	'tally pivots
	'create 4th pivot table
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet2.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	'add, organize and format Pivot Table fields
	With myPivotTable4
		'add calculated field
		.CalculatedFields.Add "AWV Rate", "= 'AWV Numerator'/'Age 66+ at Year End'", True 
		.CalculatedFields.Add "Coastal Core Prev 1 Rate", "= 'Numerator'/'Denominator'", True 
		.CalculatedFields.Add "AWV Exceptions", "= '66+ Medicare MA or Dual'-'AWV Numerator'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				End If
			Next i
		End With
		With .PivotFields("Empaneled Flag")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("AWV Numerator")
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("AWV Exceptions")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With 
		With .PivotFields("AWV Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With 
		With .PivotFields("Coastal Core Prev 1 Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Due Next 3 Months and Not Scheduled FINAL")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With 
		'Filters
		With .PivotFields("Empaneled Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("66+ Medicare MA or Dual")
			.Orientation = xlPageField
			.Position = 3
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With myDestinationWorksheet2.Range("A1")
		.Value = "Coastal Core - Annual Wellness Visit 66+"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	With myDestinationWorksheet2.Range("A2")
		.Value = "Measurement Year between  Jan 1, 2019 and Dec 31, 2019"
		.Font.Size = 11
	End With
	With myDestinationWorksheet2.Range("I28")
		.Value = "VALIDATION - COUNT OF PATIENTS WITH AWV EXCEPTION STATUS  IN DNS REPORT"
		.Interior.ColorIndex = 6
	End With
	'create 5th pivot
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet2.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	'add, organize and format Pivot Table fields
	With myPivotTable5
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				End If
			Next i
		End With
		With .PivotFields("AWV Exception Status")
			.Orientation = xlColumnField 
		End With
		With .PivotFields("Patient Account Number")
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlCount
		End With
		 
		With .PivotFields("AWV Exception Status")
			.Orientation = xlDataField
			.Function = xlCount
			.Calculation = xlPercentOfRow
			.NumberFormat = "0.00%"
		End With 
		'Filters
		With .PivotFields("Empaneled Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("66+ Medicare MA or Dual")
			.Orientation = xlPageField
			.Position = 3
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'create 6th pivot
	Set myPivotTable6 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet2.Name & "!" & myDestinationRange6, TableName:="PivotTableNewSheet6")
	'add, organize and format Pivot Table fields
	With myPivotTable6
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Numerator'/'Denominator'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				End If
			Next i
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With 
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Due Next 3 Months and Not Scheduled FINAL")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With 
		'Filters
		With .PivotFields("Empaneled Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 18+ at Year End")
			.Orientation = xlPageField
			.Position = 3
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 4
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With myDestinationWorksheet2.Range("A58")
		.Value = "ALL EMPANELED - Due Next 3 Months and Not Scheduled"
		.Font.FontStyle = "Bold"
		.Font.Size = 12
	End With
	With myDestinationWorksheet2.Range("F88")
		.Value = "VALIDATION - TOTAL PATIENT COUNT IN DNS Report"
		.Interior.ColorIndex = 6
	End With
		'clean up
	With myDestinationWorksheet2.Rows("10")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet2.Rows("36")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet2.Rows("67")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createBMIPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Deduped Page").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = core & "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A11").Address(ReferenceStyle:=xlR1C1)
	'myDestinationRange2 = myDestinationWorksheet.Range("G11").Address(ReferenceStyle:=xlR1C1)
	'myDestinationRange3 = myDestinationWorksheet.Range("G39").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet.Range("G11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet.Range("G39").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange6 = myDestinationWorksheet.Range("M11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange7 = myDestinationWorksheet.Range("M39").Address(ReferenceStyle:=xlR1C1)
	'myDestinationRange8 = myDestinationWorksheet.Range("S11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange9 = myDestinationWorksheet.Range("U11").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Core Numerator_W/ Excl'/ 'Core Denominator_ W/ Excl'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Core Numerator_W/ Excl")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Core Denominator_ W/ Excl")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		'Filters
		With .PivotFields("Age 18-74_HEDIS")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Last Elig Enc_Visit_Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Core Denominator_ W/ Excl")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With Range("A1")
		.Value = "Core RATE by Visit"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create second pivot table
	' Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	' 'add, organize and format Pivot Table fields
	' With myPivotTable2
		' 'add calculated fields
		' .CalculatedFields.Add "Rate", "= 'HEDIS_Num'/ 'HEDIS_Denom_Excl Applied'", True 
		' With .PivotFields("Demographics PCP Facility")
			' .Orientation = xlRowField
			' .Position = 1
			' ' .PivotItems("East Greenwich").Visible = True
			' ' For Each Pi In .PivotItems
				' ' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' ' Pi.Visible = True
				' ' Else
					' ' Pi.Visible = False
				' ' End If
			' ' Next Pi
			' .AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		' End With
		' With .PivotFields("HEDIS_Num")
			' .Orientation = xlDataField
			' .Position = 1
			' .Function = xlSum
		' End With
		' With .PivotFields("HEDIS_Denom_Excl Applied")
			' .Orientation = xlDataField
			' .Position = 2
			' .Function = xlSum
		' End With
		' With .PivotFields("Rate")
			' .Orientation = xlDataField
			' .Position = 3
			' .Function = xlSum
			' .NumberFormat = "0.00%"
		' End With
		' With .PivotFields("Scheduled This Year")
			' .Orientation = xlDataField
			' '.Position = 2
			' .Function = xlSum
		' End With
		' 'Filters
		' With .PivotFields("BCBS MA Denom BMI")
			' .Orientation = xlPageField
			' .Position = 1
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("Deceased")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' '.CurrentPage = "No"
			' On Error GoTo 0
		' End With
		' With .PivotFields("Status")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' '.CurrentPage = "Active"
			' On Error GoTo 0
		' End With
		' With .PivotFields("Age 18-74_HEDIS")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("HEDIS_Num")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' '.CurrentPage = "1"
			' On Error GoTo 0
		' End With
	' End With
	' 'add title
	' With Range("G1")
		' .Value = "BCBSRI MA"
		' .Font.FontStyle = "Bold"
		' .Font.Size = 14
	' End With
	' 'create third pivot table
	' Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	' 'add, organize and format Pivot Table fields
	' With myPivotTable3
		' 'add calculated fields
		' .CalculatedFields.Add "Rate", "= 'HEDIS_Num'/ 'HEDIS_Denom_Excl Applied'", True 
		' With .PivotFields("Demographics PCP Facility")
			' .Orientation = xlRowField
			' .Position = 1
			' ' .PivotItems("East Greenwich").Visible = True
			' ' For Each Pi In .PivotItems
				' ' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' ' Pi.Visible = True
				' ' Else
					' ' Pi.Visible = False
				' ' End If
			' ' Next Pi
			' .AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		' End With
		' With .PivotFields("HEDIS_Num")
			' .Orientation = xlDataField
			' .Position = 1
			' .Function = xlSum
		' End With
		' With .PivotFields("HEDIS_Denom_Excl Applied")
			' .Orientation = xlDataField
			' .Position = 2
			' .Function = xlSum
		' End With
		' With .PivotFields("Rate")
			' .Orientation = xlDataField
			' .Position = 3
			' .Function = xlSum
			' .NumberFormat = "0.00%"
		' End With
		' With .PivotFields("Scheduled This Year")
			' .Orientation = xlDataField
			' '.Position = 2
			' .Function = xlSum
		' End With
		' 'Filters
		' With .PivotFields("BCBS Comm Denom BMI")
			' .Orientation = xlPageField
			' .Position = 1
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("Deceased")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' '.CurrentPage = "No"
			' On Error GoTo 0
		' End With
		' With .PivotFields("Status")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' '.CurrentPage = "Active"
			' On Error GoTo 0
		' End With
		' With .PivotFields("Age 18-74_HEDIS")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("HEDIS_Num")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' '.CurrentPage = "1"
			' On Error GoTo 0
		' End With
	' End With
	' 'add title
	' With Range("G31")
		' .Value = "BCBSRI COMM "
		' .Font.FontStyle = "Bold"
		' .Font.Size = 14
	' End With
	'4th pivot
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	With myPivotTable4
	'add calculated fields
		.CalculatedFields.Add "Rate", "='HEDIS_Num'/ 'HEDIS_Denom_Excl Applied'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("BCBS MA Denom BMI")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("HEDIS_Num_Max Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("BCBS MA Denom BMI")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Age 18-74_HEDIS")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS_Num")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("G1")
		.Value = "BCBSRI MA"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'5th pivot
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	With myPivotTable5
		'add calculated fields
		.CalculatedFields.Add "Rate", "='HEDIS_Num'/ 'HEDIS_Denom_Excl Applied'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			'.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("BCBS Comm Denom BMI")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("HEDIS_Num_Max Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			''.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("BCBS Comm Denom BMI")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Age 18-74_HEDIS")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS_Num")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("G31")
		.Value = "BCBSRI COMM"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'6th pivot
	Set myPivotTable6 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange6, TableName:="PivotTableNewSheet6")
	With myPivotTable6
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'HEDIS_Num'/ 'HEDIS_Denom_Excl Applied'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			'.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("HEDIS_Num")
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("HEDIS_Denom_Excl Applied")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			''.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Tufts Attributed")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS_Denom_Excl Applied")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Last Elig Enc_CPT_Flag")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS_Num")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("M1")
		.Value = "Tufts"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'7th pivot
	Set myPivotTable7 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange7, TableName:="PivotTableNewSheet7")
	With myPivotTable7
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'HEDIS_Num'/ 'HEDIS_Denom_Excl Applied'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			'.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("HEDIS_Num")
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("HEDIS_Denom_Excl Applied")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			''.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("UHC MA Attributed")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Last Elig Enc_CPT_Flag")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS_Denom_Excl Applied")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS_Num")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("M31")
		.Value = "UHC MA"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create eight pivot table
	' Set myPivotTable8 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange8, TableName:="PivotTableNewSheet8")
	' With myPivotTable8
		' 'add calculated fields
		' .CalculatedFields.Add "Rate", "= 'UHC Comm Num'/ 'UHC Comm Denom'", True 
		' With .PivotFields("Demographics PCP Facility")
			' .Orientation = xlRowField
			' '.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			' .AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		' End With
		' With .PivotFields("UHC Comm Num")
			' .Orientation = xlDataField
			' '.Position = 1
			' .Function = xlSum
		' End With
		' With .PivotFields("UHC Comm Denom")
			' .Orientation = xlDataField
			' '.Position = 2
			' .Function = xlSum
		' End With
		' With .PivotFields("Rate")
			' .Orientation = xlDataField
			' '.Position = 3
			' .Function = xlSum
			' .NumberFormat = "0.00%"
		' End With
		' With .PivotFields("Scheduled This Year")
			' .Orientation = xlDataField
			' ''.Position = 2
			' .Function = xlSum
		' End With
		' 'Filters
		' With .PivotFields("UHC Comm PCOR Attributed")
			' .Orientation = xlPageField
			' '.Position = 1
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("Active Non Deceased Flag")
			' .Orientation = xlPageField
			' '.Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("UHC Comm Denom")
			' .Orientation = xlPageField
			' '.Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("UHC Comm Num")
			' .Orientation = xlPageField
			' '.Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' '.CurrentPage = "1"
			' On Error GoTo 0
		' End With
	' End With
	' 'add title
	' With Range("S1")
		' .Value = "UHC Comm PCOR"
		' .Font.FontStyle = "Bold"
		' .Font.Size = 14
	' End With
	'create ninth pivot table
	Set myPivotTable9 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange9, TableName:="PivotTableNewSheet9")
	With myPivotTable9
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'HEDIS_Num'/ 'HEDIS_Denom_Excl Applied'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			'.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("HEDIS_Num")
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("HEDIS_Denom_Excl Applied")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			''.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS_Num")
			.Orientation = xlPageField
			''.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS_Denom_Excl Applied")
			.Orientation = xlPageField
			''.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Last Elig Enc_CPT_Flag")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 18-74_HEDIS")
			.Orientation = xlPageField
			''.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Insurance Coverage Type")
			.Orientation = xlPageField
			''.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Medicaid"
			On Error GoTo 0
		End With
		With .PivotFields("Primary Insurance Product")
			.Orientation = xlPageField
			''.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.ClearAllFilters
				.PivotItems("NHP Medicaid").Visible = True
				For Each Pi In .PivotItems
					If Pi.Value = "NHP Medicaid" Or Pi.Value = "UHC Medicaid" Then
						Pi.Visible = True
					Else
						Pi.Visible = False
					End If
				Next Pi
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("U1")
		.Value = "AE Medicaid Measure"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'clean up
	With myDestinationWorksheet.Rows("12")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("40")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createMammographyPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Detail").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = core & "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("H11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("N11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet.Range("T11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet.Range("Z11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange6 = myDestinationWorksheet.Range("AF11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange7 = myDestinationWorksheet.Range("AL11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange8 = myDestinationWorksheet.Range("AR11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange9 = myDestinationWorksheet.Range("AZ11").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With Range("A1")
		.Value = "Core 3: Mammography Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'add notes
	With Range("A30")
		.Value = "Filters"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("A31")
		.Value = "- Adult & Family Med ONLY"
		'.Font.Size = 14
	End With
	With Range("A32")
		.Value = "- CIC, Specialty, To Be Reviewed, Other - Removed"
		'.Font.Size = 14
	End With
	With Range("A34")
		.Value = "Note:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("A35")
		.Value = "- Use Max Denom"
		'.Font.Size = 14
	End With
	
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("H1")
		.Value = "CMS: Mammography Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'add notes
	With Range("H30")
		.Value = "Filters"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("H31")
		.Value = "- Adult & Family Med ONLY"
		'.Font.Size = 14
	End With
	With Range("H32")
		.Value = "- CIC, Specialty, To Be Reviewed, Other - Removed"
		'.Font.Size = 14
	End With
	With Range("H34")
		.Value = "Note:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("H35")
		.Value = "- Use Max Denom"
		'.Font.Size = 14
	End With
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator_BCBSRI'/ 'BCBS MA Denom Mammo'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator_BCBSRI")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("BCBS MA Denom Mammo")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("BCBS MA Denom Mammo")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator_BCBSRI")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("N1")
		.Value = "BCBSRI-MA: Mammography Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'add notes
	With Range("N30")
		.Value = "Post Processing"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("N31")
		.Value = "- BCBSRI Numerator by Collection Date Only"
		'.Font.Size = 14
	End With
	With Range("N33")
		.Value = "Notes:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("N34")
		.Value = "- Use BCBSRI's Denominators"
		'.Font.Size = 14
	End With
	With Range("N35")
		.Value = "- Inactive, Deceased are excluded"
		'.Font.Size = 14
	End With
	With Range("N36")
		.Value = "- CIC, Specialty, To Be Reviewed, Other are not excluded"
		'.Font.Size = 14
	End With
	
	'4th pivot
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	With myPivotTable4
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator_BCBSRI'/ 'BCBS Comm Denom Mammo'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator_BCBSRI")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("BCBS Comm Denom Mammo")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("BCBS Comm Denom Mammo")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator_BCBSRI")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("T1")
		.Value = "BCBSRI-COMM: Mammography Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'add notes
	With Range("T30")
		.Value = "Post Processing"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("T31")
		.Value = "- BCBSRI Numerator by Collection Date Only"
		'.Font.Size = 14
	End With
	With Range("T33")
		.Value = "Notes:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("T34")
		.Value = "- Use BCBSRI's Denominators"
		'.Font.Size = 14
	End With
	With Range("T35")
		.Value = "- Inactive, Deceased are excluded"
		'.Font.Size = 14
	End With
	With Range("T36")
		.Value = "- CIC, Specialty, To Be Reviewed, Other are not excluded"
		'.Font.Size = 14
	End With
	'5th pivot
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	With myPivotTable5
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator_UHC_COMM PCOR'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator_UHC_COMM PCOR")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("UHC Comm PCOR Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator_UHC_COMM PCOR")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("Z1")
		.Value = "UHC COMM PCOR : Mammography Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'add notes
	With Range("Z30")
		.Value = "Post Processing"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("Z31")
		.Value = "- UHC Numerator has a different measurement period"
		'.Font.Size = 14
	End With
	With Range("Z33")
		.Value = "Filters"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("Z34")
		.Value = "- Active, non-deceased"
		'.Font.Size = 14
	End With
	With Range("Z35")
		.Value = "- Other - Removed"
		'.Font.Size = 14
	End With
	With Range("Z37")
		.Value = "Note:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("Z38")
		.Value = "- Use Max Denom"
		'.Font.Size = 14
	End With
	'6th pivot
	Set myPivotTable6 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange6, TableName:="PivotTableNewSheet6")
	With myPivotTable6
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("UHC MA Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AF1")
		.Value = "UHC MA: Mammography Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'add notes
	With Range("AF30")
		.Value = "Filters"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("AF31")
		.Value = "- Active, non-deceased"
		'.Font.Size = 14
	End With
	With Range("AF32")
		.Value = "- Other - Removed"
		'.Font.Size = 14
	End With
	With Range("AF34")
		.Value = "Note:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("AF35")
		.Value = "- Use Max Denom"
		'.Font.Size = 14
	End With
	'7th pivot
	Set myPivotTable7 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange7, TableName:="PivotTableNewSheet7")
	With myPivotTable7
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Tufts Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AL1")
		.Value = "Tufts: Mammography Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'add notes
	With Range("AL30")
		.Value = "Filters"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("AL31")
		.Value = "- Active, non-deceased"
		'.Font.Size = 14
	End With
	With Range("AL32")
		.Value = "- Other - Removed"
		'.Font.Size = 14
	End With
	With Range("AL34")
		.Value = "Note:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("AL35")
		.Value = "- Use Max Denom"
		'.Font.Size = 14
	End With
	'create eight pivot table
	Set myPivotTable8 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange8, TableName:="PivotTableNewSheet8")
	With myPivotTable8
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Tufts LS Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AR1")
		.Value = "Tufts-Lifespan: Mammography Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'add notes
	With Range("AR30")
		.Value = "Filters"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("AR31")
		.Value = "- Active, non-deceased"
		'.Font.Size = 14
	End With
	With Range("AR32")
		.Value = "- Other - Removed"
		'.Font.Size = 14
	End With
	With Range("AR34")
		.Value = "Note:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("AR35")
		.Value = "- Use Max Denom"
		'.Font.Size = 14
	End With
	'create ninth pivot table
	Set myPivotTable9 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange9, TableName:="PivotTableNewSheet9")
	With myPivotTable9
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 50-74 Yrs in Msmt Yr")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Insurance Coverage Type")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Medicaid"
			On Error GoTo 0
		End With
		With .PivotFields("Primary Insurance Product")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.ClearAllFilters
				.PivotItems("NHP Medicaid").Visible = True
				For Each Pi In .PivotItems
					If Pi.Value = "NHP Medicaid" Or Pi.Value = "UHC Medicaid" Then
						Pi.Visible = True
					Else
						Pi.Visible = False
					End If
				Next Pi
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AZ1")
		.Value = "AE Medicaid Measure"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'clean up
	With myDestinationWorksheet.Rows("12")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createColonCancerPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Detail").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = core & "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("H9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("N9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet.Range("N40").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet.Range("T9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange6 = myDestinationWorksheet.Range("T40").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange7 = myDestinationWorksheet.Range("Z9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange8 = myDestinationWorksheet.Range("Z40").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographic PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With Range("A1")
		.Value = "Core 4: Colon Cancer Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'UHC Colo Screening (Comm/PCOR)'/ 'UHC Comm PCOR Attributed'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographic PCP Facility"
		End With
		With .PivotFields("UHC Colo Screening (Comm/PCOR)")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("UHC Comm PCOR Attributed")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("UHC Comm PCOR Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC Colo Screening (Comm/PCOR)")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("H1")
		.Value = "UHC COMM PCOR: Colon Cancer Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographic PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Tufts LS Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("N1")
		.Value = "Tufts LS: Colon Cancer Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'4th pivot
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	With myPivotTable4
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'BCBSRI Num'/ 'BCBS-COMM Attributed'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Other" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographic PCP Facility"
		End With
		With .PivotFields("BCBSRI Num")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("BCBS-COMM Attributed")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("BCBS-COMM Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("BCBSRI Num")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("N32")
		.Value = "BCBS COMM: Colon Cancer Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'5th pivot
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	With myPivotTable5
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographic PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Tufts Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("T1")
		.Value = "Tufts: Colon Cancer Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'6th pivot
	Set myPivotTable6 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange6, TableName:="PivotTableNewSheet6")
	With myPivotTable6
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'BCBSRI Num'/ 'BCBS-MA Attributed'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Other" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographic PCP Facility"
		End With
		With .PivotFields("BCBSRI Num")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("BCBS-MA Attributed")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("BCBS-MA Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("BCBSRI Num")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("T32")
		.Value = "BCBS MA: Colon Cancer Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'7th pivot
	Set myPivotTable7 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange7, TableName:="PivotTableNewSheet7")
	With myPivotTable7
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographic PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("Z1")
		.Value = "CMS: Colon Cancer Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'create eight pivot table
	Set myPivotTable8 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange8, TableName:="PivotTableNewSheet8")
	With myPivotTable8
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographic PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("UHC MA Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("Z32")
		.Value = "UHC MA:  Colon Cancer Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'clean up
	With myDestinationWorksheet.Rows("10")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("41")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createDepressionPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Page2_InfoFieldsRemoved").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = core & "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A10").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("A35").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("G10").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet.Range("O10").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet.Range("V10").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "CORE Rate", "= 'YTD Numerator'/ 'Denominator'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("YTD Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("CORE Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 12+ Begin Msmt Yr")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 18+ Begin Msmt Yr")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With Range("A1")
		.Value = "CORE | Prev 5: Depression Screening and Follow Up Plan"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		'add calculated fields
		.CalculatedFields.Add "CORE Rate", "= 'YTD Numerator'/ 'Denominator'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			.PivotItems("Bald Hill Pedi").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "Bald Hill Pedi" Or Pi.Value = "Narragansett Bay Pedi" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Toll Gate Pedi" Or Pi.Value = "Waterman Pedi" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
				With .PivotFields("YTD Numerator")
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("CORE Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 12-21 at Year End")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	' 'add title
	With Range("A28")
		.Value = "Pedi Core"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated fields
		.CalculatedFields.Add "CORE Rate", "= 'YTD Numerator'/ 'Denominator'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("YTD Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("CORE Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Tufts Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 18+ Begin Msmt Yr")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("G1")
		.Value = "Tufts"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'create fourth pivot table
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	With myPivotTable4
		'add calculated fields
		.CalculatedFields.Add "CMS Rate", "= 'YTD Numerator'/ 'YTD Denominator-CMS'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("YTD Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("YTD Denominator-CMS")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("CMS Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Denominator-CMS")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 12+ Begin Msmt Yr")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("O1")
		.Value = "CMS Payer Specific (UPDATED per cms specs)"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create fifth pivot table
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	With myPivotTable5
		'add calculated fields
		.CalculatedFields.Add "CORE Rate", "= 'YTD Numerator'/ 'Denominator'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				ElseIf .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				ElseIf .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
			Next i
		End With
		With .PivotFields("YTD Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("CORE Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age 12+ Begin Msmt Yr")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Insurance Coverage Type")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Medicaid"
			On Error GoTo 0
		End With
		With .PivotFields("Primary Insurance Product")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.ClearAllFilters
				.PivotItems("NHP Medicaid").Visible = True
				For Each Pi In .PivotItems
					If Pi.Value = "NHP Medicaid" Or Pi.Value = "UHC Medicaid" Then
						Pi.Visible = True
					Else
						Pi.Visible = False
					End If
				Next Pi
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("V1")
		.Value = "AE Medicaid Measure"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'clean up
	With myDestinationWorksheet.Rows("11")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("36")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createFallRiskPivotTable()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'set Page1 as active worksheet
		Worksheets("Detail").Activate
		Set mySourceWorksheet = ActiveSheet
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(What:="Account")
	End With
	'obtain address of destination cell range
	myDestinationRange = myDestinationWorksheet.Range("A11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("G11").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	With myPivotTable
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				End If
				If .PivotItems(i).Name = "CIC" Then
					.PivotItems("CIC").Visible = False
				End If
				If .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
				If .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				End If
			Next i
		End With
		With .PivotFields("YTD Fall Risk Assessed")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Age 65+ Yrs in Msmt Year")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		'Add calculated field
		.CalculatedFields.Add "Rate","='YTD Fall Risk Assessed'/ 'Age 65+ Yrs in Msmt Year'", True
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
	End With
	
	'Add Core pivot table name
	With Range("A2")
		.Value = "CORE: Fall Risk Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	' Add text
	With Range("A28")
		.Value = "Processing"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("A29")
		.Value = "- Identify Exclusions documented YTD (Filter out those for non compliant patients)"
		'.Font.Size = 14
	End With
	
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			'test to see if exists
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				End If
				If .PivotItems(i).Name = "CIC" Then
					.PivotItems("CIC").Visible = False
				End If
				If .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
				If .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				End If
			Next i
		End With
		With .PivotFields("YTD Fall Risk Assessed")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Age 65+ Yrs in Msmt Year")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Fall Risk Assessed")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Exclusion (YTD) - Non Compliant ONLY")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "0"
			On Error GoTo 0
		End With
		'Add calculated field
		.CalculatedFields.Add "Rate","='YTD Fall Risk Assessed'/ 'Age 65+ Yrs in Msmt Year'", True
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			.Position = 4
			.Function = xlSum
		End With
	End With
' Add pivot table name
	With Range("G2")
		.Value = "CMS: Fall Risk Screening"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	' Add text
	With Range("G33")
		.Value = "Filters"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("G34")
		.Value = "- Active, Non-Deceased"
		'.Font.Size = 14
	End With
	With Range("G35")
		.Value = "- Final CMS Attributed or Medicare Ins Flag (Built into report) "
		'.Font.Size = 14
	End With
		With Range("G37")
		.Value = "Note:"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("G38")
		.Value = "- Leave Specialty & Other in for later confirmation"
		'.Font.Size = 14
	End With
	With Range("G40")
		.Value = "Processing"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("G41")
		.Value = "- Identify Exclusions documented YTD (Filter out those for non compliant patients) "
		'.Font.Size = 14
	End With
	'clean up
	With myDestinationWorksheet.Rows("11")
		.Replace what:="Values", replacement:="", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("12")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The source sheet name does not match the actual sheet name.")				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field or column name: " & Err.Description
				MsgBox(MsgTxt)
				'Delete the pivot table sheet if cannot create table
				Application.DisplayAlerts = False 'turn off alert
				'ActiveWorkbook.Sheets("PivotTables").Delete 
				ActiveSheet.Delete
				Application.DisplayAlerts = True 'turn alert back on 
			Case Else
				MsgBox("There was an undefined error")
		End Select
End Sub
Sub createInfluenzaPivotTablePivotTable()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'set Page1 as active worksheet
		Worksheets("Page1").Activate
		Set mySourceWorksheet = ActiveSheet
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(What:="Account")
	End With
	'obtain address of destination cell range
	myDestinationRange = myDestinationWorksheet.Range("A12").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("G12").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	With myPivotTable
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		'Add calculated field
		With .PivotFields("Seen This Flu Season")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		'Add calculated field
		.CalculatedFields.Add "Rate","=Numerator /Denominator", True
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
	End With
	
	'Add Core pivot table name
	With Range("A2")
		.Value = "CORE - Influenza Immunization"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			'test to see if exists
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Seen This Flu Season")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		'Add calculated field
		.CalculatedFields.Add "Rate","=Numerator /Denominator", True
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			.Position = 4
			.Function = xlSum
		End With
	End With
' Add pivot table name
	With Range("G2")
		.Value = "CMS - Influenza Immunization"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'clean up
	With myDestinationWorksheet.Rows("12")
		.Replace what:="Values", replacement:="", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("13")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With

	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The source sheet name does not match the actual sheet name.")
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field or column name: " & Err.Description
				MsgBox(MsgTxt)
				'Delete the pivot table sheet if cannot create table
				Application.DisplayAlerts = False 'turn off alert
				'ActiveWorkbook.Sheets("PivotTables").Delete 
				ActiveSheet.Delete
				Application.DisplayAlerts = True 'turn alert back on 
			Case Else
				MsgBox("There was an undefined error")
		End Select
End Sub
Sub createPneumoPivotTable()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'set Page1 as active worksheet
		Worksheets("Page1").Activate
		Set mySourceWorksheet = ActiveSheet
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(What:="Account")
	End With
	'obtain address of destination cell range
	myDestinationRange = myDestinationWorksheet.Range("A9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("G9").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	With myPivotTable
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				End If
				If .PivotItems(i).Name = "CIC" Then
					.PivotItems("CIC").Visible = False
				End If
				If .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
				If .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				End If
			Next i
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Age 65 or Older")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		'Add calculated field
		.CalculatedFields.Add "Rate","=Numerator /Age 65 or Older", True
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
	End With
	
	'Add Core pivot table name
	With Range("A2")
		.Value = "CORE: Prev 8 - Pneumo Vaccine "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).CreatePivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			'test to see if exists
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				End If
				If .PivotItems(i).Name = "CIC" Then
					.PivotItems("CIC").Visible = False
				End If
				If .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
				If .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				End If
			Next i
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Age 65 or Older")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Deceased")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Patient Status")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		'Add calculated field
		.CalculatedFields.Add "Rate","=Numerator /Age 65 or Older", True
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			.Position = 4
			.Function = xlSum
		End With
	End With
' Add pivot table name
	With Range("G2")
		.Value = "CMS: Prev 8 - Pneumo Vaccine "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	' Add text
	With Range("G28")
		.Value = "Processing"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("G29")
		.Value = "- FINAL_CMS Attributed or Medicare Patient (CMS Attributed OR Core Core Flag AND Medicare Ins. Product) "
		'.Font.Size = 14
	End With
	With Range("G32")
		.Value = "Filters"
		.Font.FontStyle = "Bold"
		'.Font.Size = 14
	End With
	With Range("G33")
		.Value = "- Active, Non-Deceased"
		'.Font.Size = 14
	End With
	With Range("G34")
		.Value = "- Final CMS Attributed or Medicare"
		'.Font.Size = 14
	End With
	'clean up
	With myDestinationWorksheet.Rows("10")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With

	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The source sheet name does not match the actual sheet name.")				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field or column name: " & Err.Description
				MsgBox(MsgTxt)
				'Delete the pivot table sheet if cannot create table
				Application.DisplayAlerts = False 'turn off alert
				'ActiveWorkbook.Sheets("PivotTables").Delete 
				ActiveSheet.Delete
				Application.DisplayAlerts = True 'turn alert back on 
			Case Else
				MsgBox("There was an undefined error")
		End Select
End Sub
Sub createTobaccoPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Page1").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = core & "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("A40").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("H9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet.Range("H40").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet.Range("O9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange6 = myDestinationWorksheet.Range("U9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange7 = myDestinationWorksheet.Range("AD9").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Numerator-Final'/ 'Coastal Core Flag'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Numerator-Final")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With Range("A1")
		.Value = "Core: Tobacco"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		'add calculated fields
		.CalculatedFields.Add "CMS Rate- Pop 2", "='Numerator- Pop2'/ 'Denominator -Pop2'", True
		.CalculatedFields.Add "CMS Rate- Pop 2 FIX", "='Numerator- Pop2 (Intervention after Assmt)'/ 'Denominator -Pop2'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Denominator -Pop2")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Numerator- Pop2") 
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("CMS Rate- Pop 2")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Numerator- Pop2 (Intervention after Assmt)")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("CMS Rate- Pop 2 FIX")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator -Pop2")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator- Pop2") 
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("*IPP_CPT and Age")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("A32")
		.Value = "CORE (Population 2)"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated fields
		.CalculatedFields.Add "BCBS Comm Rate", "= 'Numerator-Final'/ 'BCBS Comm Denom Tobacco'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			' For i = 1 To .PivotItems.Count
				' If .PivotItems(i).Name = "Other" Then
					' .PivotItems("Other").Visible = False
				' ElseIf .PivotItems(i).Name = "TO BE REVIEWED" Then
					' .PivotItems("TO BE REVIEWED").Visible = False
				' ElseIf .PivotItems(i).Name = "Specialty" Then
					' .PivotItems("Specialty").Visible = False
				' End If
			' Next i
		End With
		With .PivotFields("Numerator-Final")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("BCBS Comm Denom Tobacco")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("BCBS Comm Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("BCBS Comm Denom Tobacco")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator-Final") 
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("H1")
		.Value = "BCBS Comm: Tobacco"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	With myPivotTable4
		'add calculated fields
		.CalculatedFields.Add "Tufts Rate", "= 'Numerator-Final'/ 'Coastal Core Flag'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Numerator-Final")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Tufts Attributed")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Tufts Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Tufts Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator-Final")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("H34")
		.Value = "Tufts: Tobacco"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	With myPivotTable5
		'add calculated fields
		.CalculatedFields.Add "UHC Rate", "= 'Numerator-Final'/ 'Coastal Core Flag'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Numerator-Final")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("UHC Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("UHC Comm PCOR Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator-Final")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("O1")
		.Value = "UHC COMM PCOR: Tobacco"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	Set myPivotTable6 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange6, TableName:="PivotTableNewSheet6")
	With myPivotTable6
		'add calculated fields
		.CalculatedFields.Add "CMS Rate- Pop 2", "='Numerator- Pop2'/ 'Denominator -Pop2'", True
		.CalculatedFields.Add "CMS Rate- Pop 2 FIX", "='Numerator- Pop2 (Intervention after Assmt)'/ 'Denominator -Pop2'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Denominator -Pop2")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Numerator- Pop2") 
			.Orientation = xlDataField
			'.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("CMS Rate- Pop 2")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Numerator- Pop2 (Intervention after Assmt)")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("CMS Rate- Pop 2 FIX")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator -Pop2")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator- Pop2") 
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("U1")
		.Value = "CMS Attributed (Population 2): Tobacco"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	Set myPivotTable7 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange7, TableName:="PivotTableNewSheet7")
	With myPivotTable7
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator-Final'/ 'Coastal Core Flag'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				ElseIf .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				ElseIf .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
			Next i
		End With
		With .PivotFields("Numerator-Final")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Numerator-Final")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Insurance Coverage Type")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Medicaid"
			On Error GoTo 0
		End With
		With .PivotFields("Primary Insurance Product")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.ClearAllFilters
				.PivotItems("NHP Medicaid").Visible = True
				For Each Pi In .PivotItems
					If Pi.Value = "NHP Medicaid" Or Pi.Value = "UHC Medicaid" Then
						Pi.Visible = True
					Else
						Pi.Visible = False
					End If
				Next Pi
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AD1")
		.Value = "AE Medicaid Measure"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	

	
	'clean up
	With myDestinationWorksheet.Rows("10")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("41")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With

	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createCervicalCancerPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Page1").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = core & "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A9").Address(ReferenceStyle:=xlR1C1)
	'myDestinationRange2 = myDestinationWorksheet.Range("G9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("G9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet.Range("O9").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet.Range("O40").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator'", True
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With Range("A1")
		.Value = "Core 10: Cervical Cancer Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	With Range("A3")
		.Value = "Adult & Family Med ONLY"
		.Interior.ColorIndex = 6
	End With
	'create second pivot table
	' Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	' 'add, organize and format Pivot Table fields
	' With myPivotTable2
		' 'add calculated fields
		' .CalculatedFields.Add "UHC Rate", "= 'UHC Num (Comm/ Comm PCOR)'/ 'Denominator'", True
		' With .PivotFields("Demographic PCP Facility")
			' .Orientation = xlRowField
			' .Position = 1
			' .ClearAllFilters
			' For i = 1 To .PivotItems.Count
				' If .PivotItems(i).Name = "Other" Then
					' .PivotItems("Other").Visible = False
				' ElseIf .PivotItems(i).Name = "TO BE REVIEWED" Then
					' .PivotItems("TO BE REVIEWED").Visible = False
				' ElseIf .PivotItems(i).Name = "Specialty" Then
					' .PivotItems("Specialty").Visible = False
				' End If
			' Next i
		' End With
		' With .PivotFields("Denominator")
			' .Orientation = xlDataField
			' '.Position = 2
			' .Function = xlSum
		' End With
		' With .PivotFields("UHC Num (Comm/ Comm PCOR)")
			' .Orientation = xlDataField
			' '.Position = 1
			' .Function = xlSum
		' End With
		' With .PivotFields("Scheduled This Year")
			' .Orientation = xlDataField
			' '.Position = 2
			' .Function = xlSum
		' End With
		' With .PivotFields("UHC Rate")
			' .Orientation = xlDataField
			' '.Position = 3
			' .Function = xlSum
			' .NumberFormat = "0%"
		' End With
		' 'Filters
		' With .PivotFields("UHC Comm Attributed")
			' .Orientation = xlPageField
			' .Position = 1
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("Active Non Deceased Flag")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
		' With .PivotFields("HEDIS Denominator")
			' .Orientation = xlPageField
			' .Position = 2
			' .ClearAllFilters
			' On Error Resume Next
				' .CurrentPage = "1"
			' On Error GoTo 0
		' End With
	' End With
	' 'add title
	' With Range("G1")
		' .Value = "UHC: Cervical Cancer Screening "
		' .Font.FontStyle = "Bold"
		' .Font.Size = 14
	' End With
	' With Range("G3")
		' .Value = "Other/Specialty/CIC Removed"
		' .Interior.ColorIndex = 6
	' End With
	
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated fields
		.CalculatedFields.Add "UHC PCOR Rate", "= 'UHC Num (Comm/ Comm PCOR)'/ 'Denominator'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				ElseIf .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				ElseIf .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
			Next i
		End With
		With .PivotFields("UHC Num (Comm/ Comm PCOR)")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("UHC PCOR Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("UHC Comm PCOR Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("G1")
		.Value = "UHC PCOR: Cervical Cancer Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	With Range("G3")
		.Value = "Other/Specialty/CIC Removed"
		.Interior.ColorIndex = 6
	End With
	
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	With myPivotTable4
		'add calculated fields
		.CalculatedFields.Add "Tufts Rate", "= 'Numerator'/ 'Denominator'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				ElseIf .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				ElseIf .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
			Next i
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Tufts Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Tufts Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("O1")
		.Value = "Tufts: Cervical Cancer Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	With Range("O3")
		.Value = "Other/Specialty/CIC Removed"
		.Interior.ColorIndex = 6
	End With
	
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	With myPivotTable5
		'add calculated fields
		.CalculatedFields.Add "Tufts-Lifespan Rate", "= 'Numerator'/ 'Denominator'", True 
		With .PivotFields("Demographic PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			For i = 1 To .PivotItems.Count
				If .PivotItems(i).Name = "Other" Then
					.PivotItems("Other").Visible = False
				ElseIf .PivotItems(i).Name = "TO BE REVIEWED" Then
					.PivotItems("TO BE REVIEWED").Visible = False
				ElseIf .PivotItems(i).Name = "Specialty" Then
					.PivotItems("Specialty").Visible = False
				End If
			Next i
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Tufts-Lifespan Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Tufts-Lifespan Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("HEDIS Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("O32")
		.Value = "Tufts-Lifespan: Cervical Cancer Screening "
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	With Range("O34")
		.Value = "Other/Specialty/CIC Removed"
		.Interior.ColorIndex = 6
	End With
	
	'clean up
	With myDestinationWorksheet.Rows("10")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("41")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createDMCombinedPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Deduped Page").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = "YTDPivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A12").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("R12").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("Y12").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet.Range("AJ12").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet.Range("AJ42").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange6 = myDestinationWorksheet.Range("AT12").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange7 = myDestinationWorksheet.Range("AZ12").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange8 = myDestinationWorksheet.Range("BJ12").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange9 = myDestinationWorksheet.Range("BT12").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated fields
		.CalculatedFields.Add "YTD A1C<8 PCT", "='YTD A1C<8'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD A1C>9 or Missing PCT", "= 'YTD A1C>9 or Missing'/'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Monofilament PCT", "= 'YTD Monofilament'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Eye Exam FINAL PCT", "= 'YTD Eye Exam FINAL'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Nephropathy Screening PCT", "= 'YTD Nephropathy Screening'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "DM Pneumo (18-64) PCT", "= '18-64 Pneumo (Num)'/'18-64 Patients (Denom)'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("YTD A1C<8")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C<8 PCT")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD A1C>9 or Missing")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C>9 or Missing PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Monofilament")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD Monofilament PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Eye Exam FINAL")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD Eye Exam FINAL PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Nephropathy Screening")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD Nephropathy Screening PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("18-64 Pneumo (Num)")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("18-64 Patients (Denom)")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("DM Pneumo (18-64) PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("DMP")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "(All)"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With Range("A1")
		.Value = "COASTAL CORE"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		'add calculated fields
		.CalculatedFields.Add "YTD A1C>9 or Missing PCT", "= 'YTD A1C>9 or Missing'/'Age 18+ in Msmt Year'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		' With .PivotFields("YTD A1C<8")
			' .Orientation = xlDataField
			' .Position = 1
			' .Function = xlSum
		' End With
		With .PivotFields("YTD A1C>9 or Missing")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C>9 or Missing PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD A1C>9 or Missing")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("R1")
		.Value = "CMS FINAL"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated fields
		.CalculatedFields.Add "UHC A1C < 8 (Comm/PCOR) PCT", "= 'UHC A1C < 8 (Comm/PCOR)'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "UHC Eye Exam FINAL PCT", "= 'UHC Eye Exam FINAL'/'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "UHC Nephropathy (Comm/PCOR) PCT", "= 'UHC Nephropathy (Comm/PCOR)'/ 'Age 18+ in Msmt Year'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("UHC A1C < 8 (Comm/PCOR)")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("UHC A1C < 8 (Comm/PCOR) PCT")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("UHC Eye Exam FINAL")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("UHC Eye Exam FINAL PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("UHC Nephropathy (Comm/PCOR)")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("UHC Nephropathy (Comm/PCOR) PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC Comm PCOR Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC A1C < 8 (Comm/PCOR)")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC Nephropathy (Comm/PCOR)")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC Eye Exam FINAL")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("Y1")
		.Value = "UHC Commercial PCOR"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'4th pivot
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	With myPivotTable4
		'add calculated fields
		.CalculatedFields.Add "YTD A1C<=9 PCT", "= 'YTD A1C<=9'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Eye Exam FINAL PCT", "= 'YTD Eye Exam FINAL'/'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Nephropathy Screening PCT", "= 'YTD Nephropathy Screening'/ 'Age 18+ in Msmt Year'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("YTD Eye Exam FINAL")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("YTD Eye Exam FINAL PCT")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Nephropathy Screening")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD Nephropathy Screening PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD A1C<=9")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C<=9 PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC MA Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD A1C<=9")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Nephropathy Screening")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Eye Exam FINAL")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AJ1")
		.Value = "UHC MA PCOR"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'5th pivot
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	With myPivotTable5
		'add calculated fields
		.CalculatedFields.Add "YTD A1C<8 PCT", "='YTD A1C<8'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Nephropathy Screening PCT", "= 'YTD Nephropathy Screening'/ 'Age 18+ in Msmt Year'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("YTD Nephropathy Screening")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD Nephropathy Screening PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD A1C<8")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C<8 PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC MA Eligibility Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD A1C")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Nephropathy Screening")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AJ30")
		.Value = "UHC MA Eligibility"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'6th pivot
	Set myPivotTable6 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange6, TableName:="PivotTableNewSheet6")
	With myPivotTable6
		'add calculated fields
		.CalculatedFields.Add "YTD A1C<8 PCT", "='YTD A1C<8'/ 'Age 18+ in Msmt Year'", True  
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("YTD A1C<8")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C<8 PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Tufts Attributed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD A1C<8")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AT1")
		.Value = "Tufts"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'7th pivot
	Set myPivotTable7 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange7, TableName:="PivotTableNewSheet7")
	With myPivotTable7
		'add calculated fields
		.CalculatedFields.Add "BCBS Comm A1C< 8 PCT", "= 'BCBS Comm A1C< 8'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "BCBS Comm Nephropathy PCT", "= 'BCBS Comm Nephropathy'/'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "BCBS Comm Eye Exam FINAL PCT", "= 'BCBS Comm Eye Exam FINAL'/ 'Age 18+ in Msmt Year'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("BCBS Comm A1C< 8")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("BCBS Comm A1C< 8 PCT")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("BCBS Comm Nephropathy")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("BCBS Comm Nephropathy PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("BCBS Comm Eye Exam FINAL")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("BCBS Comm Eye Exam FINAL PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("BCBS Comm Denom DM")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("BCBS Comm A1C< 8")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("BCBS Comm Nephropathy")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("BCBS Comm Eye Exam FINAL")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AZ1")
		.Value = "BCBS-Comm"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create eight pivot table
	Set myPivotTable8 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange8, TableName:="PivotTableNewSheet8")
	With myPivotTable8
		'add calculated fields
		.CalculatedFields.Add "BCBS MA A1C Control <= 9 PCT", "= 'BCBS MA A1C Control <= 9'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "BCBS MA Nephropathy PCT", "= 'BCBS MA Nephropathy'/'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "BCBS MA Eye Exam FINAL PCT", "= 'BCBS MA Eye Exam FINAL'/ 'Age 18+ in Msmt Year'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("BCBS MA A1C Control <= 9")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("BCBS MA A1C Control <= 9 PCT")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("BCBS MA Nephropathy")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("BCBS MA Nephropathy PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("BCBS MA Eye Exam FINAL")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("BCBS MA Eye Exam FINAL PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("BCBS MA Denom DM")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("BCBS MA A1C Control <= 9")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("BCBS MA Nephropathy")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("BCBS MA Eye Exam FINAL")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("BJ1")
		.Value = "BCBS-MA"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create ninth pivot table
	Set myPivotTable9 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange9, TableName:="PivotTableNewSheet9")
	With myPivotTable9
		'add calculated fields
		.CalculatedFields.Add "YTD A1C<8 PCT", "='YTD A1C<8'/ 'Age 18+ in Msmt Year'", True  
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("YTD A1C<8")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C<8 PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("YTD A1C<8")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Insurance Coverage Type")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Medicaid"
			On Error GoTo 0
		End With
		With .PivotFields("Primary Insurance Product")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.ClearAllFilters
				.PivotItems("NHP Medicaid").Visible = True
				For Each Pi In .PivotItems
					If Pi.Value = "NHP Medicaid" Or Pi.Value = "UHC Medicaid" Then
						Pi.Visible = True
					Else
						Pi.Visible = False
					End If
				Next Pi
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("BT1")
		.Value = "AE Medicaid Measure"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'clean up
	With myDestinationWorksheet.Rows("13")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	
	'add provider pivot
	Set myDestinationWorksheet2 = ActiveWorkbook.Worksheets.Add ' add a worksheet
	myDestinationWorksheet2.Name = "YTDProviderPivot" & ActiveWorkbook.Worksheets.Count
	myDestinationRange2a = myDestinationWorksheet2.Range("A12").Address(ReferenceStyle:=xlR1C1)
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTablePCP = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet2.Name & "!" & myDestinationRange2a, TableName:="PivotTableNewSheetPCP")
	'add, organize and format Pivot Table fields
	With myPivotTablePCP
		'add calculated fields
		.CalculatedFields.Add "YTD A1C<8 PCT", "='YTD A1C<8'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD A1C>9 or Missing PCT", "= 'YTD A1C>9 or Missing'/'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Monofilament PCT", "= 'YTD Monofilament'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Eye Exam FINAL PCT", "= 'YTD Eye Exam FINAL'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Nephropathy Screening PCT", "= 'YTD Nephropathy Screening'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "DM Pneumo (18-64) PCT", "= '18-64 Pneumo (Num)'/'18-64 Patients (Denom)'", True
		With .PivotFields("Attributed Provider Name")
			.Orientation = xlRowField
			.Position = 1
		End With
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("YTD A1C<8 PCT")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD A1C>9 or Missing PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Monofilament PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Eye Exam FINAL PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Nephropathy Screening PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("DM Pneumo (18-64) PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		.ColumnGrand = False
		.RowGrand = False
		.RowAxisLayout xlTabularRow
		.DisplayErrorString = True
	End With
	'add title
	With Range("A1")
		.Value = "Providers"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'clean up
	With myDestinationWorksheet2.Rows("13")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	'add Non DMP pivot
	Set myDestinationWorksheet3 = ActiveWorkbook.Worksheets.Add ' add a worksheet
	myDestinationWorksheet3.Name = "NON DMP Rates Pivot"
	myDestinationRangeDMP = myDestinationWorksheet3.Range("A12").Address(ReferenceStyle:=xlR1C1)
	Set myPivotTableDMP = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:="'" & myDestinationWorksheet3.Name& & "'!" & myDestinationRangeDMP, TableName:="PivotTableNewSheetDMP")
	'add, organize and format Pivot Table fields
	With myPivotTableDMP
		'add calculated fields
		.CalculatedFields.Add "YTD A1C<8 PCT", "='YTD A1C<8'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD A1C>9 or Missing PCT", "= 'YTD A1C>9 or Missing'/'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Monofilament PCT", "= 'YTD Monofilament'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Eye Exam FINAL PCT", "= 'YTD Eye Exam FINAL'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "YTD Nephropathy Screening PCT", "= 'YTD Nephropathy Screening'/ 'Age 18+ in Msmt Year'", True 
		.CalculatedFields.Add "DM Pneumo (18-64) PCT", "= '18-64 Pneumo (Num)'/'18-64 Patients (Denom)'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("YTD A1C<8")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C<8 PCT")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD A1C>9 or Missing")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD A1C>9 or Missing PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Monofilament")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD Monofilament PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Eye Exam FINAL")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD Eye Exam FINAL PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Nephropathy Screening")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("YTD Nephropathy Screening PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("18-64 Pneumo (Num)")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("18-64 Patients (Denom)")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("DM Pneumo (18-64) PCT")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Age 18+ in Msmt Year")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Age Group")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "18-75 yrs"
			On Error GoTo 0
		End With
		With .PivotFields("DM Confirmed")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("DMP")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "(blank)"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'copy pivot
	' With ActiveWorkbook.Worksheets(myDestinationWorksheet3.Name)
		' Set pvt = .PivotTables("PivotTableNewSheetDMP")
		' pvt.TableRange2.Copy
	' End With
	' Set myDestinationWorksheet4 = ActiveWorkbook.Worksheets.Add ' add a worksheet
	' myDestinationWorksheet4.Name = "NON DMP Rates"
	' myDestinationWorksheet4.Range("A8").PasteSpecial xlPasteValues
	' myDestinationWorksheet4.Range("A8").PasteSpecial xlPasteFormats
	
	'copy pivot alt
	myDestinationWorksheet3.Range("A12").Select
	Call PivotCopyFormatValues
	'clean up
	With ActiveSheet.Rows("13")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With ActiveSheet.Rows("7")
		.Replace what:="1", replacement:="Y", LookAt:=xlPart, MatchCase:=False
	End With
	With ActiveSheet.Rows("9")
		.Replace what:="1", replacement:="Y", LookAt:=xlPart, MatchCase:=False
	End With
	With ActiveSheet.Rows("10")
		.Replace what:="1", replacement:="Y", LookAt:=xlPart, MatchCase:=False
	End With
	With ActiveSheet.Rows("8")
		.Replace what:="(blank)", replacement:="N", LookAt:=xlPart, MatchCase:=False
	End With
	ActiveSheet.Name = "NON DMP Rates"
	Application.DisplayAlerts = False
	myDestinationWorksheet3.Delete
	Application.DisplayAlerts = True
	'Add core pivot table name
	With Range("A1")
		.Value = "Coastal Core: Diabetes Rate Report (Non DMP)"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createCardiacPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Deduped").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("A35").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("G11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange4 = myDestinationWorksheet.Range("M11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange5 = myDestinationWorksheet.Range("S11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange6 = myDestinationWorksheet.Range("Z11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange7 = myDestinationWorksheet.Range("AF11").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange8 = myDestinationWorksheet.Range("AL11").Address(ReferenceStyle:=xlR1C1)
	'myDestinationRange9 = myDestinationWorksheet.Range("AL11").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With Range("A1")
		.Value = "Core Cardiac 19: HTN BP Control"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Attributed Provider Name")
			.Orientation = xlRowField
			.Position = 1
		End With
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		'Filters
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Coastal Core Flag")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Active Non Deceased Flag")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		'.ColumnGrand = False
		.RowGrand = False
		.RowAxisLayout xlTabularRow
		.DisplayErrorString = True
	End With
	'add title
	With Range("A29")
		.Value = "Provider Rates"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'CMS Numerator'/ 'CMS Denominator'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("CMS Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("CMS Denominator")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("CMS Numerator")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "0"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("CMS Denominator")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("MSSP Attributed")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("G1")
		.Value = "CMS: HTN BP Control"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'4th pivot
	Set myPivotTable4 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange4, TableName:="PivotTableNewSheet4")
	With myPivotTable4
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'BCBS MA Denom BP'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("BCBS MA Denom BP")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("BCBS MA Denom BP")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("M1")
		.Value = "BCBSRI-MA: HTN BP Control"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'5th pivot
	Set myPivotTable5 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange5, TableName:="PivotTableNewSheet5")
	With myPivotTable5
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'BCBS Comm Denom BP'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			' .PivotItems("East Greenwich").Visible = True
			' For Each Pi In .PivotItems
				' If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					' Pi.Visible = True
				' Else
					' Pi.Visible = False
				' End If
			' Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
		End With
		With .PivotFields("BCBS Comm Denom BP")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("BCBS Comm Denom BP")
			.Orientation = xlPageField
			'.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("S1")
		.Value = "BCBSRI-COMM: HTN BP Control"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'6th pivot
	Set myPivotTable6 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange6, TableName:="PivotTableNewSheet6")
	With myPivotTable6
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator_UHC_COMM PCOR'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator_UHC_COMM PCOR")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("UHC Comm PCOR Attributed")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Numerator_UHC_COMM PCOR")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "0"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("UHC Comm PCOR Attributed")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("Z1")
		.Value = "UHC COMM PCOR: HTN BP Control"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'7th pivot
	Set myPivotTable7 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange7, TableName:="PivotTableNewSheet7")
	With myPivotTable7
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "0"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Tufts Attributed")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AF1")
		.Value = "Tufts: HTN BP Control"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create eight pivot table
	Set myPivotTable8 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange8, TableName:="PivotTableNewSheet8")
	With myPivotTable8
		'add calculated fields
		.CalculatedFields.Add "Rate", "= 'Numerator'/ 'Denominator_For Max Rate'", True 
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
			.AutoSort Order:=xlAscending, Field:="Demographics PCP Facility"
		End With
		With .PivotFields("Numerator")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.00%"
		End With
		With .PivotFields("Scheduled This Year")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		'Filters
		With .PivotFields("Numerator")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = "0"
			On Error GoTo 0
		End With
		With .PivotFields("Deceased")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "No"
			On Error GoTo 0
		End With
		With .PivotFields("Status")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Active"
			On Error GoTo 0
		End With
		With .PivotFields("Denominator_For Max Rate")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Insurance Coverage Type")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Medicaid"
			On Error GoTo 0
		End With
		With .PivotFields("Primary Insurance Product")
			.Orientation = xlPageField
			''.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.ClearAllFilters
				.PivotItems("NHP Medicaid").Visible = True
				For Each Pi In .PivotItems
					If Pi.Value = "NHP Medicaid" Or Pi.Value = "UHC Medicaid" Then
						Pi.Visible = True
					Else
						Pi.Visible = False
					End If
				Next Pi
			On Error GoTo 0
		End With
	End With
	'add title
	With Range("AL1")
		.Value = "AE Medicaid Measure"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'clean up
	With myDestinationWorksheet.Rows("12")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub createSDOHPivotTables()
	On Error GoTo ErrorHandler
	'declare variables to hold row and column numbers that define source data cell range
	Dim myFirstRow As Long
	Dim myLastRow As Long
	Dim myFirstColumn As Long
	Dim myLastColumn As Long
	Dim foundCell As Range
	'declare variables to hold source and destination cell range address
	Dim mySourceData As String
	Dim myDestinationRange As String
	'declare object variables to hold references to source and destination worksheets, and new Pivot Table
	Dim mySourceWorksheet As Worksheet
	Dim myDestinationWorksheet As Worksheet
	Dim myPivotTable As PivotTable
	'identify source and destination worksheets
	With ActiveWorkbook
		'Set mySourceWorksheet = Worksheets(Worksheets.Count)
		Worksheets("Patients Details").Activate
		Set mySourceWorksheet = ActiveSheet
		Set myDestinationWorksheet = .Worksheets.Add ' add a worksheet
		myDestinationWorksheet.Name = core & "PivotTables" & .Worksheets.Count
		Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'-------------add formula---------------
	mySourceWorksheet.Activate
	startRow = foundCell.Row
	FoundCellColLtr = Split(Cells(1, foundCell.Column).Address, "$")(1)
	formulaStartRow = startRow + 1
	'get last row value
	lastRow = mySourceWorksheet.Range("A" & Rows.Count).End(xLUp).Row
	'get last column value
	lastCol = mySourceWorksheet.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
	FormCellNum = lastCol + 1
	FormCellLtr = Split(Cells(1, FormCellNum).Address, "$")(1)
	FormulaAddr = FormCellLtr & startRow
	'add heading
	With mySourceWorksheet.Range(FormulaAddr)
		.Value = "YTD Medicaid Numerator"
		.Interior.Color = RGB(231, 229, 229)
		.Font.Name = "Arial"
		.Font.Size = 8
		.VerticalAlignment = xlCenter
		.Borders.LineStyle = xlContinuous 
		.Borders.Weight = xlMedium 
		.Borders.Color = RGB(192, 192, 192) 
	End With
	'define formula
	'=IF(AND(AK5=1,AH5>=$AS$3),1,"") 
	'set comparison value for formula
	mySourceWorksheet.Range("AS3").Value = "2019-01-01"
	Set DenomElig = mySourceWorksheet.Range("A1:CZ20").Find(What:="Denominator (Eligible for Screening)", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	DenomEligColLtr = Split(Cells(1, DenomElig.Column).Address, "$")(1)
	Set SDOHDate = mySourceWorksheet.Range("A1:CZ20").Find(What:="SDOH Screened Date FINAL", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	SDOHDateColLtr = Split(Cells(1, SDOHDate.Column).Address, "$")(1)
	'add the formulas
	YTDMedNum = "=IF(AND(" & DenomEligColLtr & formulaStartRow & "=1," & SDOHDateColLtr &formulaStartRow & ">=$AS$3),1,"""")"
	mySourceWorksheet.Range(FormCellLtr & formulaStartRow & ":" & FormCellLtr & lastRow).Formula = YTDMedNum
	With mySourceWorksheet.Range(FormCellLtr & formulaStartRow & ":" & FormCellLtr & lastRow)
		.value = .value 'causes out of memory error
		' .Copy
		' .PasteSpecial xlPasteValues
	End With
	'------------------------------
	'obtain address of destination cell ranges
	myDestinationRange = myDestinationWorksheet.Range("A8").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange2 = myDestinationWorksheet.Range("H8").Address(ReferenceStyle:=xlR1C1)
	myDestinationRange3 = myDestinationWorksheet.Range("A33").Address(ReferenceStyle:=xlR1C1)
	'identify first row and first column of source data cell range
	myFirstRow = foundCell.Row
	myFirstColumn = 1
	With mySourceWorksheet.Cells
		.AutoFilter 'clear filters if they exist
		'find last row and last column of source data cell range
		myLastRow = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
		myLastColumn = .Find(what:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		'obtain address of source data cell range
		mySourceData = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn)).Address(ReferenceStyle:=xlR1C1)
	End With
	'create Pivot Table cache and create Pivot Table report based on that cache
	Set myPivotTable = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange, TableName:="PivotTableNewSheet")
	'add, organize and format Pivot Table fields
	Dim Pi As PivotItem
	With myPivotTable
		'add calculated field
		.CalculatedFields.Add "Adult Rate", "= 'Numerator (Screened in Current Mo.)'/ 'Denominator (Eligible for Screening)'", True
		.CalculatedFields.Add "Declined Rate", "= 'Declined Current Mo.'/ 'Denominator (Eligible for Screening)'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.PivotItems("East Greenwich").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "East Greenwich" Or Pi.Value = "EPIM" Or Pi.Value = "Garden City" Or Pi.Value = "Greenville" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Lincoln" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Newport" Or Pi.Value = "Providence Edgewood"  Or Pi.Value = "PIMS" Or Pi.Value = "Wakefield" Or Pi.Value = "Warren Ave" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Numerator (Screened in Current Mo.)")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator (Eligible for Screening)")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Adult Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.0%"
		End With
		With .PivotFields("Declined Current Mo.")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Declined Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.0%"
		End With
		'Filters
		With .PivotFields("Denominator (Eligible for Screening)")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	'Add core pivot table name
	With myDestinationWorksheet.Range("A1")
		.Value = "SDOH Screening in Adult and Family Practices"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create second pivot table
	Set myPivotTable2 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange2, TableName:="PivotTableNewSheet2")
	'add, organize and format Pivot Table fields
	With myPivotTable2
		'add calculated field
		.CalculatedFields.Add "Pedi Rate", "= 'Numerator (Screened in Current Mo.)'/ 'Denominator (Eligible for Screening)'", True
		.CalculatedFields.Add "Declined Rate", "= 'Declined Current Mo.'/ 'Denominator (Eligible for Screening)'", True
		With .PivotFields("Demographics PCP Facility")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			.PivotItems("Bald Hill Pedi").Visible = True
			For Each Pi In .PivotItems
				If Pi.Value = "Coastal Family Medicine" Or Pi.Value = "Bald Hill Pedi" Or Pi.Value = "Narragansett Bay Pedi" Or Pi.Value = "Hillside Family Medicine" Or Pi.Value = "Narragansett Family Medicine" Or Pi.Value = "Toll Gate Pedi" Or Pi.Value = "Waterman Pedi" Then
					Pi.Visible = True
				Else
					Pi.Visible = False
				End If
			Next Pi
		End With
		With .PivotFields("Numerator (Screened in Current Mo.)")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator (Eligible for Screening)")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Pedi Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0.0%"
		End With
		With .PivotFields("Declined Current Mo.")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Declined Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.0%"
		End With
		'Filters
		With .PivotFields("Denominator (Eligible for Screening)")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
	End With
	' add title
	With myDestinationWorksheet.Range("H1")
		.Value = "SDOH Screening in Pediatric and Family Practices"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	'create third pivot table
	Set myPivotTable3 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=mySourceWorksheet.Name & "!" & mySourceData).createPivotTable(TableDestination:=myDestinationWorksheet.Name & "!" & myDestinationRange3, TableName:="PivotTableNewSheet3")
	'add, organize and format Pivot Table fields
	With myPivotTable3
		'add calculated field
		.CalculatedFields.Add "Adult Rate", "= 'Numerator (Screened in Current Mo.)'/ 'Denominator (Eligible for Screening)'", True
		'.CalculatedFields.Add "Declined Rate in Current Mo.", "= 'Declined Current Mo.'/ 'Denominator (Eligible for Screening)'", True
		.CalculatedFields.Add "YTD Rate", "='YTD Medicaid Numerator'/'Denominator (Eligible for Screening)'", True
		With .PivotFields("Primary Insurance Product")
			.Orientation = xlRowField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.ClearAllFilters
				.PivotItems("RI Medicaid").Visible = True
				For Each Pi In .PivotItems
					If Pi.Value = "RI Medicaid" Or Pi.Value = "Tufts Medicaid" Then
						Pi.Visible = True
					Else
						Pi.Visible = False
					End If
				Next Pi
			On Error GoTo 0
		End With
		With .PivotFields("Numerator (Screened in Current Mo.)")
			.Orientation = xlDataField
			.Position = 1
			.Function = xlSum
		End With
		With .PivotFields("Denominator (Eligible for Screening)")
			.Orientation = xlDataField
			.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("Adult Rate")
			.Orientation = xlDataField
			.Position = 3
			.Function = xlSum
			.NumberFormat = "0%"
		End With
		With .PivotFields("YTD Medicaid Numerator")
			.Orientation = xlDataField
			'.Position = 2
			.Function = xlSum
		End With
		With .PivotFields("YTD Rate")
			.Orientation = xlDataField
			'.Position = 3
			.Function = xlSum
			.NumberFormat = "0.0%"
		End With
		' With .PivotFields("Declined Current Mo.")
			' .Orientation = xlDataField
			' '.Position = 2
			' .Function = xlSum
		' End With
		' With .PivotFields("Declined Rate in Current Mo.")
			' .Orientation = xlDataField
			' '.Position = 3
			' .Function = xlSum
			' .NumberFormat = "0.0%"
		' End With
		' With .PivotFields("Scheduled This Year")
			' .Orientation = xlDataField
			' '.Position = 2
			' .Function = xlSum
		' End With
		'Filters
		With .PivotFields("Denominator (Eligible for Screening)")
			.Orientation = xlPageField
			.Position = 1
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "1"
			On Error GoTo 0
		End With
		With .PivotFields("Insurance Coverage Type")
			.Orientation = xlPageField
			.Position = 2
			.ClearAllFilters
			On Error Resume Next
				.CurrentPage = "Medicaid"
			On Error GoTo 0
		End With
		With .PivotFields("YTD Medicaid Numerator")
			.Orientation = xlPageField
			'.Position = 2
			.ClearAllFilters
			On Error Resume Next
				'.CurrentPage = ""
			On Error GoTo 0
		End With
	End With
	'add title
	With myDestinationWorksheet.Range("A27")
		.Value = "AE Medicaid Measure"
		.Font.FontStyle = "Bold"
		.Font.Size = 14
	End With
	
	'clean up
	With myDestinationWorksheet.Rows("11")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	With myDestinationWorksheet.Rows("31")
		.Replace what:="Sum of", replacement:="", LookAt:=xlPart, MatchCase:=False
		.Replace what:="Row Labels", replacement:="PODs", LookAt:=xlPart, MatchCase:=False
	End With
	
	'ERROR HANDLING
	'source: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/on-error-statement
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler: 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9	' worksheet does not exist
				MsgBox("The expected sheet name 'Page1' does not match the actual sheet name.")
				' Insert code to handle this error
				
			Case 1004	' pivot table data source start range is wrong; check table location
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox(MsgTxt)
			Case Else
				MsgBox("There was an undefined cause for error: " & Err.Number)
		End Select
End Sub
Sub PivotCopyFormatValues()
	'source: https://www.contextures.com/excel-vba-pivot-table-paste-format.html#Manual
	'select pivot table cell first
	Dim ws As Worksheet
	Dim pt As PivotTable
	Dim rngPT As Range
	Dim rngPTa As Range
	Dim rngCopy As Range
	Dim rngCopy2 As Range
	Dim lRowTop As Long
	Dim lRowsPT As Long
	Dim lRowPage As Long
	Dim msgSpace As String
	
	On Error Resume Next
		Set pt = ActiveCell.PivotTable
		Set rngPTa = pt.PageRange
	On Error GoTo errHandler
	
	If pt Is Nothing Then
		MsgBox "Could not copy pivot table for active cell"
		GoTo exitHandler
	End If
	
	If pt.PageFieldOrder = xlOverThenDown Then
	  If pt.PageFields.Count > 1 Then
		msgSpace = "Horizontal filters with spaces." _
		  & vbCrLf _
		  & "Could not copy Filters formatting."
	  End If
	End If
	
	Set rngPT = pt.TableRange1
	lRowTop = rngPT.Rows(1).Row
	lRowsPT = rngPT.Rows.Count
	Set ws = Worksheets.Add
	Set rngCopy = rngPT.Resize(lRowsPT - 1)
	Set rngCopy2 = rngPT.Rows(lRowsPT)
	
	rngCopy.Copy Destination:=ws.Cells(lRowTop, 1)
	rngCopy2.Copy _
	  Destination:=ws.Cells(lRowTop + lRowsPT - 1, 1)
	
	If Not rngPTa Is Nothing Then
		lRowPage = rngPTa.Rows(1).Row
		rngPTa.Copy Destination:=ws.Cells(lRowPage, 1)
	End If
	
	ws.Columns.AutoFit
	If msgSpace <> "" Then
	  MsgBox msgSpace
	End If
	exitHandler:
		Exit Sub
	errHandler:
    MsgBox "Could not copy pivot table for active cell"
    Resume exitHandler
End Sub
