
''***********************ABOUT THIS PROCEDURE***************************************************
'*******What It Does**********************
'This procedure creates the QA and POD Exception Reports for either adult or pediatric patients using the FINAL OUTPUT sheet from the Unified Report. 
'For both populations, it adds payer attribution, supportive care, care management team, and next and last appointment details. For the adult populations, it also adds diabetes, colonoscopy, and mammography details.
'The QA report includes all of the added details and is saved as a single file; reports for each POD are saved as separate files and do not include all of the added detail.
'
'*******Prerequisites*********************
'Assumptions for use:
'-- You are using a FINAL OUTPUT sheet and the Unified Report as the primary processing file
'-- You have access to the following files
'-- 	*Active and Attributed Patient file
'-- 	*Colon Cancer Screening file
'-- 	*Mammography Screening file
'-- 	*Diabetes Information file
'
'*************Settings***************************
'The settings for this procedure are as follows:
' -- sMammoFname: enter the file name of the Mammography Screening file. Make sure to include the file extension
' -- sMammoFolder: enter the file path for the Mammography Screening file. Make sure to include the last backslash.
' -- sColoFname: enter the file name of the Colon Cancer Screening file. Make sure to include the file extension
' -- sColoFolder: enter the file path for the Colon Cancer Screening file. Make sure to include the last backslash.
' -- sDMFname: enter the file name of the Diabetes Information file. Make sure to include the file extension
' -- sDMFolder: enter the file path for the Diabetes Information file. Make sure to include the last backslash.
' -- sAAFname: enter the file name of the Active and Attributed file. Make sure to include the file extension
' -- sAAFolder: enter the file path for the Active and Attributed file. Make sure to include the last backslash.
' -- Adult: enter either True or False. True processes the adult exception reports; False processes the pediatric exception reports.
'
'************************************************************************************************
'**************************** REVISION HISTORY ********************************************************
'3/11/19: Initial release
'5/23/19: Added LTC care flag, made multiple output formatting revisions

'******************************************************************************************************

'Declare Public Variables
Public sMammoFname As String, sMammoFolder As String, sColoFname As String, sColoFolder As String, sDMFname As String, sDMFolder As String, sAAFname As String, sAAFolder As String, Adult As Boolean
'************************************************************************************************


Sub CreatePodExceptionReport()
	'************************SET THE FOLLOWING VARIABLE BEFORE USING*******************************
	'mammography file info
	sMammoFname = "Prev 3 - Mammography.xlsx" 'make sure to include extension
	sMammoFolder = "C:\Users\phodgkins\Documents\Processing\vlookuptest\" 'make sure to include the last backslash
	'colonoscopy file info
	sColoFname = "Prev 4 - Colon Cancer Screening.xlsx" 'make sure to include extension
	sColoFolder = "C:\Users\phodgkins\Documents\Processing\vlookuptest\" 'make sure to include the last backslash
	'diabetes file info
	sDMFname = "DM Combined - Adult.xlsx" 'make sure to include extension
	sDMFolder = "C:\Users\phodgkins\Documents\Processing\vlookuptest\" 'make sure to include the last backslash
	'active and attributed file info
	sAAFname = "Active and Attributed Patients v4.0404.xlsx" 'make sure to include extension
	sAAFolder = "C:\Users\phodgkins\Documents\Processing\vlookuptest\" 'make sure to include the last backslash
	Adult = True 'enter True for adult, False for pedi
	'************************************************************************************************

	Application.ScreenUpdating = False
	If Adult Then
		Call AddMammo(sMammoFname, sMammoFolder)
		Call AddDM(sDMFname, sDMFolder)
		Call AddColo(sColoFname, sColoFolder)
	End If
	Call AddDemos(sAAFname, sAAFolder)
	Call ReorderColumns
	Call CreateQACompleteException(Adult)
	Call CreateComposite(Adult)
	Call AddFormula4Split(Adult)
	Call SplitData
	Call CleanUp
	Application.ScreenUpdating = True
End Sub

Sub AddMammo(fileName As String, folderName As String)
	'This procedure adds vlookups to the Mammography Screening file
	'define variables
	Dim wbMammo As Workbook, wbDemo As Workbook
	Dim ws As Worksheet, shtMammo As Worksheet
	Dim lastRow As Long, lastCol As Long
	Dim foundCell As Range
	Dim startRow As String
	Dim formulaStartRow As String
	Dim wasOpen As Boolean
	
	
	Set wbDemo = ThisWorkbook
	wbDemo.Activate
	'identify source worksheet
	With wbDemo
		exists = VerifySheetname("FINAL OUTPUT")
		If Not exists Then
			MsgBox("Could not find the FINAL OUTPUT worksheet. Exiting procedure.")
			Exit Sub
		Else
			Set ws = Worksheets("FINAL OUTPUT") 
			ws.AutoFilterMode = False 'remove all filters
			Set foundCell = ws.Range("A1:Z20").Find(What:="Account", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
		End If
	End With
	
	ws.Activate
	'get last row value
	lastRow = ws.Range("A" & Rows.Count).End(xLUp).Row
	'get last column value
	lastCol = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
	'insert columns
	'get column before exceptions
	Set insertCell = ws.Range("A1:AZ20").Find(What:="Exception Status", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
	insertCellCol = insertCell.Column
	insertCellColLtr = Split(Cells(1, insertCell.Column).Address, "$")(1)
	numFormulas = 1
	For i=1 to numFormulas
		Columns(""& insertCellColLtr &":" & insertCellColLtr &"").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
	Next i
	startRow = foundCell.Row
	FoundCellColLtr = Split(Cells(1, foundCell.Column).Address, "$")(1)
	formulaStartRow = startRow + 1
	'get the column numbers
	firstFormCellNum = insertCellCol
	firstFormCellLtr = Split(Cells(1, firstFormCellNum).Address, "$")(1)
	firstFormAddr = firstFormCellLtr & startRow
	With ws.Range(firstFormAddr)
		.Value = "Mammo Open Order Date"
	End With
	'get mammo
	If IsWorkbookOpen(fileName) = False Then
		Set wbMammo = Workbooks.Open(folderName & fileName)
		wasOpen = False
	Else
		Workbooks(fileName).Activate
		Set wbMammo = ActiveWorkbook
		wasOpen = True
	End If
	
	wbMammo.Activate
	'set search values
	With wbMammo
		Set shtMammo = Worksheets("Detail") 
		Set mammoAcctNum = shtMammo.Range("A1:Z20").Find(What:="Account Number", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set mammoOrderDate = shtMammo.Range("A1:CZ20").Find(What:="Mammo Open Order Date", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'define the formulas; column letters
	mammoStartRow = mammoAcctNum.Row + 1 'add one to header row
	mammoStartColLtr = Split(Cells(1, mammoAcctNum.Column).Address, "$")(1)
	mammoOrderDateColLtr = Split(Cells(1, mammoOrderDate.Column).Address, "$")(1)
	mammoColsNum = ABS(mammoOrderDate.Column - mammoAcctNum.Column) + 1
	mammoLastRow = shtMammo.Range("A" & Rows.Count).End(xLUp).Row
	'vlookups
	'order date
	vLkupMammo = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & shtMammo.Name & "'!$" & mammoOrderDateColLtr & "$" & mammoStartRow  & ":$" & mammoStartColLtr & "$" & mammoLastRow & "," & mammoColsNum & ",FALSE)"
	'format as general
	ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).NumberFormat = "General"
	'add the formulas
	ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).Formula = vLkupMammo
	With ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow)
		.value = .value
		'remove 0 and #N/A
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.NumberFormat = "mmm d, yyyy"
	End With
	'close the file if it wasn't open before
	If Not wasOpen Then
		wbmammo.Close savechanges:=False
	End If
	
	ws.Activate 'set focus to demo file

End Sub

Sub AddColo(fileName As String, folderName As String)
	'This procedure adds vlookups to the Colon Cancer Screening file
	'define variables
	Dim wbColo As Workbook, wbDemo As Workbook
	Dim ws As Worksheet, shtColo As Worksheet
	Dim lastRow As Long, lastCol As Long
	Dim foundCell As Range
	Dim startRow As String
	Dim formulaStartRow As String
	Dim wasOpen As Boolean
	
	
	Set wbDemo = ThisWorkbook
	'identify source worksheet
	With wbDemo
		exists = VerifySheetname("FINAL OUTPUT")
		If Not exists Then
			MsgBox("Could not find the FINAL OUTPUT worksheet. Exiting procedure.")
			Exit Sub
		Else
			Set ws = Worksheets("FINAL OUTPUT") 
			ws.AutoFilterMode = False 'remove all filters
			Set foundCell = ws.Range("A1:Z20").Find(What:="Account", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
		End If
	End With
	
	ws.Activate
	'get last row value
	lastRow = ws.Range("A" & Rows.Count).End(xLUp).Row
	'get last column value
	lastCol = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
	'insert columns
	'get column before exceptions
	Set insertCell = ws.Range("A1:AZ20").Find(What:="Exception Status", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
	insertCellCol = insertCell.Column
	insertCellColLtr = Split(Cells(1, insertCell.Column).Address, "$")(1)
	numFormulas = 5
	For i=1 to numFormulas
		Columns(""& insertCellColLtr &":" & insertCellColLtr &"").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
	Next i
	startRow = foundCell.Row
	FoundCellColLtr = Split(Cells(1, foundCell.Column).Address, "$")(1)
	formulaStartRow = startRow + 1
	'get the column numbers
	firstFormCellNum = insertCellCol
	firstFormCellLtr = Split(Cells(1, firstFormCellNum).Address, "$")(1)
	firstFormAddr = firstFormCellLtr & startRow
	With ws.Range(firstFormAddr)
		.Value = "FOBT Open Order Date"
	End With
	secondFormCellNum = insertCellCol +1
	secondFormCellLtr = Split(Cells(1, secondFormCellNum).Address, "$")(1)
	secondFormAddr = secondFormCellLtr & startRow
	With ws.Range(secondFormAddr)
		.Value = "Cologuard Open Order Date"
	End With
	thirdFormCellNum = insertCellCol +2
	thirdFormCellLtr = Split(Cells(1, thirdFormCellNum).Address, "$")(1)
	thirdFormAddr = thirdFormCellLtr & startRow
	With ws.Range(thirdFormAddr)
		.Value = "Colonoscopy Open Order Date"
	End With
	fourthFormCellNum = insertCellCol +3
	fourthFormCellLtr = Split(Cells(1, fourthFormCellNum).Address, "$")(1)
	fourthFormAddr = fourthFormCellLtr & startRow
	With ws.Range(fourthFormAddr)
		.Value = "Colo Referral date"
	End With
	fifthFormCellNum = insertCellCol +4
	fifthFormCellLtr = Split(Cells(1, fifthFormCellNum).Address, "$")(1)
	fifthFormAddr = fifthFormCellLtr & startRow
	With ws.Range(fifthFormAddr)
		.Value = "Colo Referral status"
	End With
	'get Colo
	If IsWorkbookOpen(fileName) = False Then
		Set wbColo = Workbooks.Open(folderName & fileName)
		wasOpen = False
	Else
		Workbooks(fileName).Activate
		Set wbColo = ActiveWorkbook
		wasOpen = True
	End If
	
	wbColo.Activate
	'set search values
	With wbColo
		Set shtColo = Worksheets("Detail") 
		Set ColoAcctNum = shtColo.Range("A1:Z20").Find(What:="Account Number", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set FOBTOrderDate = shtColo.Range("A1:CZ20").Find(What:="FOBT Open Order Date", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set ColoOrderDate = shtColo.Range("A1:CZ20").Find(What:="Colonoscopy Open Order Date", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set CologuardOrderDate = shtColo.Range("A1:CZ20").Find(What:="Cologuard Open Order Date", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set ColoReferDate = shtColo.Range("A1:CZ20").Find(What:="Colo Referral date", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set ColoStatus = shtColo.Range("A1:CZ20").Find(What:="Colo Referral status", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'define the formulas; column letters
	ColoStartRow = ColoAcctNum.Row + 1 'add one to header row
	ColoStartColLtr = Split(Cells(1, ColoAcctNum.Column).Address, "$")(1)
	ColoOrderDateColLtr = Split(Cells(1, ColoOrderDate.Column).Address, "$")(1)
	ColoColsNum = ABS(ColoOrderDate.Column - ColoAcctNum.Column) + 1
	FOBTOrderDateColLtr = Split(Cells(1, FOBTOrderDate.Column).Address, "$")(1)
	FOBTColsNum = ABS(FOBTOrderDate.Column - ColoAcctNum.Column) + 1
	CologuardOrderDateColLtr = Split(Cells(1, CologuardOrderDate.Column).Address, "$")(1)
	CologuardColsNum = ABS(CologuardOrderDate.Column - ColoAcctNum.Column) + 1
	ColoRefDateColLtr = Split(Cells(1, ColoReferDate.Column).Address, "$")(1)
	ColoRefColsNum = ABS(ColoReferDate.Column - ColoAcctNum.Column) + 1
	ColoStatusColLtr = Split(Cells(1, ColoStatus.Column).Address, "$")(1)
	ColoStatusColsNum = ABS(ColoStatus.Column - ColoAcctNum.Column) + 1
	ColoLastRow = shtColo.Range("A" & Rows.Count).End(xLUp).Row
	
	'vlookups
	vLkupColo = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & shtColo.Name & "'!$" & ColoOrderDateColLtr & "$" & ColoStartRow  & ":$" & ColoStartColLtr & "$" & ColoLastRow & "," & ColoColsNum & ",FALSE)"
	vLkupFOBT = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & shtColo.Name & "'!$" & FOBTOrderDateColLtr & "$" & ColoStartRow  & ":$" & ColoStartColLtr & "$" & ColoLastRow & "," & FOBTColsNum & ",FALSE)"
	vLkupCologuard = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & shtColo.Name & "'!$" & CologuardOrderDateColLtr & "$" & ColoStartRow  & ":$" & ColoStartColLtr & "$" & ColoLastRow & "," & CologuardColsNum & ",FALSE)"
	vLkupRefDate = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & shtColo.Name & "'!$" & ColoRefDateColLtr & "$" & ColoStartRow  & ":$" & ColoStartColLtr & "$" & ColoLastRow & "," & ColoRefColsNum & ",FALSE)"
	vLkupColoStatus = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & shtColo.Name & "'!$" & ColoStatusColLtr & "$" & ColoStartRow  & ":$" & ColoStartColLtr & "$" & ColoLastRow & "," & ColoStatusColsNum & ",FALSE)"
	'format as general
	ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).NumberFormat = "General"
	'add the formulas
	ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).Formula = vLkupColo
	With ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow)
		.value = .value
		'remove 0 and #N/A
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.NumberFormat = "mmm d, yyyy"
	End With
	ws.Range(secondFormCellLtr & formulaStartRow & ":" & secondFormCellLtr & lastRow).NumberFormat = "General"
	ws.Range(secondFormCellLtr & formulaStartRow & ":" & secondFormCellLtr & lastRow).Formula = vLkupFOBT
	With ws.Range(secondFormCellLtr & formulaStartRow & ":" & secondFormCellLtr & lastRow)
		.value = .value
		'format the cells; remove 0 and #N/A
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.NumberFormat = "mmm d, yyyy"
	End With
	ws.Range(thirdFormCellLtr & formulaStartRow & ":" & thirdFormCellLtr & lastRow).NumberFormat = "General"
	ws.Range(thirdFormCellLtr & formulaStartRow & ":" & thirdFormCellLtr & lastRow).Formula = vLkupCologuard
	With ws.Range(thirdFormCellLtr & formulaStartRow & ":" & thirdFormCellLtr & lastRow)
		.value = .value
		'format the cells; remove 0 and #N/A
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.NumberFormat = "mmm d, yyyy"
	End With
	ws.Range(fourthFormCellLtr & formulaStartRow & ":" & fourthFormCellLtr & lastRow).NumberFormat = "General"
	ws.Range(fourthFormCellLtr & formulaStartRow & ":" & fourthFormCellLtr & lastRow).Formula = vLkupRefDate
	With ws.Range(fourthFormCellLtr & formulaStartRow & ":" & fourthFormCellLtr & lastRow)
		.value = .value
		'format the cells; remove 0 and N/A
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.NumberFormat = "mmm d, yyyy"
	End With
	ws.Range(fifthFormCellLtr & formulaStartRow & ":" & fifthFormCellLtr & lastRow).NumberFormat = "General"
	ws.Range(fifthFormCellLtr & formulaStartRow & ":" & fifthFormCellLtr & lastRow).Formula = vLkupColoStatus
	With ws.Range(fifthFormCellLtr & formulaStartRow & ":" & fifthFormCellLtr & lastRow)
		.value = .value
		'remove 0 and #N/A
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	'close the file if it wasn't open before
	If Not wasOpen Then
		wbColo.Close savechanges:=False
	End If
	
	ws.Activate 'set focus to demo file

End Sub

Sub AddDM(fileName As String, folderName As String)
	'This procedure adds vlookups to the Diabetes Information file
	'define variables
	Dim wbDM As Workbook, wbDemo As Workbook
	Dim ws As Worksheet, shtDM As Worksheet
	Dim lastRow As Long, lastCol As Long
	Dim foundCell As Range
	Dim startRow As String
	Dim formulaStartRow As String
	Dim wasOpen As Boolean
	
	
	Set wbDemo = ThisWorkbook
	'identify source worksheet
	With wbDemo
		exists = VerifySheetname("FINAL OUTPUT")
		If Not exists Then
			MsgBox("Could not find the FINAL OUTPUT worksheet. Exiting procedure.")
			Exit Sub
		Else
			Set ws = Worksheets("FINAL OUTPUT") 
			ws.AutoFilterMode = False 'remove all filters
			Set foundCell = ws.Range("A1:Z20").Find(What:="Account", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
		End If
	End With
	
	ws.Activate
	'get last row value
	lastRow = ws.Range("A" & Rows.Count).End(xLUp).Row
	'get last column value
	lastCol = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
	'insert columns
	'get column before exceptions
	Set insertCell = ws.Range("A1:AZ20").Find(What:="Exception Status", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
	insertCellCol = insertCell.Column
	insertCellColLtr = Split(Cells(1, insertCell.Column).Address, "$")(1)
	numFormulas = 2
	For i=1 to numFormulas
		Columns(""& insertCellColLtr &":" & insertCellColLtr &"").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
	Next i
	startRow = foundCell.Row
	FoundCellColLtr = Split(Cells(1, foundCell.Column).Address, "$")(1)
	formulaStartRow = startRow + 1
	'get the column numbers
	firstFormCellNum = insertCellCol
	firstFormCellLtr = Split(Cells(1, firstFormCellNum).Address, "$")(1)
	firstFormAddr = firstFormCellLtr & startRow
	With ws.Range(firstFormAddr)
		.Value = "A1C Open Order Date"
	End With
	secondFormCellNum = insertCellCol +1
	secondFormCellLtr = Split(Cells(1, secondFormCellNum).Address, "$")(1)
	secondFormAddr = secondFormCellLtr & startRow
	With ws.Range(secondFormAddr)
		.Value = "Microalbumin Open Order Date"
	End With
	'get DM
	If IsWorkbookOpen(fileName) = False Then
		Set wbDM = Workbooks.Open(folderName & fileName)
		wasOpen = False
	Else
		Workbooks(fileName).Activate
		Set wbDM = ActiveWorkbook
		wasOpen = True
	End If
	
	wbDM.Activate
	'set search values
	With wbDM
		Set shtDM = Worksheets("Deduped Page") 
		Set DMAcctNum = shtDM.Range("A1:Z20").Find(What:="Account Number", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set A1COrderDate = shtDM.Range("A1:CZ20").Find(What:="A1C Open Order Date", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set MicroOrderDate = shtDM.Range("A1:CZ20").Find(What:="Microalbumin Open Order Date", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	End With
	'define the formulas; column letters
	DMStartRow = DMAcctNum.Row + 1 'add one to header row
	DMStartColLtr = Split(Cells(1, DMAcctNum.Column).Address, "$")(1)
	A1COrderDateColLtr = Split(Cells(1, A1COrderDate.Column).Address, "$")(1)
	A1CColsNum = ABS(A1COrderDate.Column - DMAcctNum.Column) + 1
	MicroOrderDateColLtr = Split(Cells(1, MicroOrderDate.Column).Address, "$")(1)
	MicroColsNum = ABS(MicroOrderDate.Column - DMAcctNum.Column) + 1
	DMLastRow = shtDM.Range("A" & Rows.Count).End(xLUp).Row
	'vlookups
	vLkupA1C = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & shtDM.Name & "'!$" & A1COrderDateColLtr & "$" & DMStartRow  & ":$" & DMStartColLtr & "$" & DMLastRow & "," & A1CColsNum & ",FALSE)"
	vLkupMicro = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & shtDM.Name & "'!$" & MicroOrderDateColLtr & "$" & DMStartRow  & ":$" & DMStartColLtr & "$" & DMLastRow & "," & MicroColsNum & ",FALSE)"
	'format as general
	ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).NumberFormat = "General"
	'add the formulas
	ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).Formula = vLkupA1C
	With ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow)
		.value = .value
'		'format the cells; remove 0 and #N/A
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.NumberFormat = "mmm d, yyyy"
	End With
	ws.Range(secondFormCellLtr & formulaStartRow & ":" & secondFormCellLtr & lastRow).NumberFormat = "General"
	ws.Range(secondFormCellLtr & formulaStartRow & ":" & secondFormCellLtr & lastRow).Formula = vLkupMicro
	With ws.Range(secondFormCellLtr & formulaStartRow & ":" & secondFormCellLtr & lastRow)
		.value = .value
'		'format the cells; remove 0 and #N/A
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.NumberFormat = "mmm d, yyyy"
	End With
	
	'close the file if it wasn't open before
	If Not wasOpen Then
		wbDM.Close savechanges:=False
	End If
	
	ws.Activate 'set focus to demo file

End Sub

Sub AddDemos(fileName As String, folderName As String)
	'This procedure adds payer attribution, supportive care, care management team, and next and last appointment details information to the FINAL OUTPUT file. It uses the Active and Attributed file to add the information.
	'On Error GoTo ErrorHandler
	Dim wbAllPts As Workbook, wbDemo As Workbook
	Dim ws As Worksheet 
	Dim lastRow As Long, lastCol As Long
	Dim foundCell As Range
	Dim startRow As String
	Dim formulaStartRow As String
	Dim wasOpen As Boolean
	
	
	Set wbDemo = ThisWorkbook
	'identify source worksheet
	With wbDemo
		exists = VerifySheetname("FINAL OUTPUT")
		If Not exists Then
			MsgBox("Could not find the FINAL OUTPUT worksheet. Exiting procedure.")
			Exit Sub
		Else
			Set ws = Worksheets("FINAL OUTPUT") 
			ws.AutoFilterMode = False 'remove all filters
			Set foundCell = ws.Range("A1:Z20").Find(What:="Account", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
		End If
	End With
	
	ws.Activate
	'insert columns
	'get column before exceptions
	Set insertCell = ws.Range("A1:AZ20").Find(What:="Exception Status", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
	insertCellCol = insertCell.Column
	insertCellColLtr = Split(Cells(1, insertCell.Column).Address, "$")(1)
	numFormulas = 16
	For i=1 to numFormulas
		Columns(""& insertCellColLtr &":" & insertCellColLtr &"").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
	Next i
	'get last row value
	lastRow = ws.Range("A" & Rows.Count).End(xLUp).Row
	'get last column value
	lastCol = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
	FoundCellColLtr = Split(Cells(1, foundCell.Column).Address, "$")(1)
	lastColLtr = Split(Cells(1, lastCol).Address, "$")(1)
	startRow = foundCell.Row
	'sort on account number to speed processing
	sortingRange = FoundCellColLtr & foundCell.Row
	startAddress = "A" & foundCell.Row
	With ws.Sort
		.SortFields.Add Key:=Range(sortingRange), Order:=xlAscending
		.SetRange Range(startAddress).Resize(lastRow, lastCol)
		.Header = xlYes
		.Apply
	End With
	formulaStartRow = startRow + 1
	'get the column numbers
	firstFormCellNum = insertCellCol
	firstFormCellLtr = Split(Cells(1, firstFormCellNum).Address, "$")(1)
	firstFormAddr = firstFormCellLtr & startRow
	With ws.Range(firstFormAddr)
		.Value = "Last Appointment Time"
	End With
	secondFormCellNum = insertCellCol +1
	secondFormCellLtr = Split(Cells(1, secondFormCellNum).Address, "$")(1)
	secondFormAddr = secondFormCellLtr & startRow
	With ws.Range(secondFormAddr)
		.Value = "Last Appointment Provider"
	End With
	thirdFormCellNum = insertCellCol +2
	thirdFormCellLtr = Split(Cells(1, thirdFormCellNum).Address, "$")(1)
	thirdFormAddr = thirdFormCellLtr & startRow
	With ws.Range(thirdFormAddr)
		.Value = "Last Appointment Facility"
	End With
	fourthFormCellNum = insertCellCol +3
	fourthFormCellLtr = Split(Cells(1, fourthFormCellNum).Address, "$")(1)
	fourthFormAddr = fourthFormCellLtr & startRow
	With ws.Range(fourthFormAddr)
		.Value = "Last Appointment Visit Type"
	End With
	fifthFormCellNum = insertCellCol +4
	fifthFormCellLtr = Split(Cells(1, fifthFormCellNum).Address, "$")(1)
	fifthFormAddr = fifthFormCellLtr & startRow
	With ws.Range(fifthFormAddr)
		.Value = "Next Appointment Time"
	End With
	sixthFormCellNum = insertCellCol +5
	sixthFormCellLtr = Split(Cells(1, sixthFormCellNum).Address, "$")(1)
	sixthFormAddr = sixthFormCellLtr & startRow
	With ws.Range(sixthFormAddr)
		.Value = "Next Appointment Provider"
	End With
	seventhFormCellNum = insertCellCol +6
	seventhFormCellLtr = Split(Cells(1, seventhFormCellNum).Address, "$")(1)
	seventhFormAddr = seventhFormCellLtr & startRow
	With ws.Range(seventhFormAddr)
		.Value = "Next Appointment Facility"
	End With
	eighthFormCellNum = insertCellCol +7
	eighthFormCellLtr = Split(Cells(1, eighthFormCellNum).Address, "$")(1)
	eighthFormAddr = eighthFormCellLtr & startRow
	With ws.Range(eighthFormAddr)
		.Value = "Next Appointment Visit Type"
	End With
	ninthFormCellNum = insertCellCol +8
	ninthFormCellLtr = Split(Cells(1, ninthFormCellNum).Address, "$")(1)
	ninthFormAddr = ninthFormCellLtr & startRow
	With ws.Range(ninthFormAddr)
		.Value = "Payer Attribution"
	End With
	tenthFormCellNum = insertCellCol +9
	tenthFormCellLtr = Split(Cells(1, tenthFormCellNum).Address, "$")(1)
	tenthFormAddr = tenthFormCellLtr & startRow
	With ws.Range(tenthFormAddr)
		.Value = "Supportive Care"
	End With
	eleventhFormCellNum = insertCellCol +10
	eleventhFormCellLtr = Split(Cells(1, eleventhFormCellNum).Address, "$")(1)
	eleventhFormAddr = eleventhFormCellLtr & startRow
	With ws.Range(eleventhFormAddr)
		.Value = "Risk Strata"
	End With
	twelfthFormCellNum = insertCellCol + 11
	twelfthFormCellLtr = Split(Cells(1, twelfthFormCellNum).Address, "$")(1)
	twelfthFormAddr = twelfthFormCellLtr & startRow
	With ws.Range(twelfthFormAddr)
		.Value = "Care Management Team"
	End With
	thirteenthFormCellNum = insertCellCol + 12
	thirteenthFormCellLtr = Split(Cells(1, thirteenthFormCellNum).Address, "$")(1)
	thirteenthFormAddr = thirteenthFormCellLtr & startRow
	With ws.Range(thirteenthFormAddr)
		.Value = "Care Management Frequency"
	End With
	fourteenthFormCellNum = insertCellCol +13
	fourteenthFormCellLtr = Split(Cells(1, fourteenthFormCellNum).Address, "$")(1)
	fourteenthFormAddr = fourteenthFormCellLtr & startRow
	With ws.Range(fourteenthFormAddr)
		.Value = "Next Appointment"
	End With
	fifthteenthFormCellNum = insertCellCol +14
	fifthteenthFormCellLtr = Split(Cells(1, fifthteenthFormCellNum).Address, "$")(1)
	fifthteenthFormAddr = fifthteenthFormCellLtr & startRow
	With ws.Range(fifthteenthFormAddr)
		.Value = "High Risk Codes"
	End With
	sixthteenthFormCellNum = insertCellCol +15
	sixthteenthFormCellLtr = Split(Cells(1, sixthteenthFormCellNum).Address, "$")(1)
	sixthteenthFormAddr = sixthteenthFormCellLtr & startRow
	With ws.Range(sixthteenthFormAddr)
		.Value = "LTC"
	End With
	'get active and attributed
	Dim shtAllPts As Worksheet
	If IsWorkbookOpen(fileName) = False Then
		Set wbAllPts = Workbooks.Open(folderName & fileName)
		wasOpen = False
	Else
		Workbooks(fileName).Activate
		Set wbAllPts = ActiveWorkbook
		wasOpen = True
	End If
	
	wbAllPts.Activate
	'set search values
	With wbAllPts
		Set shtAllPts = .Worksheets(Worksheets.Count)
		Set allPtsAcctNum = shtAllPts.Range("A1:Z20").Find(What:="Account Number", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsProvAttr = shtAllPts.Range("A1:CZ20").Find(What:="Payer Attribution", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsSuptCare = shtAllPts.Range("A1:CZ20").Find(What:="Supportive Care", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsLstApptTime = shtAllPts.Range("A1:CZ20").Find(What:="Last Appointment Time", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsLstApptProv = shtAllPts.Range("A1:CZ20").Find(What:="Last Appointment Provider", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsLstApptFac = shtAllPts.Range("A1:CZ20").Find(What:="Last Appointment Facility", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsLstApptType = shtAllPts.Range("A1:CZ20").Find(What:="Last Appointment Visit Type", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsNxtAppt = shtAllPts.Range("A1:CZ20").Find(What:="Next Appointment", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsNxtApptTime = shtAllPts.Range("A1:CZ20").Find(What:="Next Appointment Time", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsNxtApptProv = shtAllPts.Range("A1:CZ20").Find(What:="Next Appointment Provider", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsNxtApptFac = shtAllPts.Range("A1:CZ20").Find(What:="Next Appointment Facility", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsNxtApptType = shtAllPts.Range("A1:CZ20").Find(What:="Next Appointment Visit Type", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsRisk = shtAllPts.Range("A1:CZ20").Find(What:="Risk Strata", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsCareTeam = shtAllPts.Range("A1:CZ20").Find(What:="Care Management Team", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsHighRisk = shtAllPts.Range("A1:CZ20").Find(What:="High Risk Codes", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set allPtsCareFreq = shtAllPts.Range("A1:CZ20").Find(What:="Care Management Frequency", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		allPtsLastRow = shtAllPts.Range("A" & Rows.Count).End(xLUp).Row
		allPtsSheetName = shtAllPts.Name
		allPtsLastCol = shtAllPts.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
		AllPtsLastColLtr = Split(Cells(1, allPtsLastCol).Address, "$")(1)
		allPtsStartColLtr = Split(Cells(1, allPtsAcctNum.Column).Address, "$")(1)
	End With
	'sort account numbers to speed processing
	allPtsSortingRange = allPtsStartColLtr & allPtsAcctNum.Row
	allPtsStartAddress = "A" & allPtsAcctNum.Row
	With ws.Sort
		.SortFields.Add Key:=Range(allPtsSortingRange), Order:=xlAscending
		.SetRange Range(allPtsStartAddress).Resize(allPtsLastRow, allPtsLastCol)
		.Header = xlYes
		.Apply
	End With
	'define the formulas; column letters
	allPtsStartRow = allPtsAcctNum.Row + 1 'add one to header row
	allPtsProvAttrColLtr = Split(Cells(1, allPtsProvAttr.Column).Address, "$")(1)
	allPtsColsNum = ABS(allPtsProvAttr.Column - allPtsAcctNum.Column) + 1
	allPtsSuptCareColLtr = Split(Cells(1, allPtsSuptCare.Column).Address, "$")(1)
	allPtsSuptCareNum = ABS(allPtsSuptCare.Column - allPtsAcctNum.Column) + 1
	allPtsLstApptTimeColLtr = Split(Cells(1, allPtsLstApptTime.Column).Address, "$")(1)
	allPtsLstApptTimeNum = ABS(allPtsLstApptTime.Column - allPtsAcctNum.Column) + 1
	allPtsLstApptProvColLtr = Split(Cells(1, allPtsLstApptProv.Column).Address, "$")(1)
	allPtsLstApptProvNum = ABS(allPtsLstApptProv.Column - allPtsAcctNum.Column) + 1
	allPtsLstApptFacColLtr = Split(Cells(1, allPtsLstApptFac.Column).Address, "$")(1)
	allPtsLstApptFacNum = ABS(allPtsLstApptFac.Column - allPtsAcctNum.Column) + 1
	allPtsLstApptTypeColLtr = Split(Cells(1, allPtsLstApptType.Column).Address, "$")(1)
	allPtsLstApptTypeNum = ABS(allPtsLstApptType.Column - allPtsAcctNum.Column) + 1
	allPtsNxtApptColLtr = Split(Cells(1, allPtsNxtAppt.Column).Address, "$")(1)
	allPtsNxtApptNum = ABS(allPtsNxtAppt.Column - allPtsAcctNum.Column) + 1
	allPtsNxtApptTimeColLtr = Split(Cells(1, allPtsNxtApptTime.Column).Address, "$")(1)
	allPtsNxtApptTimeNum = ABS(allPtsNxtApptTime.Column - allPtsAcctNum.Column) + 1
	allPtsNxtApptProvColLtr = Split(Cells(1, allPtsNxtApptProv.Column).Address, "$")(1)
	allPtsNxtApptProvNum = ABS(allPtsNxtApptProv.Column - allPtsAcctNum.Column) + 1
	allPtsNxtApptFacColLtr = Split(Cells(1, allPtsNxtApptFac.Column).Address, "$")(1)
	allPtsNxtApptFacNum = ABS(allPtsNxtApptFac.Column - allPtsAcctNum.Column) + 1
	allPtsNxtApptTypeColLtr = Split(Cells(1, allPtsNxtApptType.Column).Address, "$")(1)
	allPtsNxtApptTypeNum = ABS(allPtsNxtApptType.Column - allPtsAcctNum.Column) + 1
	allPtsRiskColLtr = Split(Cells(1, allPtsRisk.Column).Address, "$")(1)
	allPtsRiskNum = ABS(allPtsRisk.Column - allPtsAcctNum.Column) + 1
	allPtsCareTeamColLtr = Split(Cells(1, allPtsCareTeam.Column).Address, "$")(1)
	allPtsCareTeamNum = ABS(allPtsCareTeam.Column - allPtsAcctNum.Column) + 1
	allPtsCareFreqColLtr = Split(Cells(1, allPtsCareFreq.Column).Address, "$")(1)
	allPtsCareFreqNum = ABS(allPtsCareFreq.Column - allPtsAcctNum.Column) + 1
	allPtsHighRiskColLtr = Split(Cells(1, allPtsHighRisk.Column).Address, "$")(1)
	allPtsHighRiskNum = ABS(allPtsHighRisk.Column - allPtsAcctNum.Column) + 1
	'vlookups
	'provider attribution
	vLkupAllPtsPayerAttr = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsProvAttrColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsColsNum & ",FALSE)"
	'supportive care
	vLkupAllPtsSuptCare = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsSuptCareColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsSuptCareNum & ",FALSE)"
	'last appt
	vLkupAllPtsLstApptTime = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsLstApptTimeColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsLstApptTimeNum & ",FALSE)"
	vLkupAllPtsLstApptProv = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsLstApptProvColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsLstApptProvNum & ",FALSE)"
	vLkupAllPtsLstApptFac = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsLstApptFacColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsLstApptFacNum & ",FALSE)"
	vLkupAllPtsLstApptType = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsLstApptTypeColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsLstApptTypeNum & ",FALSE)"
	'next appt
	vLkupAllPtsNxtAppt = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsNxtApptColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsNxtApptNum & ",FALSE)"
	vLkupAllPtsNxtApptTime = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsNxtApptTimeColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsNxtApptTimeNum & ",FALSE)"
	vLkupAllPtsNxtApptProv = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsNxtApptProvColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsNxtApptProvNum & ",FALSE)"
	vLkupAllPtsNxtApptFac = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsNxtApptFacColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsNxtApptFacNum & ",FALSE)"
	vLkupAllPtsNxtApptType = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsNxtApptTypeColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsNxtApptTypeNum & ",FALSE)"
	'care teams
	vLkupAllPtsRisk = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsRiskColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsRiskNum & ",FALSE)"
	vLkupAllPtsCareTeam = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsCareTeamColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsCareTeamNum & ",FALSE)"
	vLkupAllPtsCareFreq = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsCareFreqColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsCareFreqNum & ",FALSE)"
	'high risk
	vLkupAllPtsHighRisk = "=VLOOKUP(" & FoundCellColLtr & formulaStartRow & ",'" & folderName & "[" & fileName & "]" & allPtsSheetName & "'!$" & allPtsHighRiskColLtr & "$" & allPtsStartRow  & ":$" & allPtsStartColLtr & "$" & allPtsLastRow & "," & allPtsHighRiskNum & ",FALSE)"
	ltcMapping = "=IF(ISBLANK(" & tenthFormCellLtr & allPtsStartRow & "),"""", ""Y"")"
	'add the formulas
	ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).Formula = vLkupAllPtsLstApptTime
	With ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(secondFormCellLtr & formulaStartRow & ":" & secondFormCellLtr & lastRow).Formula = vLkupAllPtsLstApptProv
	With ws.Range(secondFormCellLtr & formulaStartRow & ":" & secondFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(thirdFormCellLtr & formulaStartRow & ":" & thirdFormCellLtr & lastRow).Formula = vLkupAllPtsLstApptFac
	With ws.Range(thirdFormCellLtr & formulaStartRow & ":" & thirdFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(fourthFormCellLtr & formulaStartRow & ":" & fourthFormCellLtr & lastRow).Formula = vLkupAllPtsLstApptType
	With ws.Range(fourthFormCellLtr & formulaStartRow & ":" & fourthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(fifthFormCellLtr & formulaStartRow & ":" & fifthFormCellLtr & lastRow).Formula = vLkupAllPtsNxtApptTime
	With ws.Range(fifthFormCellLtr & formulaStartRow & ":" & fifthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(sixthFormCellLtr & formulaStartRow & ":" & sixthFormCellLtr & lastRow).Formula = vLkupAllPtsNxtApptProv
	With ws.Range(sixthFormCellLtr & formulaStartRow & ":" & sixthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(seventhFormCellLtr & formulaStartRow & ":" & seventhFormCellLtr & lastRow).Formula = vLkupAllPtsNxtApptFac
	With ws.Range(seventhFormCellLtr & formulaStartRow & ":" & seventhFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(eighthFormCellLtr & formulaStartRow & ":" & eighthFormCellLtr & lastRow).Formula = vLkupAllPtsNxtApptType
	With ws.Range(eighthFormCellLtr & formulaStartRow & ":" & eighthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(ninthFormCellLtr & formulaStartRow & ":" & ninthFormCellLtr & lastRow).Formula = vLkupAllPtsPayerAttr
	With ws.Range(ninthFormCellLtr & formulaStartRow & ":" & ninthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(tenthFormCellLtr & formulaStartRow & ":" & tenthFormCellLtr & lastRow).Formula = vLkupAllPtsSuptCare
	With ws.Range(tenthFormCellLtr & formulaStartRow & ":" & tenthFormCellLtr & lastRow)
		.value = .value
		.NumberFormat = "General"
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(eleventhFormCellLtr & formulaStartRow & ":" & eleventhFormCellLtr & lastRow).Formula = vLkupAllPtsRisk
	With ws.Range(eleventhFormCellLtr & formulaStartRow & ":" & eleventhFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(twelfthFormCellLtr & formulaStartRow & ":" & twelfthFormCellLtr & lastRow).Formula = vLkupAllPtsCareTeam
	With ws.Range(twelfthFormCellLtr & formulaStartRow & ":" & twelfthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(thirteenthFormCellLtr & formulaStartRow & ":" & thirteenthFormCellLtr & lastRow).Formula = vLkupAllPtsCareFreq
	With ws.Range(thirteenthFormCellLtr & formulaStartRow & ":" & thirteenthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End With
	ws.Range(fourteenthFormCellLtr & formulaStartRow & ":" & fourteenthFormCellLtr & lastRow).Formula = vLkupAllPtsNxtAppt
	With ws.Range(fourteenthFormCellLtr & formulaStartRow & ":" & fourteenthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		'format as date
		.NumberFormat = "mmm d, yyyy"
	End With
	ws.Range(fifthteenthFormCellLtr & formulaStartRow & ":" & fifthteenthFormCellLtr & lastRow).Formula = vLkupAllPtsHighRisk
	With ws.Range(fifthteenthFormCellLtr & formulaStartRow & ":" & fifthteenthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		'format as text
		.NumberFormat = "General"
	End With
	ws.Range(sixthteenthFormCellLtr & formulaStartRow & ":" & sixthteenthFormCellLtr & lastRow).Formula = ltcMapping
	With ws.Range(sixthteenthFormCellLtr & formulaStartRow & ":" & sixthteenthFormCellLtr & lastRow)
		.value = .value
		'remove 0
		'.Replace what:=0, Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		'.Replace what:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
		'format as date
		.NumberFormat = "General"
	End With
	'close the A/A file if it wasn't open before
	If Not wasOpen Then
		wbAllPts.Close savechanges:=False
	End If
	
	ws.Activate 'set focus to demo file
	'resize columns
	'ThisWorkbook.ActiveSheet.UsedRange.EntireColumn.AutoFit
	
	Exit Sub	' Exit to avoid handler if no errors occur
	ErrorHandler:	 'Error-handling
		Select Case Err.Number   ' Evaluate error number.
			Case 9  ' worksheet does not exist
				MsgBox ("The expected sheet name does not match the actual sheet name.")
			Case 91 'sort column names does not exist
				MsgBox ("The expected name of the column does not match the actual column name.")
			Case 1004	   ' data source start range is wrong or field name is wrong
				MsgTxt = "There was an issue with a data field: " & Err.Description
				MsgBox (MsgTxt)
			Case Else
				MsgBox ("There was an undefined cause for error: " & Err.Number)
		End Select
	Application.ScreenUpdating = True
End Sub

Sub ReorderColumns()
	'This procedure reorders the columns according to the order provided in the array assignment below.
	Dim arrColOrder As Variant, ndx As Integer
	Dim Found As Range, counter As Integer
	Dim foundCell As Range
	
	arrColOrder = Array("Patient Account Number", "Patient Name", "Patient Date of Birth", "Demographic PCP Facility", "Demographics PCP Name", "Demographics Rendering Provider Name", "Demographics Rendering Facility", "Attributed Provider", "Patient Deceased", "Patient Status", "Age at Begin Msmt Yr", "Age as of End Msmt Yr", "Primary Insurance Name", "Primary Insurance Subscriber No", "Primary Insurance Group No", "Insurance Product", "Last Appointment", "Last Appointment Time","Last Appointment Provider","Last Appointment Facility","Last Appointment Visit Type","Next Appointment","Next Appointment Time","Next Appointment Provider","Next Appointment Facility","Next Appointment Visit Type","Risk Strata","Care Management Team","Care Management Frequency", "Payer Attribution","Supportive Care","Patient Phone Number", "Email", "Email Status", "Web Enabled Flag", "Phone", "Voice Enabled Flag", "Text Enabled Flag", "Mobile")
	
	arrColOrderAlt = Array("Patient Account Number", "Patient Name", "Patient Date of Birth", "Demographics PCP Facility", "Demographics PCP Name", "Demographics Rendering Provider Name", "Demographics Rendering Facility", "Attributed Provider", "Patient Deceased", "Patient Status", "Age at Begin Msmt Yr", "Age as of End Msmt Yr", "Primary Insurance Name", "Primary Insurance Subscriber No", "Primary Insurance Group No", "Insurance Product", "Last Appointment", "Last Appointment Time","Last Appointment Provider","Last Appointment Facility","Last Appointment Visit Type","Next Appointment","Next Appointment Time","Next Appointment Provider","Next Appointment Facility","Next Appointment Visit Type","Risk Strata","Care Management Team","Care Management Frequency", "Payer Attribution","Supportive Care","Patient Phone Number", "Email", "Email Status", "Web Enabled Flag", "Phone", "Voice Enabled Flag", "Text Enabled Flag", "Mobile")
	
	counter = 1
	Set wbDemo = ThisWorkbook
	wbDemo.Activate
	'identify source worksheet
	With wbDemo
		exists = VerifySheetname("FINAL OUTPUT")
		If Not exists Then
			MsgBox("Could not find the FINAL OUTPUT worksheet. Exiting procedure.")
			Exit Sub
		Else
			Set ws = Worksheets("FINAL OUTPUT") 
			ws.AutoFilterMode = False 'remove all filters
			Set foundCell = ws.Range("A1:Z20").Find(What:="Account", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
			Set isDemographics = ws.Range("A1:Z20").Find(What:="Demographic PCP Facility", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
		End If
	End With
	
	myRows = foundCell.Row & ":" & foundCell.Row
	
	'rename next appointment column that contains numbers, not dates
	myFirstColumn = 1
	myLastColumn = ActiveSheet.Cells.Find(what:="*", LookIn:=xlFormulas, lookat:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
	
	For col = myFirstColumn to myLastColumn
		If ActiveSheet.Cells(2,col).value = "Next Appointment" AND ActiveSheet.Cells(3,col).NumberFormat <> "mmm d, yyyy" Then
			ActiveSheet.Cells(2,col).value = "Next Appointment Number" 
			Exit For
		End If
	
	Next col
	
	'use this array if Demographic PCP Facility
	If (Not isDemographics Is Nothing) Then
		For ndx = LBound(arrColOrder) To UBound(arrColOrder)
			Set Found = Rows(myRows).Find(arrColOrder(ndx), LookIn:=xlValues, LookAt:=xlWhole, _
							  SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
			If Not Found Is Nothing Then
				If Found.Column <> counter Then
					Found.EntireColumn.Cut
					Columns(counter).Insert Shift:=xlToRight
					Application.CutCopyMode = False
				End If
				counter = counter + 1
			End If
		Next ndx
	Else
		For ndx = LBound(arrColOrderAlt) To UBound(arrColOrderAlt)
			Set Found = Rows(myRows).Find(arrColOrderAlt(ndx), LookIn:=xlValues, LookAt:=xlWhole, _
							  SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
			If Not Found Is Nothing Then
				If Found.Column <> counter Then
					Found.EntireColumn.Cut
					Columns(counter).Insert Shift:=xlToRight
					Application.CutCopyMode = False
				End If
				counter = counter + 1
			End If
		Next ndx
	End If
	
End Sub

Sub CreateQACompleteException(Adult As Boolean)
	'This procedure creates a QA exception file, which includes all of the added columns.
	Dim fileNameQA As String
	Dim choice As Boolean : choice = Adult
	If choice Then
		fileNameQA = "Adult Complete Exception Report.xlsx"
	Else
		fileNameQA = "Pedi Complete Exception Report.xlsx"
	End If
	'move a copy of the FINAL OUTPUT sheet to a new book
	With ThisWorkBook
		Worksheets("FINAL OUTPUT").Copy
	End With
	For currentColumn = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
	columnHeading = ActiveSheet.UsedRange.Cells(2, currentColumn).Value
		Select Case columnHeading
			Case "Supportive Care", "Next Appointment Number"
				'Delete the column 
				ActiveSheet.Columns(currentColumn).Delete
			Case Else
				'Do nothing
		End Select
	Next
	'add filter
	'ActiveSheet.AutoFilterMode = False
	LastRow = ActiveSheet.UsedRange.Rows.Count
	LastColumn = ActiveSheet.UsedRange.Columns.Count
	LastColLtr = Split(Cells(1, LastColumn).Address, "$")(1)
	ActiveSheet.Range("A2:" & LastColLtr & LastRow).AutoFilter
	Call InsertNotes(choice)
	'replace title for QA
	If Adult Then
		ActiveSheet.UsedRange.Replace what:="Coastal Core: Adult Unified Exception Report", Replacement:="Coastal Core: Adult Preventative Measures Composite Report", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	Else
		ActiveSheet.UsedRange.Replace what:="Coastal Core: Pediatric Unified Exception Report", Replacement:="Coastal Core: Pediatric Preventative Measures Composite Report", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
	End If
	ActiveSheet.Cells.Font.Name = "Century Gothic"
	'save new book in same folder
	ActiveWorkbook.SaveAs ThisWorkBook.Path & "\" & fileNameQA
	ActiveWorkbook.Close savechanges:=False
End Sub

Sub CreateComposite(Adult As Boolean)
	'This procedure creates an exception file with only the columns required for the PODs. This file is used as the source for the POD exception reports.
	Dim fileName As String
	Dim currentColumn As Integer
	Dim columnHeading As String
	Dim choice As Boolean : choice = Adult
	fileName = "COMPOSITE FOR SPLIT.xlsx"
	
	'move a copy of the FINAL OUTPUT sheet to a new book
	With ThisWorkBook
		Worksheets("FINAL OUTPUT").Copy
	End With
	'save new book in same folder
	ActiveWorkbook.SaveAs ThisWorkBook.Path & "\" & fileName
	'delete extra columns
	'select start column to ensure proper formatting
	ActiveSheet.Range("A4").Select
	For currentColumn = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
	columnHeading = ActiveSheet.UsedRange.Cells(2, currentColumn).Value
		Select Case columnHeading
			Case "Last Appointment Time", "Last Appointment Provider", "Last Appointment Facility", "Last Appointment Visit Type", "Next Appointment Time", "Next Appointment Provider", "Next Appointment Facility", "Next Appointment Visit Type", "Risk Strata", "Care Management Team", "Care Management Frequency", "Voice Enabled Flag", "Text Enabled Flag", "Web Enabled Flag", "Email", "Email Status", "HTN Flag", "DM Flag", "FOBT Open Order Date", "Cologuard Open Order Date", "Colonoscopy Open Order Date", "Colo Referral date", "Colo Referral status", "A1C Open Order Date", "Microalbumin Open Order Date", "Mammo Open Order Date", "High Risk Codes", "Supportive Care", "Next Appointment Number"
				'Delete the column 
				ActiveSheet.Columns(currentColumn).Delete
			Case Else
				'Do nothing
		End Select
	Next
	'add filter
	'ActiveSheet.AutoFilterMode = False
	LastRow = ActiveSheet.UsedRange.Rows.Count
	LastColumn = ActiveSheet.UsedRange.Columns.Count
	LastColLtr = Split(Cells(1, LastColumn).Address, "$")(1)
	ActiveSheet.Range("A2:" & LastColLtr & LastRow).AutoFilter
	Call InsertNotes(choice)
	ActiveSheet.Cells.Font.Name = "Century Gothic"
End Sub

Sub InsertNotes(Choice As Boolean)
	'This procedure adds the notes to the top of the FINAL OUTPUT sheet
	Dim ws As Worksheet
	'Set ws = Worksheets("FINAL OUTPUT")
	Set ws = ActiveSheet
	myMonth = DateTime.Month(Date)
	myYear = DateTime.Year(Date)
	'to account for end of year
	If myMonth = 1 Then
		myMonth = 12
		myYear = myYear-1
	Else
		myMonth = myMonth-1
	End If
	myMonthName = MonthName(myMonth)
	ws.Activate
	myLastColumn = Range("A2:CZ3").Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
	lastDataColLtr = Split(Cells(1, myLastColumn).Address, "$")(1)
	myRange = "A3:" & lastDataColLtr &"3"
	With ws
		.Rows("1:4").Insert Shift:=xlShiftDown
		'merge cells
		With Range(myRange)
			.VerticalAlignment = xlCenter
			.Merge
			.interior.Color = RGB(204, 204, 204)
		End With
		'add the notes
		With Range("A1")
			If Choice Then
				.Value = "Coastal Core: Adult Unified Exception Report" 'POD
				'.Value = "Coastal Core: Adult Preventitive Measures Composite Report" ' QA
			Else
				.Value = "Coastal Core: Pediatric Unified Exception Report" 'POD
				'.Value = "Coastal Core: Pediatric Preventitive Measures Composite Report" 'QA
			End If
			.Font.FontStyle = "Bold"
			.Font.Size = 14
			.EntireRow.AutoFit
		End With
		With Range("A2")
			' .Value = "Report Month: " & Format(Date, "mmmm") & " " & DateTime.Year(Date) 'network version does not permit Format
			'.Value = "Report Month: " & myMonth & "/" & myYear
			.Value = "Report Month: " & myMonthName & " " & myYear
			.Font.Size = 10
			.EntireRow.AutoFit
		End With
		With Range("A3")
			If Choice Then
				.Value = CHR(149) & " Patients are included in the exception report if they have at least one of the following non-compliant Exception Statuses: "  & Chr(10) & _
				"            o Missing " & Chr(150) & " The patient has not yet received the specified screening/assessment. All patients with a missing status will have a due by date of 1/1/2019. "  & Chr(10) & _
				"            o Due in CY " & Chr(150) & " The patient's most recent screening/assessment was before the start of the calendar year. The patient must receive the specified screening/assessment (and if applicable, the appropriate follow up plan) by the designated Due by Date. "  & Chr(10) & _
				"            o Out of Range/No FU  " & Chr(150) & " The patient was screened during the calendar year, but there is  no documentation of appropriate treatment or follow-up"  & Chr(10) & _
				"            o Declined  " & Chr(150) & " The patient declined screening during the calendar year, and is not compliant for the measure"  & Chr(10) & _
				"            o Medical Reason or Allergic " & Chr(150) & " The patient has a documented medical reason for not being screened or vaccinated and is not compliant for the measure"  & Chr(10) & _
				"            o Incomplete Followup " & Chr(150) & " The patient requires a documented followup for the measure, but the documentation is incomplete."  & Chr(10) & _
				"            o Due/Not Scheduled " & Chr(150) & " The patient had a preventive visit in the prior year, but have not been scheduled for a preventive visit during the current year."  & Chr(10) & _
				"            o Scheduled " & Chr(150) & " The patient is compliant for the Adult Preventive Visit measure, but is considered an exception until the preventive visit is completed."  & Chr(10) & _
				CHR(149) & " The following Exception Statuses are used when a patient is compliant, or excluded from a particular measure:"  & Chr(10) & _
				"            o Compliant " & Chr(150) & " The patient is compliant for the measure"  & Chr(10) & _
				"            o N/A " & Chr(150) & " The patient does not qualify for the measure"  & Chr(10) & _
				"            o Exclusion " & Chr(150) & " The patient is excluded from the measure due to an allergy, but can still be counted as compliant if they receive the vaccine at a future date (applies to Flu Vaccine)"  & Chr(10) & _
				CHR(149) & " A patient's Due By Date is calculated from the patient's most recent date of service, or their birthday if they are aging into the measure (applies to Colonoscopy, Mammography, Fall Risk, and  Pneumo Vaccine)"
			Else
				.Value = "Notes and Instructions: " & Chr(10) & Chr(9) & CHR(149) & _
				" Patients are included in the exception report if they have one of the  following 'Exception Statuses':" & Chr(10) & Chr(9) & _
				"            o  Missing " & Chr(150) & " The patient has not yet received the specified screening/assessment. All patients with a missing status will have a due by date of 1/1/2018. " & Chr(10) & Chr(9) & Chr(9) & _
				"            o  Missed Opportunity " & Chr(150) & " The patient did not receive all the appropriate testing/screenings before the qualifying age specified by the measure." & Chr(10) & Chr(9) & Chr(9) & _
				"            o  Opportunity Available " & Chr(150) & " The patient has not received all the appropriate testings/screenings but has not yet aged out of the measure. The patient must receive the necessary screenings/testings by the indicated due by date. " & Chr(10) & Chr(9) & Chr(9) & _
				"            o  Due in CY " & Chr(150) & " The patient's most recent screening/assessment was before the start of the calendar year. The patient must receive the specified screening/assessment (and if applicable, the appropriate follow up plan) by the designated Due by Date. " & Chr(10) & Chr(9) & Chr(9) & _
				"            o  No Counseling OR No Percentile " & Chr(150) & " The patient's most recent BMI was taken during the calendar year, but there is either no documentation of  nutrition counseling, or if the patient is under 16, there is no BMI percentile calculation. " & Chr(10) & Chr(9) & Chr(9) & _
				"            o  N/A " & Chr(150) & " The patient does not qualify for the measure." & Chr(10) & Chr(9)  & _
				CHR(149) & " A patient's Due By Date is calculated from the patient's most recent date of service, or their birthday. "& Chr(10) & Chr(9) & CHR(149) & _
				" Prioritize patients who are missing the measure first, then focus on those who have a due by date in the next two months."
				' .Characters(1,22).Font.Bold = True
				' .Characters(132,12).Font.Italic = True
				' .Characters(290,19).Font.Italic = True
				' .Characters(434,22).Font.Italic = True
				' .Characters(663,9).Font.Italic = True
				' .Characters(913,31).Font.Italic = True
				' .Characters(1150,4).Font.Italic = True
			End If
			.Font.Size = 9
			.EntireRow.AutoFit
		End With
		If Choice Then
			.Rows("3:3").RowHeight = 210
		Else
			.Rows("3:3").RowHeight = 150
		End If
	End With
End Sub

Sub AddFormula4Split(Adult As Boolean)
	'This procedure adds formulas to the POD exception file that may be used to split the file into files for the PODs.
	Set ws = ActiveSheet
	Set foundCell = ws.Range("A1:Z20").Find(what:="Account", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	startRow = foundCell.row
	'get last row value
	lastRow = ws.Range("A" & Rows.Count).End(xLUp).Row
	formulaStartRow = startRow + 1
	With ws.Cells
		'find last column of source data cell range
		myLastColumn = .Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
	End With
	firstFormCellNum = myLastColumn + 1
	firstFormCellLtr = Split(Cells(1, firstFormCellNum).Address, "$")(1)
	firstFormAddr = firstFormCellLtr & startRow
	With ws.Range(firstFormAddr)
		.Value = "Split Key"
		.Interior.Color = RGB(255, 255, 51)
		.Font.FontStyle = "Bold"
	End With
	facility = Rows(startRow).Find("PCP Facility", LookIn:=xlValues, LookAt:=xlPart).Column
	facilityColLtr = Split(Cells(1, facility).Address, "$")(1)
	If Adult Then
		ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).Formula = "=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM($" & facilityColLtr & formulaStartRow &")&""-Adult Unified"",""Narragansett Family Medicine"",""NFM""),""Coastal Family Medicine"",""CFM""),""Hillside Family Medicine"",""HFM""),""Providence Edgewood"",""ProvEdgewood""),""Veterans Parkway"",""Veterans Pkwy""),""Narragansett Bay Pedi"",""NBP""), "" Adult"",""""),""East Greenwich"",""EG"")"
	Else
		ws.Range(firstFormCellLtr & formulaStartRow & ":" & firstFormCellLtr & lastRow).Formula = "=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM($" & facilityColLtr & formulaStartRow &")&""-Pedi Unified"",""Narragansett Family Medicine"",""NFM""),""Coastal Family Medicine"",""CFM""),""Hillside Family Medicine"",""HFM""),""Providence Edgewood"",""ProvEdgewood""),""Veterans Parkway"",""Veterans Pkwy""),""Narragansett Bay Pedi"",""NBP""), "" Pedi"",""""),""East Greenwich"",""EG"")"
	End If
End Sub

Sub SplitData()
	'this procedure splits worksheets based on a specified column value
	Dim SrcSheet As Worksheet
	Dim TrgSheet As Worksheet
	Dim SrcRow As Long
	Dim LastRow As Long
	Dim TrgRow As Long
	Dim Student As String
	
	Set SrcSheet = ActiveSheet
	Set foundName = SrcSheet.Range("A1:AZ20").Find(what:="Split Key", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
	foundSplitRow = foundName.Row
	NameCol = Rows(foundSplitRow).Find("Split Key", LookIn:=xlValues, LookAt:=xlWhole).Column
	HeaderRow = foundSplitRow
	FirstRow = foundSplitRow + 1
	LastRow = SrcSheet.Cells(SrcSheet.Rows.Count, NameCol).End(xlUp).Row
	
	For SrcRow = FirstRow To LastRow
		Student = SrcSheet.Cells(SrcRow, NameCol).Value
		Set TrgSheet = Nothing
		On Error Resume Next
		Set TrgSheet = Worksheets(Student)
		On Error GoTo 0
		If TrgSheet Is Nothing Then
			Set TrgSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
			TrgSheet.Name = Student
			SrcSheet.Rows(HeaderRow).Copy Destination:=TrgSheet.Rows(HeaderRow)
		End If
		TrgRow = TrgSheet.Cells(TrgSheet.Rows.Count, NameCol).End(xlUp).Row + 1
		SrcSheet.Rows(SrcRow).Copy Destination:=TrgSheet.Rows(TrgRow)
	Next SrcRow
End Sub

Sub CleanUp()
	'this procedure removes the formulas used to split the worksheets, adds the heading and saves each sheet as a new workbook.
	Dim ws As Worksheet
	Dim xPath As String
	xPath = Application.ActiveWorkbook.Path
	
	Set mySourceWorksheet = Worksheets("Final Output")
	Set foundCell = mySourceWorksheet.Range("A1:Z20").Find(What:="Account", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
	
	For Each ws in Worksheets
		ws.Activate
		Set foundSplitCell = Nothing
		Set foundSplitCell = ws.Range("A1:AZ20").Find(what:="Split Key", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		Set foundDemoPCPCell = Nothing
		Set foundDemoPCPCell = ws.Range("A1:AZ20").Find(what:=" PCP Facility", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
		If ws.name <> "FINAL OUTPUT" Then
			splitCellCol = foundSplitCell.Column
			Columns(splitCellCol).EntireColumn.Delete
			demoPCPCellCol = foundDemoPCPCell.Column
			'expand columns
			With ws.UsedRange
				.EntireColumn.AutoFit
			End With
			'add filter
			'ActiveSheet.AutoFilterMode = False
			LastRow = ws.UsedRange.Rows.Count
			LastColumn = ws.UsedRange.Columns.Count
			LastColLtr = Split(Cells(1, LastColumn).Address, "$")(1)
			ws.Range("A2:" & LastColLtr & LastRow).AutoFilter
			For i = 1 to foundCell.Row - 1
				mySourceWorksheet.Rows(i).Copy
				ws.Rows(i).PasteSpecial
			Next
			'save as new file
			Columns(demoPCPCellCol).EntireColumn.Delete
			ws.Copy
			Application.ActiveWorkbook.SaveAs Filename:=xPath & "\" & ws.Name & ".xlsx"
			Application.ActiveWorkbook.Close
		End If
	Next
End Sub

Function IsWorkBookOpen(Name As String) As Boolean
	Dim xWb As Workbook
	On Error Resume Next
	Set xWb = Application.Workbooks.Item(Name)
	IsWorkBookOpen = (Not xWb Is Nothing)
End Function

Function VerifySheetname(theName As String) As Boolean
	Dim mySheetName 
	mySheetName = theName
	
	'Verify that sheet name exists in the workbook
	Dim strSheetName As String, wks As Worksheet, bln As Boolean
	strSheetName = Trim(mySheetName)
	On Error Resume Next
	Set wks = ActiveWorkbook.Worksheets(strSheetName)
	On Error Resume Next
	If Not wks Is Nothing Then
		VerifySheetname = True
		Exit Function
	Else
		VerifySheetname = False
		Err.Clear
	End If
End Function

