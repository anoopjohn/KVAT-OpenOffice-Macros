Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Sub GenerateXml(Opt As String)
	Dim strXML As String
	Dim path As String
	Dim DirPath As String
	Dim oRange As Object

	DirPath = GetDirPath()

	' sWriteFile strXML, ThisWorkbook.Path & filenameinput
	If Opt = "P" Then
		oRange = FindUsedRange(ShPurchase)
		strXML = fGenerateXML(Opt, oRange, "DATA")
		path = ConvertFromURL(ConvertToURL(DirPath) & "/Purchase.XML")
	Else
		oRange = FindUsedRange(ShSales)
		strXML = fGenerateXML(Opt, oRange, "DATA")
		path = ConvertFromURL(ConvertToURL(DirPath) & "/Sales.XML")
	End If
	sWriteFile strXML, path
	MsgBox ("Completed. XML Written to " & path)
FolderError:
	Exit Sub
	Resume Next
End Sub

Sub GenerateTxt(Opt As String)
	Dim strtxt As String
	Dim path As String
	Dim oRange As Object
	Dim DirPath As String
	
	DirPath = GetDirPath()


	If Opt = "P" Then
		oRange = FindUsedRange(ShPurchase)
		strtxt = fGenerateTxt(oRange)
		path = ConvertFromURL(ConvertToURL(DirPath) & "/Purchase.txt")
	Else
		oRange = FindUsedRange(ShSales)
		strtxt = fGenerateTxt(oRange)
		path = ConvertFromURL(ConvertToURL(DirPath) & "/Sales.txt")
	End If
	sWriteFile strtxt, path
	MsgBox ("Completed. File Written to " & path)
End Sub

Function fGenerateXML(Opt As String, ByVal rngData As Object, rootNodeName As String) As String
	Const HEADER As String = "<?xml version=""1.0""?>"
	Dim TAG_BEGIN  As String
	Dim TAG_END  As String
	Dim ROW_BEGIN As String
	Dim ROW_END As String
	Dim intColCount As Double
	Dim intRowCount As Double
	Dim intColCounter As Double
	Dim intRowCounter As Double
	Dim rngCell As Object
	Dim strXML As String
	Dim strTabCol() As String
	'initial tag
	TAG_BEGIN = vbCrLf & "<" & rootNodeName & ">"
	TAG_END = vbCrLf & "</" & rootNodeName & ">"
	ROW_BEGIN = vbCrLf & "<ROW>"
	ROW_END = "</ROW>"
	strXML = HEADER
	strXML = strXML & TAG_BEGIN
	With rngData
		'Discover dimensions of the data we
		'will be dealing with...
		intColCount = .Columns.Count
		
		intRowCount = .Rows.Count
		
		Dim strColNames() As String
		
		ReDim strColNames(intColCount)
		ReDim strTabCol(intColCount)
		If Opt = "P" Then
		'Hard Coded to Get the Table Column
			strTabCol(1) = "RPDISERIAL_ID"
			strTabCol(2) = "RPDVINVOICE_NO"
			strTabCol(3) = "RPDDINVOICE_DATE"
			strTabCol(4) = "RPDVSALES_DEALER_NAME"
			strTabCol(5) = "RPDVSALES_DEALER_ADDRESS"
			strTabCol(6) = "RPDISALES_REGN_ID"
			strTabCol(7) = "RPDIVALUE_OF_GOODS"
			strTabCol(8) = "RPDIVAT_AMOUNT_COLLECT"
			strTabCol(9) = "RPDITOTAL_INVOICE_AMOUNT"
		Else
			strTabCol(1) = "RPDISERIAL_ID"
			strTabCol(2) = "RPDVINVOICE_NO"
			strTabCol(3) = "RPDDINVOICE_DATE"
			strTabCol(4) = "RPDVPURCHASE_DEALER_NAME"
			strTabCol(5) = "RPDVPURCHASE_DEALER_ADDRESS"
			strTabCol(6) = "RPDIPURCHASE_REGN_ID"
			strTabCol(7) = "RPDIVALUE_OF_GOODS"
			strTabCol(8) = "RPDIVAT_AMOUNT_COLLECT"
			strTabCol(9) = "RPDITOTAL_INVOICE_AMOUNT"
		End If
		'First Row is the Field/Tag names
		If intRowCount >= 1 Then
			'   Loop accross columns...
			For intColCounter = 1 To intColCount
				
				'   Mark the cell under current scrutiny by setting
				'   an object variable...
				Set rngCell = .Cells(1, intColCounter)
				'   Is the cell merged?..
				If Not rngCell.MergeArea.Address = rngCell.Address Then
					  MsgBox ("!! Cell Merged ... Invalid format")
					  Exit Function
				End If
				'strColNames(intColCounter) = rngCell.Text
				strColNames(intColCounter) = strTabCol(intColCounter)
			Next
		End If
		'Loop down the table's rows
		For intRowCounter = 2 To intRowCount
			strXML = strXML & ROW_BEGIN & vbCrLf
			ReDim NodeStack(0)
			'Loop accross columns...
			For intColCounter = 1 To intColCount + 1
				'The Column is iterated to last column +1 to find the row terminator
				If intColCounter <= intColCount Then
				Set rngCell = .Cells(intRowCounter, intColCounter)
					'   Is the cell merged?..
					If Not rngCell.MergeArea.Address = rngCell.Address Then
						  MsgBox ("!! Cell Merged ... Invalid format")
						  Exit Function
					End If
					'If Trim(rngCell.Text) <> "" Then
						  strXML = strXML & "<" & Trim(strColNames(intColCounter)) & ">" & Trim(rngCell.Text) & "</" & Trim(strColNames(intColCounter)) & ">" & vbCrLf
					'End If
				Else
					strXML = strXML & ROW_END & vbCrLf
				End If
			Next
		Next
	End With
	strXML = strXML & TAG_END
	'Return the HTML string...
	fGenerateXML = strXML
End Function

Function fGenerateTxt(ByVal rngData As Object) As String
	Dim intColCount As Double
	Dim intRowCount As Double
	Dim intColCounter As Double
	Dim intRowCounter As Double
	Dim rngCell As Object
	Dim strtxt As String
	Dim strTemp As String
	With rngData
		strtxt = ""
		intColCount = .Columns.Count
		intRowCount = .Rows.Count
		'Loop down the table's rows
		For intRowCounter = 3 To intRowCount - 1
			strTemp = ""
			'looping for numcols + 1 to add the CRLF in the last go
			For intColCounter = 0 To intColCount
				If intColCounter < intColCount Then
					rngCell = rngData.getCellByPosition(intColCounter, intRowCounter)
					'Is the cell merged?..
					If rngCell.getIsMerged() Then
						MsgBox ("!! Cell Merged ... Invalid format")
						Exit Function
					End If
					' skip if no content
					If Trim(rngCell.String) <> "" Then
						If strTemp <> "" Then
							strTemp = strTemp & "|" & Trim(rngCell.String)
						Else
							strTemp = strTemp & intRowCounter - 2 & "|" & Trim(rngCell.String)
						End If
					End If
				Else
					strTemp = strTemp & vbCRLF
				End If
			Next
			strtxt = strtxt & strTemp
		Next
	End With
	fGenerateTxt = strtxt
End Function

Function GetDirPath()
	Dim DirPath As String
REDO:	
	DirPath = InputBox ("Please enter the path to the directory where you wish to save the files. Remember that this will overwrite the previous Purchase.txt and Sales.txt if any present in the folder", "Path to writable direcotry") 
	'DirPath = "/home/user01/test"
	If DirPath = "" Then
		End
	End If
	If Dir(DirPath, vbDirectory) = "" Then
		MsgBox DirPath & " is not a valid directory name. Please try again"
		Goto REDO
	End If
	GetDirPath = DirPath
End Function 

' Function for writing plain string into a file
Sub sWriteFile(strXML As String, strFullFileName As String)
	Dim intFileNum As String
	intFileNum = FreeFile
	Open strFullFileName For Output As #intFileNum
	Print #intFileNum, strXML
	Close #intFileNum
End Sub

'Finds the range of cells that constitute the minimum rectangular area used in the sheet
Function FindUsedRange(ByVal oSheet as Object)
	Dim oCell As Object
	Dim oCursor As Object
	oCell = oSheet.getCellByPosition(0, 0)
	oCursor = oSheet.createCursorByRange(oCell)
	oCursor.gotoStartOfUsedArea(False)
	oCursor.GotoEndOfUsedArea(True)
	FindUsedRange = oCursor
End Function

'Initializes global variables
Sub InitProc
	ShPurchase = ThisComponent.Sheets.getByName("Purchases")
	ShSales = ThisComponent.Sheets.getByName("Sales")
	ShValidation = ThisComponent.Sheets.getByName("Validation")
End Sub

'Finds the index of a sheet by name
Function findSheetIndex(SheetName as String) as Integer
	Dim i as integer
	for i = 0 to ThisComponent.Sheets.Count - 1
		if ThisComponent.Sheets.getByIndex(i).Name = _
			SheetName _
		then
			findSheetIndex = i
			exit function
		end if
	next i
	findSheetIndex = -1
End Function

'Sets the text in a Cell
Sub SetText(ByVal oSheet as Object, Address as String, Value as String)
    Dim oRange as Object	
    Dim oCell as Object
    oRange = oSheet.getCellRangeByName(Address)
    oCell = oSheet.getCellByPosition(oRange.RangeAddress.StartColumn, oRange.RangeAddress.StartRow)
    oCell.String = Value
End Sub

'Get the Index of a column given the column name. 
Function GetColIndex(ByVal oSheet as Object, ColName as String)
	Dim oRange as Object
    oRange = oSheet.getCellRangeByName(ColName & "1")
    GetColIndex = oRange.RangeAddress.StartColumn
End Function

'Get the address of a cell
Function GetCellAddress(ByRef oCell as Object)
	Dim Address as String
	XRay oCell
	GetCellAddress = Address
End Function

'Select a given object
Sub SelectObject(obj)
	ThisComponent.getCurrentController().Select(obj)
End Sub
