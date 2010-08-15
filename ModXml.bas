Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

'@file - ModXml.bas

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
	
'	DirPath = GetDirPath()
    If Dir("c:\KVATS",16)= "" then ' Does the directory exist ?
		MkDir "c:\KVATS"
	End If

	If Opt = "P" Then
		oRange = FindUsedRange(ShPurchase)
		strtxt = fGenerateTxt(oRange)
		'path = ConvertFromURL(ConvertToURL(DirPath) & "/Purchase.txt")
		 path = "C:\KVATS\Purchase.txt"
	Else
		oRange = FindUsedRange(ShSales)
		strtxt = fGenerateTxt(oRange)
	'	path = ConvertFromURL(ConvertToURL(DirPath) & "/Sales.txt")
	    path = "C:\KVATS\Sales.txt"
	End If
	sWriteFile strtxt, path
	MsgBox ("Completed. File Written to " & path)
End Sub

Function fGenerateTxt(ByVal rngData As Object) As String
	Dim intColCount As Double
	Dim intRowCount As Double
	Dim intColCounter As Double
	Dim intRowCounter As Double
	Dim rngCell As Object
	Dim rngCell1 as object
	Dim rngCell2 as object
	Dim strtxt As String
	Dim strTemp As String
	With rngData
		strtxt = ""
		intColCount = .Columns.Count
		intRowCount = .Rows.Count
		'Loop down the table's rows
		For intRowCounter = 1 To intRowCount -1
			strTemp = ""
			'looping for numcols + 1 to add the CRLF in the last go
			For intColCounter = 0 To intColCount 
				If intColCounter < intColCount -1 Then
					rngCell = rngData.getCellByPosition(intColCounter,intRowCounter)
					If strTemp <> "" Then
							strTemp = strTemp & "|" & Trim(rngCell.String)
					Else
							strTemp = strTemp & intRowCounter & "|" & Trim(rngCell.String)
					End If
				'Added to Sum up the Cess Amt, Value of Goods and VAT amount
                ElseIf intColCounter = intColCount -1 Then
                    rngCell1 = rngData.getCellByPosition(intColCounter - 3,intRowCounter)
                    rngCell2 = rngData.getCellByPosition(intColCounter - 2,intRowCounter)
                    strTemp = strTemp & "|" & clng(Trim(rngCell1.String)) + clng(Trim(rngCell2.String)) + clng(Trim(rngCell.String)) & vbCRLF
				'Else
				'	strTemp = strTemp & vbCRLF
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
