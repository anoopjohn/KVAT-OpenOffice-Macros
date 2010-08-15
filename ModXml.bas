Rem Attribute VBA_ModuleType=VBAModule

Option Explicit
Option VBASupport 1

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

'@file - ModXml.bas

Sub GenerateXml(Opt As String)
	Dim strXML As String
	Dim path As String
	Dim DirPath As String
	Dim oRange As Object

	DirPath = GetDirPath()

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
	MsgBox ("Completed. XML Written to " & path & ".")
	
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
	MsgBox ("Completed. File Written to " & path & ".")
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
		For intRowCounter = 1 To intRowCount - 1
			strTemp = ""
			'looping for numcols + 1 to add the CRLF in the last go
			For intColCounter = 0 To intColCount 
				If intColCounter < intColCount - 1 Then
					rngCell = rngData.getCellByPosition(intColCounter,intRowCounter)
					If strTemp <> "" Then
						strTemp = strTemp & "|" & Trim(rngCell.String)
					Else
						strTemp = strTemp & intRowCounter & "|" & Trim(rngCell.String)
					End If
				'Added to Sum up the Cess Amt, Value of Goods and VAT amount
                ElseIf intColCounter = intColCount - 1 Then
                    rngCell1 = rngData.getCellByPosition(intColCounter - 3,intRowCounter)
                    rngCell2 = rngData.getCellByPosition(intColCounter - 2,intRowCounter)
                    strTemp = strTemp & "|" & clng(Trim(rngCell1.String)) + clng(Trim(rngCell2.String)) + clng(Trim(rngCell.String)) & vbCRLF
				End If
			Next
			strtxt = strtxt & strTemp
		Next
	End With
	fGenerateTxt = strtxt
End Function

Function GetDirPath()
	On Error Resume Next
	Dim DirPath As String
	Dim strTempPath As String
	Dim i As Integer
	Randomize 2^14-1
	'If C:\KVAT exists then use that else if C:\ exists then try creating C:\KVAT 
	'else if the directory mentioned in the settings file exists then use that 
	'else ask the user 
	Do While True 
	    i = i + 1
	    If i > 2 Then
			DirPath = InputBox ("Please enter the path to the directory where you wish to save the files. Remember that this will overwrite the previous Purchase.txt and Sales.txt if any present in the folder", "Path to writable direcotry") 
			'DirPath = "/home/user01/test"
	    ElseIf i = 2 Then
	    	DirPath = ShSettings.getCellByPosition(1,0).String
	    ElseIf i = 1 Then
	    	DirPath = "C:\KVATS"
		End If	
		'Stop if user cancels the Input box
		If DirPath = "" Then
			End
		End If
		If Dir(DirPath, vbDirectory) = "" Then
			If i > 2 Then
				MsgBox DirPath & " is not a valid directory name. Please try again."
			Else 
				'Try creating the directory
				MkDir DirPath
			End If	
		Else
			'Check if we have write access to the directory by creating and deleting a test file
			strTempPath = ConvertFromURL(ConvertToURL(DirPath) & "/tmp" & Rnd & ".txt"
			sWriteFile "Test", strTempPath
			If Not Err Then
				Kill strTempPath
				If Not Err Then
					Exit Do
				End If
			Else
				MsgBox "Cannot write to " & DirPath & ". Please give a directory with write access."
			End If		
		End If
	Loop
	GetDirPath = DirPath
End Function 

' Function for writing plain string into a file
Sub sWriteFile(strText As String, strFullFileName As String)
	Dim intFileNum As String
	intFileNum = FreeFile
	Open strFullFileName For Output As #intFileNum
	Print #intFileNum, strText
	Close #intFileNum
End Sub

'Finds the range of cells that constitute the minimum rectangular area used in the sheet
Function FindUsedRange(ByVal oSheet as Object)
	Dim oCell As Object
	Dim oCursor As Object
	Set oCell = oSheet.getCellByPosition(0, 0)
	oCursor = oSheet.createCursorByRange(oCell)
	oCursor.gotoStartOfUsedArea(False)
	oCursor.GotoEndOfUsedArea(True)
	FindUsedRange = oCursor
End Function

'Initializes global variables
Sub InitProc
	ShPurchase = ThisComponent.Sheets.getByName("Purchases")
	ShSales = ThisComponent.Sheets.getByName("Sales")
	ShSettings = ThisComponent.Sheets.getByName("Settings")
End Sub

'Finds the index of a sheet by name
Function FindSheetIndex(SheetName As String) As Integer
	Dim i As Integer
	For i = 0 To ThisComponent.Sheets.Count - 1
		If ThisComponent.Sheets.getByIndex(i).Name = SheetName Then
			FindSheetIndex = i
			Exit Function
		End If
	Next i
	FindSheetIndex = -1
End Function

'Sets the text in a Cell
Sub SetText(ByVal oSheet As Object, Address As String, Value As String)
    Dim oRange As Object	
    Dim oCell As Object
    oRange = oSheet.getCellRangeByName(Address)
    oCell = oSheet.getCellByPosition(oRange.RangeAddress.StartColumn, oRange.RangeAddress.StartRow)
    oCell.String = Value
End Sub

'Get the Index of a column given the column name. 
Function GetColIndex(ByVal oSheet As Object, ColName As String)
	Dim oRange As Object
    oRange = oSheet.getCellRangeByName(ColName & "1")
    GetColIndex = oRange.RangeAddress.StartColumn
End Function

'Get the address of a cell
Function GetCellAddress(ByRef oCell As Object)
	Dim Address As String
	XRay oCell
	GetCellAddress = Address
End Function

'Select a given object
Sub SelectObject(obj)
	ThisComponent.getCurrentController().Select(obj)
End Sub
