Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Public ShPurchase As Object
Public ShSales As Object
Public ShValidation As Object

Public Function ProCheckDate(Opt As String, ColName As String, i_intRowCount As Integer) As Boolean
	Dim l_intCount As Integer
	Dim l_intRowCount As Integer
	Dim curCol As Integer
	Dim oSheet As Object
	Dim oCell As Object
	
	If Opt = "P" Then
		oSheet = ShPurchase
	Else
		oSheet = ShSales
	End If
	curCol = GetColIndex(oSheet, ColName)	
	For l_intCount = 3 To i_intRowCount - 1
		oCell = oSheet.getCellByPosition(curCol, l_intCount)
		If ValidateDate(oCell.String) = False Then
			MsgBox "Invalid Date [" & oCell.String & "] in cell " & oCell.AbsoluteName & "." & vbCRLF & "You have to enter date in DD-MM-YYYY format."
			SelectObject(oCell)
			ProCheckDate = False
			Exit Function
		End If
	Next l_intCount
	ProCheckDate = True
End Function

'This Function is used to Replace < ,' and & which creates error duing uploading
Public Sub ValidateString(Opt As String, ColName As String, i_intRowCount As Integer)
	Dim l_intLength As Integer
	Dim l_intCount As Integer
	Dim oSheet as Object
	Dim oCell as Object
	Dim curCol as Integer
	If Opt = "P" Then
		oSheet = ShPurchase
	Else
		oSheet = ShSales
	End If
	curCol = GetColIndex(oSheet, ColName)

	'Range(cell).Select
	For l_intCount = 3 To i_intRowCount - 1
		oCell = oSheet.getCellByPosition(curCol, l_intCount) 
		If InStr(oCell.String, "<") <> 0 Then
			oCell.String = Replace(oCell.String, "<", " ")
		ElseIf InStr(oCell.String, "&") <> 0 Then
			oCell.String = Replace(oCell.String, "&", " ")
		ElseIf InStr(oCell.String, "'") <> 0 Then
			oCell.String = Replace(oCell.String, "'", " ")
		ElseIf InStr(oCell.String, """") <> 0 Then
			oCell.String = Replace(oCell.String, """", " ")
		ElseIf InStr(oCell.String, ">") <> 0 Then
			oCell.String = Replace(oCell.String, ">", " ")
		ElseIf InStr(oCell.String, "|") <> 0 Then
			oCell.String = Replace(oCell.String, "|", " ")
		End If
	Next l_intCount
End Sub

Sub proValidate_Purchase()
	Dim oRange as Object
	Dim l_intRowCount as Integer
	oRange = FindUsedRange(ShPurchase)
	
	l_intRowCount = oRange.Rows.Count
	If l_intRowCount = 2 Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		MsgBox "No Data Entered", vbInformation + vbOKOnly
		Exit Sub
	End If
	If ProCheckMandatory("P", "A", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	If ProCheckDate("P", "B", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	If ProCheckMandatory("P", "C", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	ValidateString "P", "D", l_intRowCount
	ValidateString "P", "E", l_intRowCount
	If ProCheckMandatory("P", "F", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	If ProCheckMandatory("P", "G", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	If ProCheckMandatory("P", "H", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	ShValidation.Unprotect ("cmctcs")
	SetText(ShValidation, "B1", "Validated")
	ShValidation.Protect ("cmctcs")
	'ActiveWorkbook.SaveAs Filename:="c:\Deepu\TIN.csv", FileFormat:=xlCSV
	'GenerateXml "P"
	GenerateTxt "P"
End Sub

Sub proValidate_Sales()
	Dim l_intRowCount as Integer
	Dim oRange as Object
	oRange = FindUsedRange(ShSales)
	l_intRowCount = oRange.Rows.Count
	If l_intRowCount = 2 Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		MsgBox "No Data Entered", vbInformation + vbOKOnly
		Exit Sub
	End If
	If ProCheckMandatory("S", "A", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	If ProCheckDate("S", "B", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	If ProCheckMandatory("S", "C", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	ValidateString "S", "D", l_intRowCount
	ValidateString "S", "E", l_intRowCount
	If ProCheckMandatory("S", "F", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	If ProCheckMandatory("S", "G", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	If ProCheckMandatory("S", "H", l_intRowCount) = False Then
		ShValidation.Unprotect ("cmctcs")
		SetText(ShValidation, "B1", "Not Validated")
		ShValidation.Protect ("cmctcs")
		Exit Sub
	End If
	ShValidation.Unprotect ("cmctcs")
	SetText(ShValidation, "B1", "Validated")
	ShValidation.Protect ("cmctcs")
	'ActiveWorkbook.SaveAs Filename:="c:\Deepu\TIN.csv", FileFormat:=xlCSV
	'GenerateXml "S"
	GenerateTxt "S"
End Sub

Public Function ProCheckMandatory(Opt As String, ColName As String, i_intRowCount As Integer) As Boolean
	Dim curRow As Integer
	Dim curCol As Integer
	Dim oSheet As Object
	Dim oCell As Object
	If Opt = "P" Then
		oSheet = ShPurchase
	Else
		oSheet = ShSales
	End If
	curCol = GetColIndex(oSheet, ColName)
	'Check if any of the cells are empty and if so return false
	For curRow = 3 To i_intRowCount - 1
		oCell = oSheet.getCellByPosition(curCol, curRow)
		If oCell.String = "" Then
			MsgBox "Mandatory Information Missing [" & oCell.String & "] in cell " & oCell.AbsoluteName
			SelectObject(oCell)
			ProCheckMandatory = False
			Exit Function
		End If
	Next curRow
	ProCheckMandatory = True
End Function

Public Function ValidateDate(i_date As String) As Boolean
	If (Val(Mid(i_date, 4, 2)) < 1 Or Val(Mid(i_date, 4, 2)) > 12) Then
		ValidateDate = False
		Exit Function
	End If
	If Val(Mid(i_date, 7, 4)) < 1900 Then
		ValidateDate = False
		Exit Function
	End If
	If (IsDate(i_date)) = False Then
		ValidateDate = False
		Exit Function
	End If

	If CDate(i_date) > Date Then
		ValidateDate = False
		Exit Function
	End If
	ValidateDate = True
End Function

