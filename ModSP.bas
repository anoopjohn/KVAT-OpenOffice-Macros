Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Public ShPurchase As Object
Public ShSales As Object
Public ShValidation As Object

Public Function ProCheckDate(Opt As String, ColName As String, i_intRowCount As Integer) As Boolean
	Dim curCol As Integer
	Dim curRow As Integer
	Dim oSheet As Object
	Dim oCell As Object
	
	If Opt = "P" Then
		oSheet = ShPurchase
	Else
		oSheet = ShSales
	End If
	curCol = GetColIndex(oSheet, ColName)	
	For curRow = 1 To i_intRowCount - 1
		oCell = oSheet.getCellByPosition(curCol, curRow)
		'oCell.String = Format(oCell.String, "DD-MM-YYYY")
		If ValidateDate(oCell.String) = False Then
			MsgBox "Invalid Date [" & oCell.String & "] in cell " & oCell.AbsoluteName & "." & vbCRLF & "You have to enter date in DD-MM-YYYY format."
			SelectObject(oCell)
			ProCheckDate = False
			Exit Function
		End If
	Next curRow
	ProCheckDate = True
End Function



'Added------------------------------------------------------------------------------------------------------------------
Public Function ProCheckAmount(Opt As String, PrevCell As String,ColName As String, i_intRowCount As Double) As Boolean
Dim oSheet As Object
Dim oCell As Object
Dim l_intCell As Double
Dim l_intPrev As Double
Dim curRow As Integer
Dim curCol As Integer


If Opt = "P" Then
   oSheet = ShPurchase
Else
   oSheet = ShSales
End If
For curRow = 1 To i_intRowCount - 1
curCol = GetColIndex(oSheet, ColName)
oCell = oSheet.getCellByPosition(curCol, curRow)
l_intCell = CDbl(oCell.String)
curCol = GetColIndex(oSheet, PrevCell)
oCell = oSheet.getCellByPosition(curCol, curRow)
l_intPrev = CDbl(oCell.String)
If l_intCell > l_intPrev Then
     SelectObject(oCell)
     ProCheckAmount = False
   Exit Function
End If
Next curRow
ProCheckAmount = True
End Function


Public Sub RemoveSpace(Opt As String, cell As String, i_intRowCount As Double)
Dim oSheet As Object
Dim oCell As Object
Dim curRow As Integer
Dim curCol As Integer


If Opt = "P" Then
    oSheet = ShPurchase
Else
    oSheet = ShSales
End If
'Range(cell).Select
For curRow = 1 To i_intRowCount - 1
    oCell = oSheet.getCellByPosition(curCol, curRow)
    oCell.String = Trim(oCell.String)
Next curRow
End Sub

Public Function ProCheckLength(opt As String, ColName As String, i_intRowCount As Double, i_lngLen As Long, Cond As String) As Boolean
Dim curCol As Integer
Dim curRow As Integer
Dim oSheet As Object
Dim oCell As Object
If Opt = "P" Then
		oSheet = ShPurchase
Else
		oSheet = ShSales
End If
curCol = GetColIndex(oSheet, ColName)	
For curRow = 1 To i_intRowCount - 1
oCell = oSheet.getCellByPosition(curCol,curRow)
If Cond = "Equal" Then
    If Len(oCell.String) <> i_lngLen Then
         MsgBox "Value [" & oCell.String & "] in cell " & oCell.AbsoluteName & " Should be of Length" & i_lngLen 
         ProCheckLength = False
         SelectObject(oCell)
         Exit Function
    End If
ElseIf Cond = "Less" Then
   If Len(oCell.String) > i_lngLen Then
         MsgBox "Value [" & oCell.String & "] in cell " & oCell.AbsoluteName & " Exceeds Maximum Length" & i_lngLen 
         ProCheckLength = False
         SelectObject(oCell)
         Exit Function
    End If
End If
Next curRow
ProCheckLength = True
End Function

Public Function ProCheckNumeric(opt As String, ColName As String, i_intRowCount As Double) As Boolean
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
For curRow = 1 To i_intRowCount - 1
oCell = oSheet.getCellByPosition(curCol,curRow)
If IsNumeric(oCell.String) = False Or oCell.String = "" Then
     MsgBox " Non Numeric Value [" & oCell.String & "] Found in cell " & oCell.AbsoluteName 
     ProCheckNumeric = False
     SelectObject(oCell)
     Exit Function
ElseIf Trim(oCell.String) = "" Then
     MsgBox " Null Value [" & oCell.String & "] Found in cell " & oCell.AbsoluteName 
     ProCheckNumeric = False
     SelectObject(oCell)
     Exit Function
ElseIf InStr(Trim(oCell.String), ",") > 0 Then
     MsgBox " Non Numeric Value, Remove Comma(,) [" & oCell.String & "] Found in cell " & oCell.AbsoluteName
     ProCheckNumeric = False
     SelectObject(oCell)
     Exit Function    
ElseIf ColName = "C" Then
  If InStr(Trim(oCell.String), ".") > 0 Then
     MsgBox "Non Numeric Value, Remove Dot(.)"
     ProCheckNumeric = False
     SelectObject(oCell)
     Exit Function
  End If      
ElseIf oCell.String < 0 Then
     MsgBox "Value should not be less than zero  [" & oCell.String & "] Found in cell " & oCell.AbsoluteName
     ProCheckNumeric = False
     SelectObject(oCell)
     Exit Function
End If
If CDbl(oCell.String) = "0" Then
   oCell.String = "0"
Else
    If ColName <> "C" Then
     oCell.String = Replace(LTrim(Replace(oCell.String, "0", " ")), " ", "0")
     oCell.String = Round(oCell.String, 2)
    End If
End If
Next curRow
ProCheckNumeric = True
End Function


'This Function is used to Replace < ,' and & which creates error duing uploading
Public Sub ValidateString(Opt As String, ColName As String, i_intRowCount As Integer)
	Dim oSheet as Object
	Dim oCell as Object
	Dim curCol as Integer
	Dim curRow As Integer
	If Opt = "P" Then
		oSheet = ShPurchase
	Else
		oSheet = ShSales
	End If
	curCol = GetColIndex(oSheet, ColName)
	For curRow = 1 To i_intRowCount - 1
		oCell = oSheet.getCellByPosition(curCol,curRow) 
		If InStr(oCell.String, "<") <> 0 Then
			oCell.String = Replace(oCell.String, "<", " ")
		ElseIf InStr(oCell.String, "!") <> 0 Then
        	oCell.String = Replace(oCell.String, "!", " ")
        ElseIf InStr(oCell.String, "@") <> 0 Then
        	oCell.String = Replace(oCell.String, "@", " ")
        ElseIf InStr(oCell.String, "#") <> 0 Then
        	oCell.String = Replace(oCell.String, "#", " ")
        ElseIf InStr(oCell.String, "$") <> 0 Then
       	 	oCell.String = Replace(oCell.String, "$", " ")
        ElseIf InStr(oCell.String, "%") <> 0 Then
       		oCell.String = Replace(oCell.String, "%", " ")
       	ElseIf InStr(oCell.String, "^") <> 0 Then
       		oCell.String = Replace(oCell.String, "^", " ")
       	ElseIf InStr(oCell.String, "*") <> 0 Then
            oCell.String = Replace(oCell.String, "*", " ")
   	 	ElseIf InStr(oCell.String, "(") <> 0 Then
       		oCell.String = Replace(oCell.String, "(", " ")
    	ElseIf InStr(oCell.String, ")") <> 0 Then
        	oCell.String = Replace(oCell.String, ")", " ")
    	ElseIf InStr(oCell.String, "_") <> 0 Then
        	oCell.String = Replace(oCell.String, "_", " ")
    	ElseIf InStr(oCell.String, "-") <> 0 Then
        	oCell.String = Replace(oCell.String, "-", " ")
    	ElseIf InStr(oCell.String, "=") <> 0 Then
        	oCell.String = Replace(oCell.String, "=", " ")
    	ElseIf InStr(oCell.String, "+") <> 0 Then
        	oCell.String = Replace(oCell.String, "+", " ")
       	ElseIf InStr(oCell.String, "\") <> 0 Then
        	oCell.String = Replace(oCell.String, "\", " ")
    	ElseIf InStr(oCell.String, "{") <> 0 Then
        	oCell.String = Replace(oCell.String, "{", " ")
    	ElseIf InStr(oCell.String, "}") <> 0 Then
        	oCell.String = Replace(oCell.String, "}", " ")
    	ElseIf InStr(oCell.String, "[") <> 0 Then
        	oCell.String = Replace(oCell.String, "[", " ")
    	ElseIf InStr(oCell.String, "]") <> 0 Then
        	oCell.String = Replace(oCell.String, "]", " ")
    	ElseIf InStr(oCell.String, ":") <> 0 Then
        	oCell.String = Replace(oCell.String, ":", " ")
    	ElseIf InStr(oCell.String, ";") <> 0 Then
        	oCell.String = Replace(oCell.String, ";", " ")
       	ElseIf InStr(oCell.String, ",") <> 0 Then
        	oCell.String = Replace(oCell.String, ",", " ")
    	ElseIf InStr(oCell.String, ".") <> 0 Then
        	oCell.String = Replace(oCell.String, ".", " ")
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
		ElseIf InStr(oCell.String, "/") <> 0 Then
			oCell.String = Replace(oCell.String, "/", " ")	
		End If
		oCell.String = Replace(Replace(oCell.String, vbLf, " "), vbCr, " ")
		oCell.String = Replace(oCell.String, " ", "")
	Next curRow
End Sub





'Added------------------------------------------------------------------------------------------------------------------
Public Sub ValidateString1(Opt As String, ColName As String, i_intRowCount As Double)
Dim oSheet As Object
Dim oCell As Object
Dim curRow As Integer
Dim curCol As Integer

If Opt = "P" Then
    oSheet = ShPurchase
Else
    oSheet = ShSales
End If
curCol = GetColIndex(oSheet, ColName)
For curRow = 1 To i_intRowCount - 1
    oCell = oSheet.getCellByPosition(curCol, curRow)
    If InStr(oCell.String, "!") <> 0 Then
        oCell.String = Replace(oCell.String, "!", " ")
    ElseIf InStr(oCell.String, "#") <> 0 Then
        oCell.String = Replace(oCell.String, "#", " ")
    ElseIf InStr(oCell.String, "$") <> 0 Then
        oCell.String = Replace(oCell.String, "$", " ")
    ElseIf InStr(oCell.String, "%") <> 0 Then
        oCell.String = Replace(oCell.String, "%", " ")
    ElseIf InStr(oCell.String, "^") <> 0 Then
        oCell.String = Replace(oCell.String, "^", " ")
    ElseIf InStr(oCell.String, "*") <> 0 Then
        oCell.String = Replace(oCell.String, "*", " ")
    ElseIf InStr(oCell.String, "=") <> 0 Then
        oCell.String = Replace(oCell.String, "=", " ")
    ElseIf InStr(oCell.String, "+") <> 0 Then
        oCell.String = Replace(oCell.String, "+", " ")
    ElseIf InStr(oCell.String, "|") <> 0 Then
        oCell.String = Replace(oCell.String, "|", " ")
    ElseIf InStr(oCell.String, "\") <> 0 Then
        oCell.String = Replace(oCell.String, "\", " ")
    ElseIf InStr(oCell.String, "{") <> 0 Then
        oCell.String = Replace(oCell.String, "{", " ")
    ElseIf InStr(oCell.String, "}") <> 0 Then
        oCell.String = Replace(oCell.String, "}", " ")
    ElseIf InStr(oCell.String, "[") <> 0 Then
        oCell.String = Replace(oCell.String, "[", " ")
    ElseIf InStr(oCell.String, "]") <> 0 Then
        oCell.String = Replace(oCell.String, "]", " ")
    ElseIf InStr(oCell.String, "'") <> 0 Then
        oCell.String = Replace(oCell.String, "'", " ")
    ElseIf InStr(oCell.String, """") <> 0 Then
        oCell.String = Replace(oCell.String, """", " ")
    ElseIf InStr(oCell.String, ",") <> 0 Then
        oCell.String = Replace(oCell.String, ",", " ")
    ElseIf InStr(oCell.String, "<") <> 0 Then
        oCell.String = Replace(oCell.String, "<", " ")
    End If
    oCell.String = Replace(Replace(oCell.String, vbLf, " "), vbCr, " ")
Next curRow
End Sub
'------------------------------------------------------------------------------------------------------------------



Sub proValidate_Purchase()
	Dim oRange as Object
	Dim l_intRowCount as Integer
	oRange = FindUsedRange(ShPurchase)
	l_intRowCount = oRange.Rows.Count
	If l_intRowCount = 1 Then
	    MsgBox "No Data Entered", vbInformation + vbOKOnly
	    Exit Sub
	End If
	If ProCheckDuplicate("P", l_intRowCount) = False Then
    Exit Sub
    End If
    ValidateString "P", "A", l_intRowCount
    RemoveSpace "P", "A", l_intRowCount 
    If ProCheckMandatory("P", "A", l_intRowCount) = False Then
		Exit Sub
	End If
	If ProCheckLength("P", "A", l_intRowCount, 25, "Less") = False Then
    	Exit Sub
	End If
'ProRound "P", "B", l_intRowCount
	If ProCheckDate("P", "B", l_intRowCount) = False Then
    	Exit Sub
	End If
	RemoveSpace "P", "C", l_intRowCount
	If ProCheckNumeric("P", "C", l_intRowCount) = False Then
   		Exit Sub
	End If
	If ProCheckLength("P", "C", l_intRowCount, 11, "Equal") = False Then
    	Exit Sub
	End If
    ValidateString1 "P", "D", l_intRowCount
    If ProCheckURMandatory("P", "D", l_intRowCount) = False Then
    	Exit Sub
	End If
	If ProCheckLength("P", "D", l_intRowCount, 150, "Less") = False Then
   		Exit Sub
	End If
	ValidateString1 "P", "E", l_intRowCount
	If ProCheckURMandatory("P", "E", l_intRowCount) = False Then
    	Exit Sub
	End If
	If ProCheckLength("P", "E", l_intRowCount, 200, "Less") = False Then
    	Exit Sub
	End If
	RemoveSpace "P", "F", l_intRowCount
	If ProCheckNumeric("P", "F", l_intRowCount) = False Then
    	Exit Sub
	End If
	If ProCheckLength("P", "F", l_intRowCount, 17, "Less") = False Then
   		Exit Sub
	End If
	RemoveSpace "P", "G", l_intRowCount
	If ProCheckNumeric("P", "G", l_intRowCount) = False Then
    	Exit Sub
	End If
	If ProCheckLength("P", "G", l_intRowCount, 17, "Less") = False Then
    	Exit Sub
	End If
	If ProCheckAmount("P", "F", "G", l_intRowCount) = False Then
    	MsgBox "Vat Amount Greater than Value of Goods"
    	Exit Sub
	End If
	RemoveSpace "P", "H", l_intRowCount
	If ProCheckNumeric("P", "H", l_intRowCount) = False Then
    	Exit Sub
	End If
	If ProCheckLength("P", "H", l_intRowCount, 17, "Less") = False Then
    	Exit Sub
	End If
	If ProCheckAmount("P", "F", "H", l_intRowCount) = False Then
    	MsgBox "Cess Amount Greater than Value of Goods"
    	Exit Sub
	End If
    	GenerateTxt "P"
	End Sub
Sub proValidate_Sales()
	Dim oRange as Object
	Dim l_intRowCount as Integer
	Dim oSheet As Object
    Dim oCell As Object
	
	oRange = FindUsedRange(ShSales)
	l_intRowCount = oRange.Rows.Count
	If l_intRowCount = 1 Then
	    MsgBox "No Data Entered", vbInformation + vbOKOnly
	    Exit Sub
	End If
	
	If ProCheckDuplicate("S", l_intRowCount) = False Then
	    Exit Sub
	End If
	ValidateString "S", "A", l_intRowCount
	RemoveSpace "S", "A", l_intRowCount


	If ProCheckMandatory("S", "A", l_intRowCount) = False Then
		Exit Sub
	End If
	If ProCheckLength("S", "A", l_intRowCount, 25, "Less") = False Then
    	Exit Sub
	End If
	If ProCheckDate("S", "B", l_intRowCount) = False Then
    	Exit Sub
	End If
	RemoveSpace "S", "C", l_intRowCount
	If ProCheckNumeric("S", "C", l_intRowCount) = False Then
   		Exit Sub
	End If
	If ProCheckLength("S", "C", l_intRowCount, 11, "Equal") = False Then
    	Exit Sub
	End If
	ValidateString1 "S", "D", l_intRowCount
	If ProCheckURMandatory("S", "D", l_intRowCount) = False Then
	    Exit Sub
	End If
	If ProCheckLength("S", "D", l_intRowCount, 150, "Less") = False Then
	    Exit Sub
	End If
	ValidateString1 "S", "E", l_intRowCount
	If ProCheckURMandatory("S", "E", l_intRowCount) = False Then
	    Exit Sub
	End If
	If ProCheckLength("S", "E", l_intRowCount, 200, "Less") = False Then
	    Exit Sub
	End If
	RemoveSpace "S", "F", l_intRowCount
	If ProCheckNumeric("S", "F", l_intRowCount) = False Then
	    Exit Sub
	End If
	If ProCheckLength("S", "F", l_intRowCount, 17, "Less") = False Then
	    Exit Sub
	End If
	RemoveSpace "S", "G", l_intRowCount
	If ProCheckNumeric("S", "G", l_intRowCount) = False Then
	    Exit Sub
	End If
	If ProCheckLength("S", "G", l_intRowCount, 17, "Less") = False Then
	    Exit Sub
	End If
	If ProCheckAmount("S", "F", "G", l_intRowCount) = False Then
	    MsgBox "Vat Amount Greater than Invoice Amount"
	    Exit Sub
	End If
	RemoveSpace "S", "H", l_intRowCount
	If ProCheckNumeric("S", "H", l_intRowCount) = False Then
	    Exit Sub
	End If
	If ProCheckLength("S", "H", l_intRowCount, 17, "Less") = False Then
	    Exit Sub
	End If
	If ProCheckAmount("S", "F", "H", l_intRowCount) = False Then
	    MsgBox "Cess Amount Greater than Invoice Amount"
	    Exit Sub
	End If
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
	For curRow = 1 To i_intRowCount - 1
		oCell = oSheet.getCellByPosition(curCol, curRow)
		If oCell.String = "" Or Trim(oCell.String) = "0" Then
			MsgBox "Mandatory Information Missing or Value is zero [" & oCell.String & "] in cell " & oCell.AbsoluteName
			SelectObject(oCell)
			ProCheckMandatory = False
			Exit Function
		End If
	Next curRow
	ProCheckMandatory = True
End Function


'Added---------------------------------------------------------------------------------------------------------------
Public Function ProCheckURMandatory(Opt As String,ColName As String, i_intRowCount As Double) As Boolean
Dim oSheet As Object
Dim oCell As Object
Dim curCol As Integer
Dim curRow As Integer

If Opt = "P" Then
    oSheet = ShPurchase
Else
    oSheet = ShSales
End If
For curRow = 1 To i_intRowCount - 1
    curCol = GetColIndex(oSheet, "C")
    oCell = oSheet.getCellByPosition(curCol,curRow)
    If oCell.String = "88888888888" Or oCell.String = "99999999999" Or Mid(oCell.String, 1, 2) <> "32" Then
       	curCol = GetColIndex(oSheet, ColName)
       	oCell = oSheet.getCellByPosition(curCol,curRow)
        If Trim(oCell.String) = "" Then
             MsgBox "Mandatory Information Missing  [" & oCell.String & "] in cell " & oCell.AbsoluteName
             ProCheckURMandatory = False
             SelectObject(oCell)
             Exit Function
        End If
    End If
Next curRow
ProCheckURMandatory = True
End Function


Public Function ProCheckDuplicate(Opt As String,i_intRowCount As Double) As Boolean
Dim l_dblURCount As Double
Dim l_dblISCount As Double
Dim curCol As Integer
Dim oSheet As Object
Dim oCell As Object
Dim curRow As Integer
dim ColName As String
ColName = "C"	
l_dblURCount = 0
l_dblISCount = 0
If Opt = "P" Then
    oSheet = ShPurchase
Else
    oSheet = ShSales
End If
curCol = GetColIndex(oSheet, ColName)
For curRow = 2 To i_intRowCount
    oCell = oSheet.getCellByPosition(curCol, curRow)
    If Trim(oCell.String) = "99999999999" Then
         l_dblURCount = l_dblURCount + 1
    ElseIf Trim(oCell.String) = "88888888888" Then
         l_dblISCount = l_dblISCount + 1
    End If
Next curRow
If l_dblURCount > 1 Then
   MsgBox "Only One Entry Should be There for the Unregistered Entry "
   ProCheckDuplicate = False
   SelectObject(oCell)
   Exit Function
End If
If l_dblISCount > 1 Then
   MsgBox "Only One Entry  Should be There for the InsterState and Export Entry"
   ProCheckDuplicate = False
   SelectObject(oCell)
   Exit Function
End If
ProCheckDuplicate = True
End Function

'----------------------------------------------------------------------------------------------------------------------



Public Function ValidateDate(i_date As String) As Boolean
	If Mid(i_date, 3, 1) <> "-" Then
        ValidateDate = False
        Exit Function
    End If
    If Mid(i_date, 6, 1) <> "-" Then
        ValidateDate = False
        Exit Function
    End If
	If (Val(Mid(i_date, 4, 2)) < 1 Or Val(Mid(i_date, 4, 2)) > 12) Then
		ValidateDate = False
		Exit Function
	End If
	

'Added-----------------------------------------------------------------------------------------------------------	
	If (Val(Mid(i_date, 4, 2)) = 1 Or Val(Mid(i_date, 4, 2)) = 3 Or Val(Mid(i_date, 4, 2)) = 5 Or Val(Mid(i_date, 4, 2)) = 7 Or Val(Mid(i_date, 4, 2)) = 8 Or Val(Mid(i_date, 4, 2)) = 10 Or Val(Mid(i_date, 4, 2)) = 12) Then
        If (Val(Mid(i_date, 1, 2)) > 31) Then
            ValidateDate = False
            Exit Function
        End If
    End If
    If (Val(Mid(i_date, 4, 2)) = 4 Or Val(Mid(i_date, 4, 2)) = 6 Or Val(Mid(i_date, 4, 2)) = 9 Or Val(Mid(i_date, 4, 2)) = 11) Then
        If (Val(Mid(i_date, 1, 2)) > 30) Then
            ValidateDate = False
            Exit Function
        End If
    End If
     If (Val(Mid(i_date, 4, 2)) = 2) Then
        If (Val(Mid(i_date, 1, 2)) > 29) Then
            ValidateDate = False
            Exit Function
        End If
    End If
    If (Val(Mid(i_date, 4, 2)) < 1 Or Val(Mid(i_date, 4, 2)) > 12) Then
        ValidateDate = False
        Exit Function
    End If
'------------------------------------------------------------------------------------------------------------    
	
	
	If Val(Mid(i_date, 7, 4)) < 1900 Then
		ValidateDate = False
		Exit Function
	End If

      
	ValidateDate = True
End Function

