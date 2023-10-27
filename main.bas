Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Public Function ERR(expression)
'ERR = DERIVF(cell, "A2", 0)

Dim xRetList As Object
Dim xRegEx As Object
Dim I As Long
Dim xRet As String
Dim sum As Double
Dim abl As Double
Dim Var As String
Dim pm As String

Application.Volatile
Set xRegEx = CreateObject("VBSCRIPT.REGEXP")
With xRegEx
    .Pattern = "('?[a-zA-Z0-9\s\[\]\.]{1,99})?'?!?\$?[A-Z]{1,3}\$?[0-9]{1,7}(:\$?[A-Z]{1,3}\$?[0-9]{1,7})?"
    .Global = True
    .MultiLine = True
    .IgnoreCase = False
End With
sum = 0
Set xRetList = xRegEx.Execute(expression.Formula)
For I = 0 To xRetList.Count - 1
    Var = Application.ConvertFormula(Formula:=xRetList.item(I), FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)
    abl = Abs(DERIVF(expression, Range(Var)))
    pm = Range(Var).Offset(0, 1)
    sum = sum + abl * pm
Next
ERR = sum

End Function


Public Function DERIVF(expression, variable) As Double
'Custom function to return the first derivative of a formula in a cell.

Dim OldX As Double, OldY As Double, NewX As Double, NewY As Double
Dim FormulaString As String, XAddress As String
FormulaString = expression.Formula
OldY = expression.Value
XAddress = variable.Address 'Default is absolute reference
OldX = variable.Value
NewX = OldX * 1.00000001

FormulaString = Application.ConvertFormula(Formula:=FormulaString, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)
FormulaString = Application.Substitute(FormulaString, XAddress, NewX)
NewY = Evaluate(FormulaString)
DERIVF = (NewY - OldY) / (NewX - OldX)

End Function


Private Function ExtractCellRefs(Rg As Range) As String
'Updateby Extendoffice
    Dim xRetList As Object
    Dim xRegEx As Object
    Dim I As Long
    Dim xRet As String
    Application.Volatile
    Set xRegEx = CreateObject("VBSCRIPT.REGEXP")
    With xRegEx
        .Pattern = "('?[a-zA-Z0-9\s\[\]\.]{1,99})?'?!?\$?[A-Z]{1,3}\$?[0-9]{1,7}(:\$?[A-Z]{1,3}\$?[0-9]{1,7})?"
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With
    Set xRetList = xRegEx.Execute(Rg.Formula)
    If xRetList.Count > 0 Then
        For I = 0 To xRetList.Count - 1
            xRet = xRet & xRetList.item(I) & ", "
        Next
        ExtractCellRefs = Left(xRet, Len(xRet) - 2)
    Else
        ExtractCellRefs = "No Matches"
    End If
End Function

