Attribute VB_Name = "Text_Converter_Module"
Sub convert_text()
    'Declare Current as a worksheet object variable
    Dim Current As Worksheet
    'Declare type as type of value in cell
    Dim TypeCheck
    'Declare Equation as cell's formula
    Dim Equation
   
    'Loop through all of the worksheets in the active workbook.
    For Each Current In Worksheets
        'Set current worksheet to be active
        Current.Activate
        'Iterate throuch each used cell
        For Each Cell In ActiveSheet.UsedRange.Cells
            'Check type of value in each cell
            TypeCheck = VarType(Cell.Value)
            If Not (Cell.Font.Color = vbWhite) Then
                If Cell.HasFormula = True Then
                    'Retrieve formula from cell
                    Equation = Cell.Formula
                    'Check to see if cell contains linked number
                    If InStr(Equation, "!") Then
                        Cell.Font.Color = RGB(0, 153, 0)
                    'Check to see if cell contains formula
                    Else
                        Cell.Font.Color = RGB(0, 0, 0)
                    End If
                ElseIf TypeCheck = 5 Then
                    'Else if number is hardcoded, make font color blue
                    Cell.Font.Color = RGB(0, 0, 225)
                ElseIf TypeCheck = 6 Then
                    Cell.Font.Color = RGB(0, 0, 225)
                End If
            End If
        Next
    Next
    MsgBox "Font colors changed."
End Sub
