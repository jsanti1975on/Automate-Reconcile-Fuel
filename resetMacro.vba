Option Explicit
Sub resetSheet2()
'
' resetSheet2 Macro

    Range("E2").Select
    ActiveWindow.SmallScroll Down:=24
    Range("E2:F56").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G73").Select
End Sub

Sub resetSheet1Rev()
'Recorded Macro for prepping sheet after use

    Range("R2").Select
    ActiveCell.FormulaR1C1 = ""
    Selection.AutoFill Destination:=Range("R2:V2"), Type:=xlFillDefault
    Range("R2:V2").Select
    ActiveWindow.SmallScroll Down:=0
    Selection.AutoFill Destination:=Range("R2:V56"), Type:=xlFillDefault
    Range("R2:V56").Select
    ActiveWindow.SmallScroll Down:=-15
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    Range("G2").Select
    ActiveCell.FormulaR1C1 = ""
    Selection.AutoFill Destination:=Range("G2:H2"), Type:=xlFillDefault
    Range("G2:H2").Select
    Selection.AutoFill Destination:=Range("G2:H56"), Type:=xlFillDefault
    Range("G2:H56").Select
    ActiveWindow.SmallScroll Down:=-18
    Range("G2").Select
    
End Sub
