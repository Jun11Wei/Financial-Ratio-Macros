VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Debt1 
   Caption         =   "Debt Ratios"
   ClientHeight    =   3640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5020
   OleObjectBlob   =   "Debt1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Debt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

Range(Selection.Address) = "Debt Ratio:"
Range(Selection.Address).Offset(0, 1).Value = Liability.Value / Asset.Value
Range(Selection.Address).Offset(1, 0).Value = "Interest Coverage Ratio:"
Range(Selection.Address).Offset(1, 1).Value = EBIT.Value / Interest.Value
Range(Selection.Address).EntireColumn.AutoFit
    Range(Selection.Address, Range(Selection.Address).End(xlToRight).End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

Debt1.Hide

End Sub

Private Sub CommandButton2_Click()

Debt1.Hide

End Sub

