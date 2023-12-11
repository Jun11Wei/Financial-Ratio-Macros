VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FinancialRatios 
   Caption         =   "Financial Ratios"
   ClientHeight    =   5440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9040.001
   OleObjectBlob   =   "FinancialRatios.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FinancialRatios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Debt_Click()

    FinancialRatios.Hide
    Debt1.Show

End Sub

Private Sub Liquidity_Click()

FinancialRatios.Hide
Liquidity1.Show

End Sub

Private Sub Profitability_Click()

FinancialRatios.Hide
Profitability1.Show

End Sub

Private Sub Quit_Click()

FinancialRatios.Hide

End Sub
