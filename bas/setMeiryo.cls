VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "setMeiryo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iProcessOneSheet

Private Sub iProcessOneSheet_processoneSheet(sheet As Worksheet)
        sheet.Select
        With sheet.Cells.Font
            .Name = "ÉÅÉCÉäÉI"
            .Size = 16
        End With
        Cells.EntireColumn.AutoFit
End Sub

