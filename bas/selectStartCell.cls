VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "selectStartCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iProcessOneSheet

Private Sub iProcessOneSheet_processoneSheet(sheet As Worksheet)
        sheet.Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        sheet.Cells(1, 1).Select
End Sub
