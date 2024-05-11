Attribute VB_Name = "ProcessAllSheets"
Option Explicit

''' 全てのシートの先頭のCellを選択状態にする
Public Sub selectAllStartCell()
Attribute selectAllStartCell.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim proc As selectStartCell
    Set proc = New selectStartCell
    
        processAllSheets proc
    
    Set proc = Nothing
    
End Sub


''' 全シートのフォントをメイリオに変更する
Public Sub changeFontMeiryo()

    Dim proc As setMeiryo
    Set proc = New setMeiryo
    
        processAllSheets proc
    
    Set proc = Nothing
    Application.ScreenUpdating = True

End Sub


''' 全SheetをLoopする。
''' 但し、Sheet名の先頭が#で始まる場合は除外する。
Private Sub processAllSheets(processoneSheet As iProcessOneSheet)
On Error GoTo error_processAllSheets
    Application.ScreenUpdating = False

    Dim sheet As Worksheet
    Dim i As Long
    For i = ActiveWorkbook.Sheets.Count To 1 Step -1
        Set sheet = ActiveWorkbook.Sheets(i)
        
        If Left(sheet.Name, 1) <> "#" Then
            sheet.Select
            processoneSheet.processoneSheet sheet
        End If
        
    Next
    Set sheet = Nothing


error_processAllSheets:
    Application.ScreenUpdating = True
    
End Sub
