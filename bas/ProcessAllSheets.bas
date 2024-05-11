Attribute VB_Name = "ProcessAllSheets"
Option Explicit

''' �S�ẴV�[�g�̐擪��Cell��I����Ԃɂ���
Public Sub selectAllStartCell()
Attribute selectAllStartCell.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim proc As selectStartCell
    Set proc = New selectStartCell
    
        processAllSheets proc
    
    Set proc = Nothing
    
End Sub


''' �S�V�[�g�̃t�H���g�����C���I�ɕύX����
Public Sub changeFontMeiryo()

    Dim proc As setMeiryo
    Set proc = New setMeiryo
    
        processAllSheets proc
    
    Set proc = Nothing
    Application.ScreenUpdating = True

End Sub


''' �SSheet��Loop����B
''' �A���ASheet���̐擪��#�Ŏn�܂�ꍇ�͏��O����B
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
