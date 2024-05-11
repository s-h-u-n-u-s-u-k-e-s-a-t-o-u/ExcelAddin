Attribute VB_Name = "ListUpSheets"
Option Explicit

''' シートの一覧を先頭に追加する
Public Sub listUpSheets()
On Error GoTo Error_LisUpSheets

    Application.ScreenUpdating = False
    
    Dim rows As Long
    rows = ActiveWorkbook.Sheets.Count
        
    '貼り付ける際に2次元配列であることが条件
    Dim range As Variant
    ReDim range(1 To rows, 1 To 2)
    
    Dim sheet As Worksheet
    Dim sheetCount As Long
    sheetCount = 1
    
    Dim i As Long
    For i = 1 To rows
        Set sheet = ActiveWorkbook.Sheets(i)
                        
        If Left(sheet.Name, 1) <> "#" Then
            range(sheetCount, 1) = sheetCount
            range(sheetCount, 2) = sheet.Name
            sheetCount = sheetCount + 1
        End If
    Next
    Set sheet = Nothing
    
    Dim newName As String
    newName = "#SheetList"
    'シート存在チェック
    Dim j As Long
    j = 0
    Dim tempSheet As Worksheet
    Do While True
        On Error Resume Next
        Set tempSheet = ActiveWorkbook.Sheets(newName)
           If Err.Number <> 0 Then
               Exit Do
           End If
        On Error GoTo Error_LisUpSheets
        j = j + 1
        newName = "#SheetList_" & j
    Loop
    On Error GoTo Error_LisUpSheets
    
    Set tempSheet = ActiveWorkbook.Sheets.Add(ActiveWorkbook.Sheets(1))
    tempSheet.Name = newName
    tempSheet.Cells(1, 1) = "No."
    tempSheet.Cells(1, 2) = "Sheet"
    
    Dim destRange As range
    Set destRange = tempSheet.range("A2").Resize(rows, 2)
    destRange = range
    
   
   'Book内のリンクを設定する
    Dim k As Long
    For k = 1 To destRange.Count
        If destRange(k, 2) <> "" Then
            tempSheet.Hyperlinks.Add Anchor:=destRange(k, 2), Address:="", SubAddress:=destRange(k, 2).Value & "!A1", TextToDisplay:=destRange(k, 2).Value
        End If
    Next
    
    Set destRange = Nothing
    
    '罫線を引く
    Dim borderRnd As range
    Set borderRnd = tempSheet.Cells(1, 1).CurrentRegion
    
    With borderRnd.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With borderRnd.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With borderRnd.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With borderRnd.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With borderRnd.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With borderRnd.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    
    '見出しの下線
    Set borderRnd = tempSheet.range("A1:B1")
    With borderRnd.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Set borderRnd = Nothing

    tempSheet.Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    tempSheet.Cells(1, 1).Select
    Cells.EntireColumn.AutoFit
        
    Set tempSheet = Nothing
        
Error_LisUpSheets:
    Application.ScreenUpdating = True

End Sub
