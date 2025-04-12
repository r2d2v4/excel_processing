VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub LoadFiles()
    Dim files_names() As String
    folder_in = Cells(2, 4)
    OK_folder = Cells(3, 4)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folder_in)
    ''' Filename = Dir(Folders_Com(i) & "*.xls")
    
    files_count = folder.Files.Count
    If files_count = 0 Then
        MsgBox ("В папке " & folder_in & " нет excel файлов.")
        Exit Sub
    End If
    
    'ReDim files_names(files_count)
    
    i = 0
    For Each file In folder.Files
        ''' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        ''' Попытка отследить повторяющиеся файлы
        ''' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'        If Left(file.Name, 3) <> "arp" Then
'            fileID = Left(file.Name, 36)
'        Else
'            fileID = Right(Left(file.Name, 50), 36)
'        End If
'
'        row_empty = 10
'        flag = False
'        While ThisWorkbook.Worksheets("files").Cells(row_empty, 2) <> ""
'            row_empty = row_empty + 1
'            If ThisWorkbook.Worksheets("files").Cells(row_empty, 4) = fileID Then
'                flag = True
'                MsgBox ("Файл с ID " & fileID & " уже обрабатывался ранее.")
'            End If
'        Wend
        ''' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        ''' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        
        row_empty = 10
        While ThisWorkbook.Worksheets("files").Cells(row_empty, 2) <> ""
            row_empty = row_empty + 1
        Wend
        
        ThisWorkbook.Worksheets("files").Cells(row_empty, 2) = Format(Now, "yyyy-MM-dd")
        ThisWorkbook.Worksheets("files").Cells(row_empty, 3) = Format(Now, "HH:mm:ss")
        
        If Left(file.Name, 3) <> "arp" Then
            fileID = Left(file.Name, 36)
            ThisWorkbook.Worksheets("files").Cells(row_empty, 4) = Left(file.Name, 36)
            ThisWorkbook.Worksheets("files").Cells(row_empty, 4).HorizontalAlignment = xlLeft
            ThisWorkbook.Worksheets("files").Cells(row_empty, 4).VerticalAlignment = xlBottom
        Else
            fileID = Right(Left(file.Name, 50), 36)
            ThisWorkbook.Worksheets("files").Cells(row_empty, 4) = Right(Left(file.Name, 50), 36)
            ThisWorkbook.Worksheets("files").Cells(row_empty, 4).HorizontalAlignment = xlLeft
            ThisWorkbook.Worksheets("files").Cells(row_empty, 4).VerticalAlignment = xlBottom
        End If
        ThisWorkbook.Worksheets("files").Cells(row_empty, 5) = file.Name
        ThisWorkbook.Worksheets("files").Cells(row_empty, 5).HorizontalAlignment = xlLeft
        ThisWorkbook.Worksheets("files").Cells(row_empty, 5).VerticalAlignment = xlBottom
        
        Set file_in = Workbooks.Open(file, ReadOnly:=False)
        
        r = 16
        While file_in.Sheets(1).Cells(r, 11) <> "Всего"
            r = r + 1
        Wend
        
        Set ws = file_in.Worksheets(1)
        ws.Range(ws.Cells(16, 1), ws.Cells(r - 1, 14)).Copy
        date_report = ws.Range("M5:N5").Value
        
        row_empty = 2
        While ThisWorkbook.Worksheets("data").Cells(row_empty, 3) <> ""
            row_empty = row_empty + 1
        Wend
        '''''MsgBox ("Номер следующей пустой строки: " & row_empty)
                
        ThisWorkbook.Worksheets("data").Cells(row_empty, 2).PasteSpecial xlPasteValues
        
        For r_date = row_empty To row_empty + r - 17 Step 1
            ThisWorkbook.Worksheets("data").Cells(r_date, 1) = date_report
            ThisWorkbook.Worksheets("data").Cells(r_date, 2).NumberFormat = "m/d/yyyy"
            ThisWorkbook.Worksheets("data").Cells(r_date, 10).NumberFormat = "m/d/yyyy"
        Next
        
        Application.CutCopyMode = False
        
        file_in.Close SaveChanges:=False
        
        'MsgBox (Format(Now, "yyyy-MM-dd"))
        ok_subfolder = OK_folder & Format(Now, "yyyy-MM-dd") & "\"
        'MsgBox (ok_subfolder)
        If fso.FolderExists(ok_subfolder) Then
            'Завершаем программу, если произойдет ошибка
            'On Error Resume Next
            fso.MoveFile file, ok_subfolder
        Else
            MkDir ok_subfolder
            'Завершаем программу, если произойдет ошибка
            'On Error Resume Next
            fso.MoveFile file, ok_subfolder
        End If
        
        i = i + 1

    Next

ThisWorkbook.Worksheets("control").Activate
MsgBox ("Обработано " & i & " файлов.")
End Sub
