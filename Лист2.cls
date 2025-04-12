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
    'folder_in = "c:\tmp\xls"
    
    'Set f = fso.CreateTextFile("C:\TEMP\testfile1.txt")
    'Set file_in = Workbooks.Open("c:\tmp\xls\0195582d-9906-000c-1565-dc5960a5b90b_0531480.xlsx", ReadOnly:=True)
    'file_in.Close SaveChanges:=False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folder_in)
    ''' Filename = Dir(Folders_Com(i) & "*.xls")
    
    files_count = folder.Files.Count - 1
    If files_count = -1 Then
        MsgBox ("В папке " & folder_in & " нет excel файлов")
        Exit Sub
    End If
    
    ReDim files_names(files_count)
    
    i = 0
    For Each file In folder.Files
        
        row_empty = 10
        While ThisWorkbook.Worksheets("control").Cells(row_empty, 2) <> ""
            row_empty = row_empty + 1
        Wend
        ThisWorkbook.Worksheets("control").Cells(row_empty, 2) = Format(Now, "yyyy-MM-dd")
        ThisWorkbook.Worksheets("control").Cells(row_empty, 3) = Format(Now, "HH:mm:ss")
        If Left(file.Name, 3) <> "arp" Then
            ThisWorkbook.Worksheets("control").Cells(row_empty, 4) = Left(file.Name, 36)
            ThisWorkbook.Worksheets("control").Cells(row_empty, 4).HorizontalAlignment = xlLeft
            ThisWorkbook.Worksheets("control").Cells(row_empty, 4).VerticalAlignment = xlBottom
        Else
            ThisWorkbook.Worksheets("control").Cells(row_empty, 4) = Right(Left(file.Name, 50), 36)
            ThisWorkbook.Worksheets("control").Cells(row_empty, 4).HorizontalAlignment = xlLeft
            ThisWorkbook.Worksheets("control").Cells(row_empty, 4).VerticalAlignment = xlBottom
        End If
        ThisWorkbook.Worksheets("control").Cells(row_empty, 5) = file.Name
        ThisWorkbook.Worksheets("control").Cells(row_empty, 5).HorizontalAlignment = xlLeft
        ThisWorkbook.Worksheets("control").Cells(row_empty, 5).VerticalAlignment = xlBottom
        'ThisWorkbook.Worksheets("control").Cells(row_empty, 3) = Left(file.Name, 12)
        
        
        'MsgBox (file.Name + file.Path)
        'files_names(i) = file.Path
        'i = i + 1
    'Next
    
    'i = 0
    'For Each file_name In files_names
        '''''MsgBox ("Обрабатываем файл: " & files_names(i))
        i = i + 1
        
        Set file_in = Workbooks.Open(file, ReadOnly:=False)
        'file_in.Sheets(1).Activate
        
        r = 16
        While file_in.Sheets(1).Cells(r, 11) <> "Всего"
            'For c = 1 To 14 Step 1
            '    'a = file_in.Cells(r, c)
            '    MsgBox (file_in.Cells(r, c))
            'Next
            'MsgBox (file_in.Sheets(1).Cells(r, 11))
            r = r + 1
        Wend
        'MsgBox (r - 1)
        
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
        
        
    Next

ThisWorkbook.Worksheets("control").Activate
MsgBox ("Обработано " & i & " файлов.")
End Sub


