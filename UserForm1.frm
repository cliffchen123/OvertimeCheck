VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�[�Z�O�ˬd"
   ClientHeight    =   7695
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   17130
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public dayAmounts As Integer
Public jobTitles As Object
Function setJobTitles() As Scripting.Dictionary

    Dim jobTitles
    Set jobTitles = CreateObject("Scripting.Dictionary")
    jobTitles.Add "�p��", 1
    jobTitles.Add "�Ƹp��", 2
    jobTitles.Add "�D������", 3
    jobTitles.Add "�ժ�", 4
    jobTitles.Add "�Ʋժ�", 5
    jobTitles.Add "�D��", 6
    jobTitles.Add "�M���e��", 7
    jobTitles.Add "���", 8
    jobTitles.Add "�G�F�M���ݬ��", 9
    jobTitles.Add "����", 10
    jobTitles.Add "�G�F�M��", 11
    jobTitles.Add "���R�v", 12
    jobTitles.Add "����", 13
    jobTitles.Add "�G�F�x", 14
    jobTitles.Add "�]�p�v", 15
    jobTitles.Add "�M��", 16
    jobTitles.Add "���", 17
    jobTitles.Add "�U�z��", 18
    jobTitles.Add "�U�z�]�p�v", 19
    jobTitles.Add "���z��", 20
    jobTitles.Add "�ѰO", 21
    jobTitles.Add "�����H��", 22
    jobTitles.Add "�ոp��ƥD��", 23
    jobTitles.Add "�ոp��ƪѪ�", 24
    jobTitles.Add "�ոp��ƱM��", 25
    jobTitles.Add "�ոp��Ƭ��", 26
    jobTitles.Add "�ոp��ƽҭ�", 27
    jobTitles.Add "�˹�x", 28
    Set setJobTitles = jobTitles
    
End Function


Sub PrintFinishTime()
    ProcessingBox.Value = ProcessingBox.Value & "----------------------------�����ɶ�:" & Now() & "----------------------------" & vbCrLf
End Sub

Function nb_days_month()
    
    
    'Month / Year of the date
    var_month = MonthText.Value
    var_year = YearText.Value
    
    'Calculation for the first day of the following month
    date_next_month = DateSerial(var_year, var_month + 1, 1)
    
    'Date of the last day
    last_day_month = date_next_month - 1
    
    'Number for the last day of month (= last day)
    nb_days = Day(last_day_month)
    
    nb_days_month = nb_days
    
End Function

Function ReadSalaryFromSalary(book As Workbook)
    Dim salary
    Set salary = CreateObject("Scripting.Dictionary")
    Set sh = book.Sheets(1)
    Dim rw As Integer
    For rw = 2 To sh.UsedRange.Rows.Count
        If salary.Exists(Trim(sh.Cells(rw, 1))) Then
            salary.Item(Trim(sh.Cells(rw, 1))) = CLng(sh.Cells(rw, 2))
        Else
            salary.Add Trim(sh.Cells(rw, 1)), CLng(sh.Cells(rw, 2))
        End If
    Next
    
    Set ReadSalaryFromSalary = salary
End Function

Function ReadSalaryFromTreatment(book As Workbook, salary As Object)
    Set sh = book.Sheets(1)
    Dim rw As Integer
    For rw = 4 To sh.UsedRange.Rows.Count
        If salary.Exists(Trim(sh.Cells(rw, 4))) Then
            salary.Item(Trim(sh.Cells(rw, 4))) = CLng(sh.Cells(rw, 26))
        Else
            salary.Add Trim(sh.Cells(rw, 4)), CLng(sh.Cells(rw, 26))
        End If
    Next
    
    Set ReadSalaryFromTreatment = salary
End Function

Function ReadOvertime(book As Workbook, isProject As Boolean, project As Object)
    Dim sh As Excel.Worksheet
    Set sh = book.Sheets(1)
    Dim people
    Set people = CreateObject("Scripting.Dictionary")
    Dim person As New Collection
    
    'Read data from overtime file
    Dim readyToEnd As Integer
    Dim rw As Integer
    For rw = 1 To sh.UsedRange.Rows.Count
        If jobTitles.Exists(sh.Cells(rw, 1).Value) Then 'read people data
            person.Add sh.Cells(rw, 1).Value
            person.Add sh.Cells(rw, 2).Value
            Dim monthEnd As Integer
            monthEnd = 3 + dayAmounts - 1
            person.Add sh.Range(sh.Cells(rw, 3), sh.Cells(rw, monthEnd)).Value
            person.Add sh.Cells(rw, monthEnd + 1).Value
            person.Add sh.Cells(rw, monthEnd + 2).Value
            person.Add sh.Cells(rw, monthEnd + 3).Value
            person.Add sh.Cells(rw, monthEnd + 4).Value
            person.Add sh.Cells(rw, monthEnd + 7).Value
            If people.Exists(sh.Cells(rw, 1).Value) = False Then
                people.Add sh.Cells(rw, 1).Value, New Collection
            End If
            people.Item(sh.Cells(rw, 1).Value).Add person
            Set person = Nothing

            If isProject Then
                project.Item(sh.Cells(rw, 2).Value) = sh.Cells(rw, monthEnd + 1).Value
            End If

        End If
    Next rw
    
    Set ReadOvertime = people
    
End Function


Function WriteFile(book As Workbook, people As Object, salary As Object, isProject As Boolean, project As Object, resultFolder As String, oriFile As String)
    Dim fileName As String
    fileName = Right(oriFile, Len(oriFile) - InStrRev(oriFile, "\"))
    
    Dim sh As Excel.Worksheet
    Set sh = book.Sheets(1)
    Dim total_h As Integer
    Dim sign_h As Integer
    Dim people_h As Integer
    Dim i As Integer
    
    'Record row style of total, sign, people
    sh.Range("A4:AS4").Copy Destination:=sh.Range("A10000:AS10000")
    sh.Range("A10000:AS10000").ClearContents
    people_h = sh.Rows(4).RowHeight
    For Each rw In sh.Rows
        If sh.Cells(rw.Row, 1).Value = "�X�p" Then
            sh.Range("A" & CStr(rw.Row) + ":AS" & CStr(rw.Row)).Copy Destination:=sh.Range("A10001:AS10001")
            total_h = sh.Rows(rw.Row).RowHeight
        ElseIf sh.Cells(rw.Row, 1).Value = "���ݪ��x" Then
            sh.Range("A" & CStr(rw.Row) + ":AS" & CStr(rw.Row)).Copy Destination:=sh.Range("A10002:AS10002")
            sign_h = sh.Rows(rw.Row).RowHeight
            Exit For
        End If
    Next
    
    'List all people
    Dim monthEnd As Integer
    monthEnd = 3 + dayAmounts - 1
    rw = 4
    For Each jt In jobTitles
        If people.Exists(jt) Then
            For Each Item In people.Item(jt)
                sh.Range("A" & CStr(rw) + ":AS" & CStr(rw)).UnMerge
                sh.Range("A" & CStr(rw) + ":AS" & CStr(rw)).Clear
                sh.Range("A10000:AS10000").Copy Destination:=sh.Range("A" & CStr(rw) + ":AS" & CStr(rw))
                sh.Rows(rw).RowHeight = people_h
                sh.Cells(rw, 1).Value = Item(1)
                sh.Cells(rw, 2).Value = Item(2)
                sh.Range(sh.Cells(rw, 3), sh.Cells(rw, monthEnd)).Value = Item(3)
                Dim idx As Integer
                If Not isProject Then
                    For idx = 1 To UBound(Item(3), 2) 'check whether overtime is over >4 or >8
                        Dim wd As Integer
                        wd = Weekday(DateValue(CStr(YearText) + "/" + CStr(MonthText) + "/" + CStr(idx)))
                        If (wd = 1 Or wd = 7) Then 'sunday or saturday
                            If Item(3)(1, idx) > 8 Then
                                sh.Range(sh.Cells(rw, 1), sh.Cells(rw, monthEnd + 7)).Interior.Color = RGB(255, 126, 0)
                                ProcessingBox.Value = ProcessingBox.Value & vbTab & Item(2) & " " & "����[�Z>8�p��" & vbCrLf
                                Exit For
                            End If
                        Else
                            If Item(3)(1, idx) > 4 Then
                                sh.Range(sh.Cells(rw, 1), sh.Cells(rw, monthEnd + 7)).Interior.Color = RGB(255, 255, 0)
                                ProcessingBox.Value = ProcessingBox.Value & vbTab & Item(2) & " " & "����[�Z>4�p��" & vbCrLf
                                Exit For
                            End If
                        End If
                    Next
                End If
                sh.Cells(rw, monthEnd + 1).Value = Item(4)
                
                Dim projectHour As Integer
                If InStr(1, fileName, "�M��") = 0 And project.Exists(Item(2)) Then
                    projectHour = project.Item(Item(2))
                Else
                    projectHour = 0
                End If
                If project.Exists(Item(2)) And Item(4) + projectHour > 70 Then 'check whether overtime is over �@��+�M��>70
                    sh.Range(sh.Cells(rw, 1), sh.Cells(rw, monthEnd + 7)).Interior.Color = RGB(255, 126, 255) 'overtime is over, fill with purple
                    ProcessingBox.Value = ProcessingBox.Value & vbTab & Item(2) & " " & "�@��+�M��>70�p��" & vbCrLf
                End If
                sh.Cells(rw, monthEnd + 7).Value = Item(8)
                If salary.Exists(Item(2)) Then 'check whther exist salary information
                    sh.Cells(rw, monthEnd + 2).Value = salary.Item(Item(2))
                    sh.Cells(rw, monthEnd + 3).Value = Round(CLng(salary.Item(Item(2))) / 240)
                    sh.Cells(rw, monthEnd + 4).Value = sh.Cells(rw, monthEnd + 3).Value * Item(4)
                Else
                    sh.Range(sh.Cells(rw, 1), sh.Cells(rw, monthEnd + 7)).Interior.Color = RGB(255, 0, 0) 'if not exist salary information, fill with red
                    ProcessingBox.Value = ProcessingBox.Value & vbTab & Item(2) & " " & "�~���Ƥ�����" & vbCrLf
                End If
                rw = rw + 1
            Next
        End If
    Next
    sh.Range("A" & CStr(rw) + ":AS" & CStr(rw + 100)).UnMerge
    sh.Range("A" & CStr(rw) + ":AS" & CStr(rw + 100)).Clear
    sh.Range("A10001:AS10001").Copy Destination:=sh.Range("A" & CStr(rw) + ":AS" & CStr(rw))
    sh.Rows(rw).RowHeight = total_h
    sh.Cells(rw, 2).Value = CStr(rw - 4) + "�H"
    sh.Range("A10002:AS10002").Copy Destination:=sh.Range("A" & CStr(rw + 1) + ":AS" & CStr(rw + 1))
    sh.Rows(rw + 1).RowHeight = sign_h
    
    sh.Range("A10000:AS10000").Clear
    sh.Range("A10000:AS10001").Clear
    sh.Range("A10000:AS10002").Clear
    
    book.Application.DisplayAlerts = False
    book.SaveAs resultFolder + fileName + "(����).xls", FileFormat:=56, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    
End Function


Private Sub CheckButton_Click()
    
    'Read file name from TextBox
    Dim FileNames() As String
    FileNames = Split(OvertimeTextBox.Value, vbCrLf)
    
    'Check input data
    ProcessingBox.Value = ProcessingBox.Value & "�ˬd��Ƨ����" & vbCrLf
    If YearText = "" Or MonthText = "" Then
        MsgBox ("�~�����������")
        PrintFinishTime
        Exit Sub
    ElseIf Dir(SalaryButton.Caption) = "" Then
        MsgBox ("�~���Ƥ�����")
        PrintFinishTime
        Exit Sub
    ElseIf Dir(TreatmentButton.Caption) = "" Then
        MsgBox ("�ݹJ�M�U������")
        PrintFinishTime
        Exit Sub
    ElseIf UBound(FileNames) = -1 Then
        MsgBox ("�[�Z��Ƥ�����")
        PrintFinishTime
        Exit Sub
    End If
    
    'Get amount of days in that month
    dayAmounts = nb_days_month
    
    'Create result folder
    Dim FolderName As String
    FolderName = Left(FileNames(1), InStrRev(FileNames(1), "\"))
    Dim fileName As String
    fileName = Right(FileNames(1), Len(FileNames(1)) - InStrRev(FileNames(1), "\"))
    Dim ResultFolderName As String
    ResultFolderName = FolderName + "����\"
    Dim fso As New FileSystemObject
    If Not fso.FolderExists(ResultFolderName) Then
        fso.CreateFolder ResultFolderName
        ProcessingBox.Value = ProcessingBox.Value & "<<�Хߧ�����Ƨ�>>" & vbCrLf
    End If
    
    
    
    Dim app As New Excel.Application
    app.Visible = False 'Visible is False by default, so this isn't necessary
    Dim book As Excel.Workbook
    Dim OvertimeFileName As String
    Dim i, s As Integer
    
    'Read salary data
    Set book = app.Workbooks.Add(SalaryButton.Caption)
    Dim salary As Object
    Set salary = ReadSalaryFromSalary(book)
    book.Close SaveChanges:=False
    Set book = app.Workbooks.Add(TreatmentButton.Caption)
    Set salary = ReadSalaryFromTreatment(book, salary)
    book.Close SaveChanges:=False
    
    'Read and Write overtime data
    Dim project As Object
    Set project = CreateObject("Scripting.Dictionary")
    For runIdx = 1 To 2
        For i = 1 To UBound(FileNames)
            If (runIdx = 1 And InStr(1, FileNames(i), "�M��.xls") <> 0) Or (runIdx = 2 And InStr(1, FileNames(i), "�M��.xls") = 0) Then
                Set book = app.Workbooks.Add(FileNames(i))
                
                Dim people As Object
                Set people = ReadOvertime(book, InStr(1, FileNames(i), "�M��.xls") <> 0, project)
                ProcessingBox.Value = ProcessingBox.Value & "<<Ū��>>" & FileNames(i) & vbCrLf
                
                WriteFile book, people, salary, InStr(1, FileNames(i), "�M��.xls") <> 0, project, ResultFolderName, FileNames(i)
                fileName = Right(FileNames(i), Len(FileNames(i)) - InStrRev(FileNames(i), "\"))
                ProcessingBox.Value = ProcessingBox.Value & "<<�g��>>" & ResultFolderName & fileName & "(����).xls" & vbCrLf
                
                book.Close SaveChanges:=False
                PrintFinishTime
            End If
        Next i
    Next runIdx
    
    app.Quit
    Set app = Nothing
    PrintFinishTime
    
End Sub

Private Sub ClearButton_Click()
    OvertimeTextBox = ""
End Sub

Private Sub OvertimeButton_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = True
    .Title = "Select file"
    .ButtonName = "Confirm"
    If .Show = -1 Then
        'ok clicked
        Dim Item
        For Each Item In .SelectedItems
            OvertimeTextBox.Value = OvertimeTextBox & vbCrLf & Item
        Next
    Else
        'cancel clicked
    End If
    
    End With
End Sub

Private Sub SalaryButton_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select file"
    .ButtonName = "Confirm"
    .InitialFileName = ThisWorkbook.Path
    If .Show = -1 Then
        'ok clicked
        SalaryButton.Caption = .SelectedItems(1)
    Else
        'cancel clicked
    End If
    
    End With
    
    If InStr(1, SalaryButton.Caption, "�~����") = 0 Then
        MsgBox ("�ɦW���t���u�~���ơv�A�нT�{���L����ɮ�!!")
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim d As Date
    d = Date
    'MsgBox ("�{�b�O�G" & d)
    UserForm1.YearText.Value = Year(d - 28)
    UserForm1.MonthText.Value = Month(d - 28)
    Set jobTitles = setJobTitles
End Sub

Private Sub TreatmentButton_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Select file"
    .ButtonName = "Confirm"
    .InitialFileName = ThisWorkbook.Path
    If .Show = -1 Then
        'ok clicked
        TreatmentButton.Caption = .SelectedItems(1)
    Else
        'cancel clicked
    End If
    
    End With
    
    If InStr(1, TreatmentButton.Caption, "�ݹJ�չ�") = 0 Then
        MsgBox ("�ɦW���t���u�ݹJ�չ�v�A�нT�{���L����ɮ�!!")
    End If
End Sub
