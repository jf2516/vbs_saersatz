'# ---------------------------------------------------------------
'# SAersatz.vbs
'# Отправка файлов через ПТК с помощью транспортных конвертов
'# на основании письма Банка России от 08.12.2021 N Т258-15-4/7317
'# Дата создания      : 20.12.2021
'# Дата редактирования: 20.12.2021
'#----------------------------------------------------------------

Set Shell = WScript.CreateObject("WScript.Shell")
Set FSO   = WScript.CreateObject("Scripting.FileSystemObject")

Dim CurrentDate: CurrentDate = Date()
Dim WorkDir    : WorkDir = "p:\transport_test\"
Dim IndexDir   : IndexDir = WorkDir & "\" & "_index\"
Dim Forms      : Forms = array("2o_4512U", "4e_4498U", "4q_F601", "r3_FZVBK", "r9_CLIENT")
Dim FileTypeExt: FileTypeExt = ".index"
Dim FileExt    : FileExt = "755.005"
Dim SubDir     : SubDir = array("in", "out", "ptk", "backup") 


Function To36(Digit)
'#-перевод в 36 систему
    Abc36 = "123456789abcdefghijklmnopqrstuvwxyz"
    To36 = Mid(Abc36, Digit, 1)  
End Function

Function LeadZero(ZeroCount, Number)
'#-добавление ведущих нулей
Dim l
   l=Len(Number)       
   If l<ZeroCount Then
      LeadZero=String(ZeroCount-l, "0") & Number
   Else
      LeadZero=Number
   End If          
End Function

Sub CreateSubDir(DirName)
'#-проверка наличия каталога. Если отсутствует, то создать   
   If Not FSO.FolderExists(DirName) Then 
      Shell.Run "%comspec% /c mkdir " & DirName, 0, True
   End If
End Sub

Function CreateIndexFile(FileType)
'#-нициализация счётчика
    Set fd = FSO.OpenTextFile(IndexDir & FileType & FileTypeExt, 2, True)
    fd.WriteLine CurrentDate & "|" & 1
    fd.Close
    CreateIndexFile = FileType & To36(Day(CurrentDate)) & LeadZero(2,1) & FileExt
End Function

Function ChangeIndex(FileType)
'#-увеличение счётчика
Dim Index
Dim WorkDate
    
    '#-читаем предыдущее значение
    Set fd = FSO.OpenTextFile(IndexDir & FileType & FileTypeExt, 1, True)
    Ret = Split(fd.ReadLine,"|")
    WorkDate = Ret(0)
    Index    = Ret(1)
    fd.close
    
    '#-обновляем
    If DateDiff("d", WorkDate, CurrentDate) = 0 Then
        Index = Index + 1
        Set fd = FSO.OpenTextFile(IndexDir & FileType & FileTypeExt, 2, True)
        fd.WriteLine WorkDate & "|" & Index
        fd.Close
        NewFileName = FileType & To36(Day(WorkDate)) & LeadZero(2, Index) & FileExt
    Else
        NewFileName = CreateIndexFile(FileType)
    End if
    
    ChangeIndex = NewFileName 
End Function


If Not FSO.FolderExists(IndexDir) Then
    FSO.CreateFolder(IndexDir)
End If

For Each i in Forms
    FileType = Left(i,2)
    FormDir = Right(i,Len(i)-3)
    
    For Each j in SubDir
        CreateSubDir(WorkDir & i & "\" & j)
    Next
    
    Set OutFiles = FSO.GetFolder(WorkDir & i & "\ptk").Files
    If OutFiles.Count > 0 Then
        For Each f in OutFiles
            Select Case FSO.FileExists(IndexDir & FileType & FileTypeExt)
                Case True : FileName = ChangeIndex(FileType)
                Case False: FileName = CreateIndexFile(FileType)
            End Select

            ArjFileName = Left(FileName,Len(FileName)-4) & ".arj"
            Shell.Run "%comspec% /c copy /y " & f & " " & WorkDir & i & "\out\" & FileName, 0, True
            Shell.Run "arj32 a -ef " & WorkDir & i & "\backup\" & ArjFileName & " " & f, 0, True
            Shell.Run "%comspec% /c del /q " & f, 0, True
        Next
    End If
    
    
Next

'#-конец файла

