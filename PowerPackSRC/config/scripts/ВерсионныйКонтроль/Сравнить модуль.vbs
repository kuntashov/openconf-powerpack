$NAME Сравнить и объединить модуль
' макросы для сравнения и объединения модулей
'
'         Мой e-mail: artbear@bashnet.ru
'         Мой ICQ: 265666057
'
' БОлее точное название - "Сравнить и объединить модуль",
'т.е. скрипт позволяет изменить текущий модуль путем объединения с каким-то другим,
'удобно, например, при объединении с типовой.
' ПОКА Cкрипт работает только в окне модуля.
' Схема работы такая: у пользователя запрашивается имя файла для сравнения
'с текущим модулем, (ВНИМАНИЕ: пока файл должен быть только текстовым),
'далее текст текущего модуля сохраняется во временный файл,
'затем запускается мержилка (кдифф3), в командной строке которой указаны:
'файл для сравнения, файл текущего модуля, и файл результата.
'Пользователь выполняет (с помощью кдифф3) объединение 2 модулей в третий, конечный, и возвращается в Конфигуратор.
'Здесь скрипт анализирует наличие третьего, конечного, файла,
'и если файл присутствует, задает вопрос "Заменить текущий модуль?".
'После ответа "Да", текст текущего модуля заменяется на результат объединения,что и требуется.
' PS в качестве мержилки используется кдифф3.
' PSS командная строка для мержилки находится в начале скрипта и имеет простой, понятный формат. Можете менять ком. строку на свою любимую мержилку.
' PSSS В планах:
' 1) объединение не только модуля, но и формы, и описания.
' 2) сравнение/объединение не только 2-х файлов, но и 3 файлов в один (фича кдифф3 - очень удобно и максимально быстро, рекомендую)'
'
'
' параметры комадной строки для мержилки:
'%1 - файл, в который будет помещен текст открытого модуля
'%2 - файл, с которым идет сравнение
'%3 - файл-результат сравнения/объединения
'
sCmdLineTemplate = "kdiff3 %1 %2 -o %3"

Function GetStringFromTemplate(sTemplate, StringArray)
    sString = sTemplate
    For i = LBound(StringArray) To UBound(StringArray)
        sString = Replace(sString, "%" & CStr(i + 1), """" & StringArray(i) & """")
    Next
    For i = UBound(StringArray) + 1 To 10
        sString = Replace(sString, "%" & CStr(i + 1), "")
    Next
    GetStringFromTemplate = Trim(sString)
End Function ' GetStringFromTemplate

Sub RunDiffWithArray(arr)
    CmdLine = GetStringFromTemplate(sCmdLineTemplate, arr)
'CommonScripts.Echo CmdLine

    CommonScripts.RunCommandAndWait CmdLine, ""
End Sub ' RunDiffWithArray

Function Diff2Files(File1, File2)
    DiffFiles = ""

    arr = Split(File1 & "," & File2, ",")
    RunDiffWithArray arr
End Function ' Diff2Files

Sub TestDiff()
    Diff2Files "W:\ВремРасчет.txt", "W:\ВремРасчет рез.txt"
End Sub

Function Diff3Files(File1, File2, File3)
    DiffFiles3 = ""

    arr = Split(File1 & "," & File2 & "," & File3, ",")
    RunDiffWithArray arr
End Function ' Diff3Files

Function DiffCurrentModuleWithFile(File2)
    DiffCurrentModule = ""

    Set Doc = CommonScripts.GetTextDocIfOpened(0)
    If Doc Is Nothing Then Exit Function

    Set fs = CreateObject("Scripting.FileSystemObject")

'    On Error Resume Next
	Const TemporaryFolder  = 2
	TempFolderName = fs.GetSpecialFolder(TemporaryFolder).Path
	if Right(TempFolderName, 1) <> "\" then
		TempFolderName = TempFolderName + "\"
	end if

    sSourceTempFile = TempFolderName & "tempForDiff_56731.txt"
    sDstTempFile = TempFolderName & "tempForDiff_56732.txt"

    If fs.FileExists(sSourceTempFile) Then fs.DeleteFile sSourceTempFile
    If fs.FileExists(sDstTempFile) Then fs.DeleteFile sDstTempFile

    Doc.SaveToFile sSourceTempFile
'CommonScripts.Echo sSourceTempFile
    Diff3Files sSourceTempFile, File2, sDstTempFile

    If fs.FileExists(sDstTempFile) Then
        If MsgBox("Заменить текст модуля на результат объединения", vbOKCancel + vbQuestion, SelfScript.Name) = vbOK Then
            Doc.LoadFromFile sDstTempFile
        End If
    End If
    If fs.FileExists(sSourceTempFile) Then fs.DeleteFile sSourceTempFile
    If fs.FileExists(sDstTempFile) Then fs.DeleteFile sDstTempFile
End Function ' DiffCurrentModuleWithFile

Sub TestDiffCurrentModule()
    DiffCurrentModuleWithFile "W:\ВремРасчет рез.txt"
End Sub

Sub DiffCurrentModuleWithFileSelection()
    Set srv = CreateObject("Svcsvc.Service")
    sDiffFileName = CommonScripts.SelectFileForRead("", "Модули 1С и текстовые файлы|*.1s;*.txt|Модули 1С|*.1s|Текстовые файлы|*.txt|Все файлы|*")
    If sDiffFileName = "" Then
        Exit Sub
    End If
    DiffCurrentModuleWithFile sDiffFileName
End Sub 'DiffCurrentModuleWithFileSelection

' Возвращает полный путь к папке с исполняемым файлом KDiff3
' если имеется соответсвующая запись в реестре (дабы исключить
' необходимость изменять скрипт, когда kdiff3 не прописан в path
Private Function GetKDiff3Folder()
	GetKDiff3Folder = ""
	On Error Resume Next
	GetKDiff3Folder = CStr(CommonScripts.WSH.RegRead("HKCU\Software\KDiff3\"))
	On Error GoTo 0
End Function ' GetKDiff3Folder

Function Text2File(text, path)
    Set textStream = CommonScripts.fso.CreateTextFile(path, 1, 0)
    textStream.Write text
    textStream.Close
    Text2File = path
End Function 'Text2File

Sub DiffSelectionWithClipboard()
    tempPath = CommonScripts.fso.GetSpecialFolder(2).Path
	if Right(tempPath, 1) <> "\" then
		tempPath = tempPath + "\"
	end if
	selFile  = Text2File(CommonScripts.SelectedText, tempPath & "temp1cForDiff_2.txt")
	clipFile = Text2File(CommonScripts.GetFromClipboard, tempPath & "temp1cForDiff_1.txt")

    Diff3Files selFile, clipFile, tempPath & "temp1cForDiff_result.txt"
End Sub 'DiffSelectionWithClipboard

Sub Init(param)
    Set c = Nothing
    On Error Resume Next
    Set c = CreateObject("OpenConf.CommonServices")
    On Error GoTo 0
    If c Is Nothing Then
        Message "Не могу создать объект OpenConf.CommonServices", mRedErr
        Message "Скрипт " & SelfScript.Name & " не загружен", mInformation
        Scripts.UnLoad SelfScript.Name
		Exit Sub
    End If
    c.SetConfig(Configurator)
	SelfScript.AddNamedItem "CommonScripts", c, False
	on Error goto 0
	' формируем полный путь к исполняемому файлу KDiff3
	tmp = CommonScripts.FSO.BuildPath(GetKDiff3Folder(), "kdiff3.exe")
	If CommonScripts.FSO.FileExists(tmp) Then
		sCmdLineTemplate = Replace(sCmdLineTemplate, "kdiff3", """" & tmp & """")
	End If
End Sub

Init 0
