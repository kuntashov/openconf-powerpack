' скрипт для ОпенКонф - взято с http://itland.ru/forum/index.php?act=ST&f=37&t=2929&st=
'   Автор - AlexQC
'   если в редактируемой форме или модуле есть директива #ЗагрузитьИзФайла -
'   открывает указанный файл.
'	При этом пути, начинающиеся с "\\" или с "Д:" (Д - имя диска)- считаются абсолютными и
'	передаются как есть.
'   В остальных случаях путь считается заданным относительно каталога базы.
'
'         Опубликовал и слегка подправил Артур Аюханов aka artbear
'         Мой e-mail: artbear@bashnet.ru
'         Мой ICQ: 265666057
'
' TODO (a13x)
'	1. выгрузка модуля во внешний файл с добавлением директивы #ЗагрузитьИзФайла
'	2. обратная операция - замена директивы #ЗагрузитьИзФайла содержимым соответствующего файла
'   3. При открытии объекта или внешней обработки/отчета при наличии в модуле директивы #ЗагрузитьИзФайла
'		открывать автоматически соответствующий файл как модуль


Sub OpenText(name)
    Set fs = CreateObject("Scripting.FileSystemObject")
    If left(name,2)="\\" or mid(name,2,1)=":" then
        ' Абсолютный путь
        str = name
    Else
        ' Относительный путь
        ' избавляемся от имени файла в конце пути к обработке, если это внешняя 
        ' обработка или внешний отчет, и строи иначе - это модуль объекта конфигурации
        ' и тогда путь надо разрешать относительно каталога базы
        Set re = New RegExp
        re.Pattern = "\\[\w\s]+\.ert$"
        re.IgnoreCase = true
        Path = re.Replace(DocPath, "")
        If Path = DocPath Then
            ' это модуль внутреннего объекта конфигурации, поэтому 
            ' путь будем смотреть относительно каталога базы
            Path = IBDir
        End If  
        str = CommonScripts.ResolvePath(Path, name)
        'Documents.Open str
    End If
    If Not fs.FileExists(str) Then
        'Set File = fs.CreateTextFile(str, 2)
        'File.Close
        
        Set w = Windows.ActiveWnd
        If w Is Nothing Then
            MsgBox "Нет активного окна", vbOKOnly, "TurboMD"
            Exit Sub
        End If
        Set d = w.Document
        If d.ID < 2 Then
            MsgBox "Окно ни форма, ни модуль", vbOKOnly, "TurboMD"
            Exit Sub
        End If
        If d = docText Then ' Просто модуль
            d.SaveToFile str
        Else
            If d = docWorkBook Then ' Форма
                Set m = d.Page(1) ' Выгружаем модуль
                m.SaveToFile str
            End If
        End If
        
    End If
    Documents.Open str
End Sub ' OpenText

Sub OpenDoc(doc)
    For I = 0 To doc.LineCount-1 ' Перебираем строки
        str = trim(doc.range(i,0))  ' & vbCRLF)
        If UCase(left(str,18)) = UCase("#ЗагрузитьИзФайла ") then
            ' Нашли загрузку - открываем
            str = trim(mid(str,19))
            if left(str,1) = """" then str = mid(str,2)
            if right(str,1) = """" then str = left(str, len(str)-1)
            OpenText str
            Exit Sub
        End If
        If UCase(left(str,14)) = UCase("#LoadFromFile ") then
            ' И то же с англ. вариантом
            str = trim(mid(str,15))
            if left(str,1) = """" then str = mid(str,2)
            if right(str,1) = """" then str = left(str, len(str)-1)
            OpenText str
            Exit Sub
        End If
    Next
End Sub ' OpenDoc

Sub OpenIncludeFile()
    Set d = CommonScripts.GetTextDoc(0)
    If d Is Nothing Then Exit Sub
    OpenDoc d
End Sub

'TODO
'Sub UnloadDocToExtFile()
'	Set d = CommonScripts.GetTextDoc(0)
'	If d Is Nothing Then Exit Sub
'End Sub

'Sub LoadDocFromExtFile()
'
'End Sub

' Процедура инициализации скрипта
Sub Init(dummy) ' Фиктивный параметр, чтобы процедура не попадала в макросы
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
End Sub ' Init

Init 0 ' При загрузке скрипта выполняем инициализацию
