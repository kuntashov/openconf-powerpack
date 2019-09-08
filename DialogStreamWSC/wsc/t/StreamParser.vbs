' Sorry for my too poor... VBScript :-) I had never spoken it before, really -- a13x
' вот... “ак что не бейте. ¬роде, старалс€, чтобы было похоже на бейсик, даже 
' скобки в вызовах процедур, где можно, поубирал :-)

Sub Print(str)
    WScript.Echo str
End Sub
               
Set sp = CreateObject("OpenConf.StreamParser")

' ƒемонстраци€ работы с кавычками на VBS не удалась в виду того, что здесь они тоже
' используютс€ как разделитель строк и, кажетс€ - единственный... ј сосчитать правильно
' в конце рабочего дн€ € кавычки не смог :-(, заменил строкой End... кому интересно, 
' "в чем прикол", смотрите пример на моем любимом :-) JScript'е
sp.Source = "{""Main stream"", ""Test2"", {""Child stream"", ""0"", ""1"", ""2""}, ""End""}"

If Not sp.Parse Then
    Print sp.LastError
    WScript.Quit
End If

Print sp.Stream.Data.Item(0).Data
Print sp.Stream.Data.Item(1).Data

Set c = sp.Stream.Data.Item(2).Data

s = ""
For i=0 To c.Size - 1 
    s = s & c.Item(i).Data & " "
Next

Print "Child stream data: " & s

Print sp.Stream.Data.Item(3).Data
Print "------------------------------------------"
Print "Full source:"
Print sp.Stream.Stringify()
