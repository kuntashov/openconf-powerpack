$NAME Меню всех макросов
'      Скрипт для добавления макросов из скриптов в меню модуля по Ctrl-2 (c) artbear,  2004
'
'         Мой e-mail: artbear@bashnet.ru
'         Мой ICQ: 265666057

' Скрипт работает в паре с плагином "Телепат"
' и обрабатывающий события от него
' разместите в bin\config\scripts
'
sConstMenuPrefix = "Меню всех макросов"

Dim Telepat

' Обработчик события "Вызов меню шаблонов"
' Позволяет динамически добавить команды в меню шаблонов.
' Для этого должен вернуть строку-описатель добавляемых пунктов.
' Каждый добавляемый пункт должен располагаться на отдельной строке.
' Вложенные группы команд определяются по табуляции в начале строки.
' После названия команды через символ | можно указать символы
' d или D (от Disabled) - недоступный пункт
' c или C (от Checked)  - пункт с "галочкой"
' Далее через символ | можно указать идентификатор команды.
' В этом случае в событие OnCustomMenu вместо названия пункта меню
' будет передан этот идентификатор
' Для создания разделителей укажите имя "-"
'
Function Telepat_GetMenu()
  Telepat_GetMenu  = Telepat_GetMenu & vbTab & "Макросы (полный список)" & vbCrLf

  Set e=CreateObject("Macrosenum.Enumerator")
  iScriptCount=Scripts.Count-1
  for i=0 to iScriptCount
    sScriptName = Scripts.Name(i)
    bInsertScript = 0
    Set script=Scripts(i)
    arr=e.EnumMacros(script)                        ' Получение массива макросов объекта
    for j=0 to ubound(arr)
      if bInsertScript = 0 then
        Telepat_GetMenu  = Telepat_GetMenu & vbTab & sScriptName & vbCrLf
        bInsertScript = 1
      end if
      Telepat_GetMenu  = Telepat_GetMenu & vbTab & vbTab  & arr(j) & "| |"
      Telepat_GetMenu  = Telepat_GetMenu & sConstMenuPrefix & "Scripts(""" & sScriptName & """)." & "" & arr(j) & "" & vbCrLf
      'e.InvokeMacros Script, arr(j)  ' Вызов макроса: объект, имя макроса
    Next
  Next
End Function ' Telepat_GetMenu

' Обработчик события OnCustomMenu.
' Вызывается при выборе пользователем пункта меню,
' добавленного в "GetMenu"
' Cmd - название (или идентификатор) выбранного пункта меню.
'
Sub Telepat_OnCustomMenu(Cmd)
    if InStr(Cmd, sConstMenuPrefix) = 1 then
      on error resume next
'      Execute "" & Cmd & "" ' вызываю скрипт
      Eval "" & Replace(Cmd, sConstMenuPrefix, "") & ""
      on error goto 0
    end if
End Sub ' Telepat_OnCustomMenu

'
' Процедура инициализации скрипта
'
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
	
	' При загрузке скрипта инициализируем его
	c.AddPluginToScript SelfScript, "Телепат", "Telepat", Telepat
	
	SelfScript.AddNamedItem "CommonScripts", c, False
End Sub

Init 0
