' (с) из ветки о Телепате (IAM и еще кто-то, точно не помню,)
' Оформление и коррекция: Артур Аюханов aka artbear
'
'	Версия: $Revision: 1.5 $ 
'
' ------------------------------------------------------------------
'
' выравнивает выделенный текст по знакам равно ("=")
'
' ------------------------------------------------------------------
Sub FormatBlock()
  Dim Positions(1000)
  Dim EmptySigns(1000)

  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub

  AllLines = Split(doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol), vbCrLf)
  MaxPos = 0
  for i = 0 to UBOund(AllLines)
    Positions(i) = instr(AllLines(i), "=")
    EmptySigns(i) = 0
    if instr(AllLines(i), "=") > 0 then
      for i1 = 1 to Len(AllLines(i))
        if instr(vbTab + " ", Mid(AllLines(i), i1, 1)) = 0 then
          Positions(i) = Positions(i) - i1 + 1
          EmptySigns(i) = i1 - 1
          exit for
        end if
      Next
    end if
    if Positions(i) > MaxPos then
      MaxPos = Positions(i)
    end if
  Next
  for i = 0 to UBOund(AllLines)
    Pos = Positions(i) + EmptySigns(i)
    if Pos > 0 then
      AllLines(i) = Left(AllLines(i), Pos - 1) + Space(MaxPos - Positions(i)) + Mid(AllLines(i), Pos)
    end if
  Next
  doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = Join(AllLines, vbCrLf)
End sub ' FormatBlock()

' ------------------------------------------------------------------
'
' дополняет форматирует строку пробелами для красоты
'
' ------------------------------------------------------------------
Sub FormatLines()
  Dim SpacesPlaces(1000)

  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub

  CountD = 0
  AllLines = Split(doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol), vbCrLf)
  for i = 0 to UBOund(AllLines)
    NextLine = AllLines(i)
    inString = False
     For i1 = 1 To Len(NextLine)
         Letter = Mid(NextLine, i1, 1)
         If Letter = """" Then
             inString = Not inString
         ElseIf Not inString And InStr("*/+-=<>,", Letter) > 0 Then
             If i1 > 1 And Letter <> "," Then
                 PreLetter = Mid(NextLine, i1 - 1, 1)
                 If InStr("[*/+-=<>, ", PreLetter) = 0 Then
            CountD = CountD + 1
            SpacesPlaces(CountD) = i1 - 1
                 End If
             End If
             If i1 < Len(NextLine) Then
                 PostLetter = Mid(NextLine, i1 + 1, 1)
                 If InStr("*/+-=<>, ]", PostLetter) = 0 Then
                     CountD = CountD + 1
            SpacesPlaces(CountD) = i1
                 End If
             End If
         End If
     Next
    for i1 = 1 to CountD
      SpacePos = SpacesPlaces(i1)
      AllLines(i) = Left(AllLines(i), SpacePos - 1 + i1) + " " + Mid(AllLines(i), SpacePos + i1)
    Next
    CountD = 0
  Next

  doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = Join(AllLines, vbCrLf)
end sub ' FormatLines()

' ------------------------------------------------------------------
'
' выравнивает выделенный текст по знакам равно ("=")
' теперь удаляются пробелы перед "=" и заменяются на табы
'
' ------------------------------------------------------------------
' Например:
' 		лМеню = 2;
' 		лМенюДляВызова = 1;
' превратится в
' 		лМеню          = 2;
' 		лМенюДляВызова = 1;
' ------------------------------------------------------------------
Dim TabSize
TabSize = 4 'длина таба в символах (в настройках 1С)
CharForFormating = "="

Sub FormatBlockWithTabs()
  Dim Positions(1000)
  Dim RealPos(1000) 'реальная позиция, с учётом размера табов

  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  	
  AllLines = Split(doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol), vbCrLf)
  MaxPosReal = 0

  for i = 0 to UBOund(AllLines)
    AllString = AllLines(i)
    'Positions(i) = instr( AllString, "=") - 1
    Positions(i) = instr( AllString, CharForFormating) - 1
    MultiCharForFormating = ""
    IF InStr(1,CharForFormating,",")>0 Then MultiCharForFormating = ","
    IF InStr(1,CharForFormating,".")>0 Then MultiCharForFormating = "."
    If Len(MultiCharForFormating)>0 Then
		If instr(1, CharForFormating, MultiCharForFormating)>0 Then
			Numb = 0 + Cint(Replace(CharForFormating,MultiCharForFormating,"0"))
			If Numb<>0 Then
				Positions(i) = 0
				ArrW = Split(AllString, MultiCharForFormating)
				ttext = ""
				If Ubound(ArrW)>= Numb Then
					For ww = 0 To Numb-1
						ttext = ttext & " " & ArrW(ww)
					Next
					ttext = ttext & MultiCharForFormating
					Positions(i) = instr( ttext, MultiCharForFormating) - 1
				End if    
			End if    
		End if    
	End if    

    RealPos(i) = Positions(i)
    if Positions(i) > 0 then
      LeftString = RTrim( Left( AllString, Positions(i) ) )
      RealPos(i) = RealStringLen(LeftString)
    end if

    if RealPos(i) > MaxPosReal then MaxPosReal = RealPos(i)
  Next

  MaxPosReal = Fix( MaxPosReal /TabSize + 0.99) * TabSize 'округлим максимум до границы табов
  for i = 0 to UBOund(AllLines)
    Pos = Positions(i)
    Real_Pos = RealPos(i)
    if Pos > 0 then
      DopString = ""

      AllTabs = Fix( (MaxPosReal - Real_Pos)/TabSize + 0.99 )
      DopString = DopString + String(AllTabs, vbTab)

      AllLines(i) = RTrim( Left(AllLines(i), Pos) ) + DopString + Mid(AllLines(i), Pos +1)
    end if
  Next

  doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = Join(AllLines, vbCrLf)

End sub ' FormatBlockWithTabs

' Функция определения реальной длины строки. Заморочки с "неполными" табами.
Function RealStringLen( tekStr )
  RealStringLen = 0
  for i = 1 to Len(tekStr)
    tek_Simbol = Mid( tekStr, i, 1)

    If tek_Simbol <> vbTab Then
       RealStringLen = RealStringLen + 1
    Else
       Ostatok_Deleniya = RealStringLen - Fix( RealStringLen / TabSize) * TabSize
       RealStringLen = RealStringLen + (TabSize -Ostatok_Deleniya)
    End If
  Next
End Function

Sub ChoiseCharForFormating()
	ttext = "=" & vbCrLf & "<>" & vbCrLf & "." & vbCrLf & ".2" & vbCrLf & ".3" & vbCrLf & ".4" & vbCrLf & "{"
	ttext = ttext & vbCrLf & "}" & vbCrLf & "("  & vbCrLf & ")" & vbCrLf & "-" & vbCrLf & "+" & vbCrLf & ","
	ttext = ttext & vbCrLf & ",2" & vbCrLf & ",3"  & vbCrLf & ",4" & vbCrLf & ",5" & vbCrLf & ",6" & vbCrLf & ",7" & vbCrLf & "//"
	ttext = CommonScripts.SelectValue(ttext, "Выбор символа форматирования")
	If Len(ttext)>0 Then
		CharForFormating = ttext
	End IF
End Sub

'
' Процедура инициализации скрипта
'
Private Sub Init() ' Фиктивный параметр, чтобы процедура не попадала в макросы
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
End Sub
 
Init ' При загрузке скрипта выполняем инициализацию
