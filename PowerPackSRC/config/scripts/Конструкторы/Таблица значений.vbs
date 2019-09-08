' IAm
'Первый позволяет быстро инициализировать таблицу значений, введя имена колонок через запятую, второй позволяет по имени таблицы создать новую строку, ищет по тексту строки с методом "НоваяКолонка()" и автоматом прописывает поля. То есть если в модуле был блок
'Табл.НоваяКолонка("Контр");
'Табл.НоваяКолонка("Договор");
'Табл.НоваяКолонка("Сумма");
'то можно вызвать скрипт, ввести имя таблицы(Табл)
'и получить текст в виде:
'Табл.Контр =
'Табл.Договор =
'Табл.Сумма =
'==================================================
Sub NovayaTabliza()
  If Windows.ActiveWnd Is Nothing Then
     MsgBox "Нет активного окна"
     Exit Sub
  End If
  Set doc = Windows.ActiveWnd.Document
  If doc=docWorkBook Then Set doc=doc.Page(1)
  If doc<>docText Then
     MsgBox "Окно не текстовое"
     Exit Sub
  End If

  TabInd = vbCrLf + String(doc.SelStartCol, vbTab)

  VTName = inputBox("введите имя таблицы значений:")
  ColumnsList = inputBox("введите имена колонок через запятую")
  VTColumns = Split(ColumnsList, ",")

  Txt = VTName + " = СоздатьОбъект(""ТаблицаЗначений"");"
  for i = 0 to UBound(VTColumns)
    Txt = Txt + TabInd + VTName + ".НоваяКолонка(""" + Trim(VTColumns(i)) + """);"
  Next

  doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub 'NovayaTabliza()
'
'====================================================
Sub NovayaStroka()

  If Windows.ActiveWnd Is Nothing Then
     MsgBox "Нет активного окна"
     Exit Sub
  End If
  Set doc = Windows.ActiveWnd.Document
  If doc=docWorkBook Then Set doc=doc.Page(1)
  If doc<>docText Then
     MsgBox "Окно не текстовое"
     Exit Sub
  End If

  TabInd = vbCrLf + String(doc.SelStartCol, vbTab)

    VTName = inputBox("введите имя таблицы значений:")
  NewColumnText = UCase(VTName + ".НоваяКолонка(""")
  StrLen = Instr(NewColumnText, """")

    TextD = doc.text
  TextDUpper = UCase(TextD)

  Txt = VTName + ".НоваяСтрока();"

  StartPos = 0
  Pos = instr(TextDUpper, NewColumnText)
    do while Pos > 0
    StartPos = Pos + 10
    Pos2 = instr(Pos + StrLen, TextDUpper, """")
    if Pos2 - Pos - StrLen < 30 then
      Txt = Txt + TabInd + VTName + "." + Mid(TextD, Pos + StrLen, Pos2 - Pos - StrLen) + " = ;"
    end if
    Pos = instr(StartPos, TextDUpper, NewColumnText)
  Loop

  doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
end sub 'NovayaStroka()