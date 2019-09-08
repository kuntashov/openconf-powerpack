$NAME ����� ������ � ������ ����������� �����������
'----------------------------------------------------------------
' �����: ������ ������ (Phoenix)
' Email: PhoenixUSA#yandex(dot)ru
'
' ����� ������ � ������� ������
' � ���������� ��������
' ������������ ���������� ���������
'
' ���������� �����������: ����� ������� aka artbear
' Email: artbear@bashnet.ru
'
' �� ������ ������: 
' ������ ��� ������ ���������, ������� �� ����������� � ������������ 1C 
' ^/?([^/]/?)*���������
' ��� 
' ^(?!.*//).*���������
' ����������: �� ��������� ������� mikeA � ����� 6 c �1�
'
' ������ ��� ������ ���������, �������, ��������, ����� � ������������ 1C 
' //+.*���������
' 
' TODO:
'     1) ����� ���� ������ ������ ��������� � �� ���������� ������� �������� ���� ������
'     2) ����� ����� ��������� ������ ��������� ������. 
'         � ������ ��������� ������������ ����� ������������ � ������� ��� ������ ������ ������ ��������� (?���� ��� ����� ����)
'
'======= ������������� =============================
Set RE = New RegExp

PrevRegPattern = "" ' ��� ���������� ������      
' ������ ���������� ������
Dim RegPatternCollection  ' "OpenConf.Collection"

'
' ��������� ������������� �������
'
Sub Init(dummy) ' ��������� ��������, ����� ��������� �� �������� � �������
    Set c = Nothing
    On Error Resume Next
    Set c = CreateObject("OpenConf.CommonServices")
    On Error GoTo 0
    If c Is Nothing Then
        Message "�� ���� ������� ������ OpenConf.CommonServices", mRedErr
        Message "������ " & SelfScript.Name & " �� ��������", mInformation
        Scripts.UnLoad SelfScript.Name
		Exit Sub
    End If
    c.SetConfig(Configurator)
	SelfScript.AddNamedItem "CommonScripts", c, False
	
    On Error Resume Next
	set RegPatternCollection = CreateObject("OpenConf.Collection")
    On Error GoTo 0
    If RegPatternCollection Is Nothing Then
        Message "�� ���� ������� ������ OpenConf.Collection", mRedErr
        Message "������ " & SelfScript.Name & " �� ��������", mInformation
        Scripts.UnLoad SelfScript.Name
		Exit Sub
    End If
	RegPatternCollection.Add("^/?([^/]/?)*�����������������������������")
	RegPatternCollection.Add("//+.*����������������������������")
End Sub
 
Init 0 ' ��� �������� ������� ��������� �������������

'----------------------------------------------------------------
sub GoFindText(DoBookmark)
  Dim strDoc, fStr
  
  Set Doc = CommonScripts.GetTextDocIfOpened(0)
  If Doc Is Nothing Then Exit Sub

  if DoBookmark <> -1 then ' ���� �� ������� �������� 
    if RegPatternCollection.Size > 0 then 
      PrevRegPattern = RegPatternCollection.Item(RegPatternCollection.Size-1)
    end if
  end if

  vDefaultWord = doc.CurrentWord
  vDefaultWord = Trim(vDefaultWord)
  if DoBookmark <> -1 then ' ���� �� ������� �������� 
  	if vDefaultWord <> "" then
	  RegPatternCollection.Add(vDefaultWord)
    end if

	arrForSelect = RegPatternCollection.SaveToString(vbCrLf)

	fStr = CommonScripts.SelectValue(arrForSelect, "������� �������� ��� ������", "", true, true)
	if fStr <> vDefaultWord then
      if vDefaultWord <> "" then
		RegPatternCollection.Remove(vDefaultWord)
      end if
	end if
	
  else
	fStr = PrevRegPattern
  end if
  if fStr = "" then
  	Exit Sub
  end if

    if DoBookmark = 1 then
	  if doc.NextBookmark(0) > 0 then
	  	'if fStr <> PrevRegPattern then ' TODO ? ���� ����� �����������, ������ ��� �������� �������� �� �����
          if MsgBox("������ �������� ����������� ������?",vbYesNo,"�����") = vbYes then
            'doc.ClearAllBookMark
		    ClearAllPreviousFindBookmarks() ' �������� ��������� ������� �� ���������� �������
          end if
        'end if
      end if
    end if

   RegPatternCollection.Add(fStr)

  	
    RE.Pattern = fStr
    RE.IgnoreCase = true

	vals = ""
	
    For i = 0 To doc.LineCount - 1
      strDoc = doc.Range(i, 0)
      if RE.Test(strDoc) then
        strDoc = Trim(strDoc)
        vals = vals & vbCrLf & "(" & CStr(i+1) & ") " & Replace(strDoc,vbTab,"")

        if DoBookmark = 1 then
          doc.Bookmark(i) = true
        end if
        if DoBookmark = -1 then
          doc.Bookmark(i) = false
        end if
      End If
    Next
    if DoBookmark = -1 then
	  exit sub
    end if
    
	if Trim(Vals) = "" then
	  exit sub 
	end if
	' ���� ������� ����� ���� �������, ����� � �������� �� ���
	arrVals = split(Vals, vbCrLf)
	if UBound(arrVals) - LBound(arrVals) = 1 then
	  fStr = arrVals(1)
	else
	  fStr = CommonScripts.SelectValue(Vals)
	end if
    if fStr <> "" then
      strDoc = Mid( fStr, InStr(fStr, "(")+1 , InStr(fStr, ")")-2)
      doc.MoveCaret CInt(strDoc)-1,0
    end if

End sub

'----------------------------------------------------------------
sub FindWithAddBookmarks()
  GoFindText(1)
end sub

'----------------------------------------------------------------
sub FindWithoutAddBookmarks()
  GoFindText(0)
end sub

'----------------------------------------------------------------
' ����������� ������� ��� ��������, ��������� ���������� �������
sub ClearAllPreviousFindBookmarks()
  GoFindText(-1)
end sub