$NAME ���� ���� ��������
'      ������ ��� ���������� �������� �� �������� � ���� ������ �� Ctrl-2 (c) artbear,  2004
'
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057

' ������ �������� � ���� � �������� "�������"
' � �������������� ������� �� ����
' ���������� � bin\config\scripts
'
sConstMenuPrefix = "���� ���� ��������"

Dim Telepat

' ���������� ������� "����� ���� ��������"
' ��������� ����������� �������� ������� � ���� ��������.
' ��� ����� ������ ������� ������-��������� ����������� �������.
' ������ ����������� ����� ������ ������������� �� ��������� ������.
' ��������� ������ ������ ������������ �� ��������� � ������ ������.
' ����� �������� ������� ����� ������ | ����� ������� �������
' d ��� D (�� Disabled) - ����������� �����
' c ��� C (�� Checked)  - ����� � "��������"
' ����� ����� ������ | ����� ������� ������������� �������.
' � ���� ������ � ������� OnCustomMenu ������ �������� ������ ����
' ����� ������� ���� �������������
' ��� �������� ������������ ������� ��� "-"
'
Function Telepat_GetMenu()
  Telepat_GetMenu  = Telepat_GetMenu & vbTab & "������� (������ ������)" & vbCrLf

  Set e=CreateObject("Macrosenum.Enumerator")
  iScriptCount=Scripts.Count-1
  for i=0 to iScriptCount
    sScriptName = Scripts.Name(i)
    bInsertScript = 0
    Set script=Scripts(i)
    arr=e.EnumMacros(script)                        ' ��������� ������� �������� �������
    for j=0 to ubound(arr)
      if bInsertScript = 0 then
        Telepat_GetMenu  = Telepat_GetMenu & vbTab & sScriptName & vbCrLf
        bInsertScript = 1
      end if
      Telepat_GetMenu  = Telepat_GetMenu & vbTab & vbTab  & arr(j) & "| |"
      Telepat_GetMenu  = Telepat_GetMenu & sConstMenuPrefix & "Scripts(""" & sScriptName & """)." & "" & arr(j) & "" & vbCrLf
      'e.InvokeMacros Script, arr(j)  ' ����� �������: ������, ��� �������
    Next
  Next
End Function ' Telepat_GetMenu

' ���������� ������� OnCustomMenu.
' ���������� ��� ������ ������������� ������ ����,
' ������������ � "GetMenu"
' Cmd - �������� (��� �������������) ���������� ������ ����.
'
Sub Telepat_OnCustomMenu(Cmd)
    if InStr(Cmd, sConstMenuPrefix) = 1 then
      on error resume next
'      Execute "" & Cmd & "" ' ������� ������
      Eval "" & Replace(Cmd, sConstMenuPrefix, "") & ""
      on error goto 0
    end if
End Sub ' Telepat_OnCustomMenu

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
	
	' ��� �������� ������� �������������� ���
	c.AddPluginToScript SelfScript, "�������", "Telepat", Telepat
	
	SelfScript.AddNamedItem "CommonScripts", c, False
End Sub

Init 0
