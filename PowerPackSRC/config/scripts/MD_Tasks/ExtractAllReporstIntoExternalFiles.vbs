$NAME Extract All Reporst Into External Files

'"���������� �������/��������� �� ������� �����"
' MetaEditor (shotfire@inbox.ru)

Dim FName ' ��� ����� ��� ��������������� � ������ ����������
Dim flAutoSave ' ���� - ������� ��������������� ���������� ������ �� ����� �������
flAutoSave = false

'========================================================================
Sub SaveToERT()
Set srv = CreateObject("Svcsvc.Service")
SaveFolder = srv.SelectFolder("��������� �:")
If SaveFolder = "" Then Exit Sub : End If

SaveObjects = srv.SelectValue("������|c" & vbCrLf & "���������|c","��� ���������...",True)

If SaveObjects = "" Then Exit Sub : End If

Set GotoTreeScript = scripts("NavigationTools")
flAutoSave = true

If InStr(SaveObjects,"������") > 0 Then
Set Reps = MetaData.TaskDef.Childs(5) ' ������
For i = 0 To Reps.Count - 1
	RepName = "�����." & Reps(i).Name
	FName = SaveFolder & "\" & RepName & ".ert"
	GotoTreeScript.GoToMDTreeItem RepName,33239, false
	message FName & " ���������...", mInformation
Next                  
End If

If InStr(SaveObjects,"���������") > 0 Then
Set Calcs = MetaData.TaskDef.Childs(6) ' ���������
For i = 0 To Calcs.Count - 1
	CalcName = "���������." & Calcs(i).Name
	FName = SaveFolder & "\" & CalcName & ".ert"
	GotoTreeScript.GoToMDTreeItem CalcName,33239, false
	message FName & " ���������...", mInformation
Next                           
End If

message ""
message "��������� ���������...", mInformation

flAutoSave = false
End Sub  

'========================================================================
Sub Configurator_OnFileDialog(Saved, Caption, Filter, FileName, Answer) 
	If flAutoSave = false then Exit Sub
	If (Instr(Caption,"��������� ���") = 1) and (InStr(UCase(Filter), "*.ERT") > 0) Then 
		FileName = FName
		Answer = mbaOK
	End If
End Sub
