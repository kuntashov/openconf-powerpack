' ������ �������������
Set srv=CreateObject("Svcsvc.Service")
Vals=""
for i=1 to 99
	vals=vals & "��������" & CStr(i) & vbCrLf
next

Vals=Vals & "��������100"
MsgBox srv.FilterValue(Vals,1)

MsgBox srv.SelectFolder("������� ��� ����","c:\",0)
MsgBox srv.SelectFolder("������� �������","c:\",1)
MsgBox srv.SelectFolder("������� ������� ��� ����","c:\",1 + &h4000,"c:\")

MsgBox srv.SelectFile(false,"c:\1","��� �����|*",true)

MsgBox srv.SelectValue(Vals,"��������, ����",False)

Vals=Array("�������� 1","�������� 2|c","�������� 3")
Vals=Join(Vals,vbCrLf)
MsgBox srv.SelectValue(Vals,"������� ������ ��������",True)

Vals= _
"����� 1|e" & vbCrLf & _
vbTab & "�������� 1_1" & vbCrLf & _
vbTab & "�������� 1_2" & vbCrLf & _
vbTab & vbTab & "�������� 1_2_1" & vbCrLf & _
"����� 2" & vbCrLf & _
vbTab & "�������� 2_1" & vbCrLf & _
vbTab & "�������� 2_2" & vbCrLf & _
"�������� 3"

MsgBox srv.SelectInTree(Vals,"�������� ��������",False)
MsgBox srv.SelectInTree(Vals,"�������� �������� ��� �����",False,False)

Vals= _
"����� 1|ce" & vbCrLf & _
vbTab & "�������� 1_1" & vbCrLf & _
vbTab & "�������� 1_2|c" & vbCrLf & _
vbTab & vbTab & "�������� 1_2_1" & vbCrLf & _
"����� 2|e" & vbCrLf & _
vbTab & "�������� 2_1|c" & vbCrLf & _
vbTab & "�������� 2_2" & vbCrLf & _
"�������� 3"
MsgBox srv.SelectInTree(Vals,"�������� ������ ������",True)

Vals= _
"����� 1" & vbCrLf & _
vbTab & "�������� 1_1|d" & vbCrLf & _
vbTab & "�������� 1_2|c" & vbCrLf & _
vbTab & vbTab & "�������� 1_2_1|dc" & vbCrLf & _
"����� 2" & vbCrLf & _
vbTab & "�������� 2_1| |val21" & vbCrLf & _
vbTab & "�������� 2_2|c|val22" & vbCrLf & _
"�������� 3"
MsgBox srv.PopupMenu(vals,0)
MsgBox srv.PopupMenu(vals,1)
MsgBox srv.PopupMenu(vals,2,300,300)