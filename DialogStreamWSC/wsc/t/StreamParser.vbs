' Sorry for my too poor... VBScript :-) I had never spoken it before, really -- a13x
' ���... ��� ��� �� �����. �����, ��������, ����� ���� ������ �� ������, ���� 
' ������ � ������� ��������, ��� �����, �������� :-)

Sub Print(str)
    WScript.Echo str
End Sub
               
Set sp = CreateObject("OpenConf.StreamParser")

' ������������ ������ � ��������� �� VBS �� ������� � ���� ����, ��� ����� ��� ����
' ������������ ��� ����������� ����� �, ������� - ������������... � ��������� ���������
' � ����� �������� ��� � ������� �� ���� :-(, ������� ������� End... ���� ���������, 
' "� ��� ������", �������� ������ �� ���� ������� :-) JScript'�
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
