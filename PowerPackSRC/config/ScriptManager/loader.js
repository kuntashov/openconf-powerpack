// ��������� ��� "������-����������" ������� ������-�������
Scripts.Load(BinDir + "config\\system\\ScriptManager\\HotKeysHandler.js")
// ������ ��� �������������� ������� (����� ������� �������� �� �������)
Scripts.Load(BinDir + "config\\system\\ScriptManager\\HotKeysEditor.js")
// ...� ��������� ���� �� �������������
Scripts.UnLoad(SelfScript.Name)