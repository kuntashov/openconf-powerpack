RegWrite("HKCU\\TestKey\\", 1)

R.CurrentKey = "HKCU\\TestKey\\"
assign(R.CurrentKey, "HKCU\\TestKey\\")
assign(R.SubKeysCount(), 0)

close() // ����� ����� �� ������� ��������� �����
RegDelete("HKCU\\TestKey\\")
