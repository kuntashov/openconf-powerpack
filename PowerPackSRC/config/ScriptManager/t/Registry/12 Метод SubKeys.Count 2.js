RegWrite("HKCU\\TestKey\\", 1)

R.CurrentKey = "HKCU\\TestKey\\"
assign(R.CurrentKey, "HKCU\\TestKey\\")
var sk = R.SubKeys;
assign(sk.Count(), 0)

close() // ����� ����� �� ������� ��������� �����
RegDelete("HKCU\\TestKey\\")
