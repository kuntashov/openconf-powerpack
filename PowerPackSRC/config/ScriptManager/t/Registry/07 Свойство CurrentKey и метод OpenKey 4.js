// ���������� ���� 2
R.CurrentKey = "HKEY_CLASSES_ROOT\\.txt"
assign(R.RootKey, "HKEY_CLASSES_ROOT")
assign(R.CurrentKey, "HKEY_CLASSES_ROOT\\.txt")

close() // ����� ����� �� ������� ����
RegDelete("HKCU\\TestKey\\SubKey\\")
RegDelete("HKCU\\TestKey\\")
