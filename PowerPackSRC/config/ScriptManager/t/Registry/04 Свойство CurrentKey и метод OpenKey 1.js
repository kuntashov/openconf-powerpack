RegWrite("HKCU\\TestKey\\SubKey\\", "Hello, World!")

// ���������� ���� 1
R.OpenKey("HKCU\\TestKey")
assign(R.CurrentKey, "HKCU\\TestKey")
assign(R.KeyName, "TestKey")
