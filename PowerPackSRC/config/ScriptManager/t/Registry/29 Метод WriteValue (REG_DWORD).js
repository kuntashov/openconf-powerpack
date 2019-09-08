RegWrite("HKCU\\TestKey\\", "1") 

R.CurrentKey = "HKCU\\TestKey"
assign(R.ValuesCount(), 1) // значение по умолчанию 

R.WriteValue("Integer", 0x666, "REG_DWORD")
assign(R.ValuesCount(), 2);
assign(R.ReadValue("Integer"), 0x666)

RegDelete("HKCU\\TestKey\\Integer")
RegDelete("HKCU\\TestKey\\")