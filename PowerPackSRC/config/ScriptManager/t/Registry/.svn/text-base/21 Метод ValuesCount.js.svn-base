RegWrite("HKCU\\TestKey\\", "")

R.CurrentKey = "HKCU\\TestKey"

assign(R.ValuesCount(), 1); // (значение по умолчанию)

RegWrite("HKCU\\TestKey\\Param1", "Value1")
RegWrite("HKCU\\TestKey\\Param2", "Value2")

assign(R.ValuesCount(), 3);

RegDelete("HKCU\\TestKey\\Param1")
RegDelete("HKCU\\TestKey\\Param2")
RegDelete("HKCU\\TestKey\\")