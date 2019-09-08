RegWrite("HKCU\\TestKey\\Param1", "Value1")
RegWrite("HKCU\\TestKey\\Param2", "Value2")

R.CurrentKey = "HKCU\\TestKey"
assign(R.Values.Count(), 2);

RegDelete("HKCU\\TestKey\\Param1")
RegDelete("HKCU\\TestKey\\Param2")
RegDelete("HKCU\\TestKey\\")