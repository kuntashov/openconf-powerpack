RegWrite("HKCU\\TestKey\\Param", "Value") 

R.CurrentKey = "HKCU\\TestKey"
assign(R.ValuesCount(), 1)

R.DeleteValue("Param");
assign(R.ValuesCount(), 0);

RegDelete("HKCU\\TestKey\\")
