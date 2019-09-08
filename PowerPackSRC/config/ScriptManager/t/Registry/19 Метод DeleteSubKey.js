RegWrite("HKCU\\TestKey\\SubKey\\", 1)

R.CurrentKey = "HKCU\\TestKey"

assign(R.SubKeysCount(), 1)
assign(R.DeleteSubKey("SubKey"), true)
assign(R.SubKeysCount(), 0)

RegDelete("HKCU\\TestKey\\")
