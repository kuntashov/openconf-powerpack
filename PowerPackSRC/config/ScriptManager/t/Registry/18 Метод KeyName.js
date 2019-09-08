RegWrite("HKCU\\TestKey\\SubKey\\")

R.CurrentKey = "TestKey"
assign(R.KeyName, "TestKey")

R.CurrentKey = "\\TestKey"
assign(R.KeyName, "TestKey")

R.CurrentKey = "\\TestKey\\SubKey"
assign(R.KeyName, "SubKey")

RegDelete("HKCU\\TestKey\\SubKey\\")
RegDelete("HKCU\\TestKey\\")
