RegWrite("HKCU\\TestKey\\")

R.CurrentKey = "HKCU\\TestKey\\"

assign(R.SubKeysCount(), 0)
assign(R.SubKeys.Add("SubKey"), true)
assign(R.SubKeysCount(), 1)
assign(R.SubKeyName(0), "SubKey")
assign(R.SubKeys(0).CurrentKey, "HKCU\\TestKey\\SubKey")

close()
RegDelete("HKCU\\TestKey\\SubKey\\")
RegDelete("HKCU\\TestKey\\")
