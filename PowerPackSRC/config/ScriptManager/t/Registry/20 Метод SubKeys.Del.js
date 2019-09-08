RegWrite("HKCU\\TestKey\\SubKey\\", 1)

R.CurrentKey = "HKCU\\TestKey"

var sk = R.SubKeys
assign(sk.Count(), 1)
assign(sk.Del("SubKey"), true)
assign(sk.Count(), 0)
assign(R.SubKeysCount(), 0)

RegDelete("HKCU\\TestKey\\")
