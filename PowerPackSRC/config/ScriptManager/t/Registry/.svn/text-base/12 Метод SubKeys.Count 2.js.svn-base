RegWrite("HKCU\\TestKey\\", 1)

R.CurrentKey = "HKCU\\TestKey\\"
assign(R.CurrentKey, "HKCU\\TestKey\\")
var sk = R.SubKeys;
assign(sk.Count(), 0)

close() // чтобы могли бы удалить созданные ключи
RegDelete("HKCU\\TestKey\\")
