RegWrite("HKCU\\TestKey\\SubKey00\\", 1)
RegWrite("HKCU\\TestKey\\SubKey01\\", 2)
RegWrite("HKCU\\TestKey\\SubKey02\\", 3)
RegWrite("HKCU\\TestKey\\SubKey03\\", 4)
RegWrite("HKCU\\TestKey\\SubKey04\\", 5)

R.CurrentKey = "HKCU\\TestKey\\"

for(var i=0; i<R.SubKeysCount(); i++) {
	assign(R.SubKeyName(i), "SubKey0" + i)
}

close() // чтобы могли бы удалить созданные ключи
RegDelete("HKCU\\TestKey\\SubKey00\\")
RegDelete("HKCU\\TestKey\\SubKey01\\")
RegDelete("HKCU\\TestKey\\SubKey02\\")
RegDelete("HKCU\\TestKey\\SubKey03\\")
RegDelete("HKCU\\TestKey\\SubKey04\\")
RegDelete("HKCU\\TestKey\\")
