// Путь относительно раздела
R.CurrentKey = "\\TestKey"
assign(R.CurrentKey, "HKEY_CURRENT_USER\\TestKey")
assign(R.KeyName, "TestKey")
// Путь относительно текущего ключа
R.OpenKey("SubKey")
assign(R.CurrentKey, "HKEY_CURRENT_USER\\TestKey\\SubKey")
assign(R.KeyName, "SubKey")
