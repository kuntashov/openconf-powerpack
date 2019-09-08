assign(R.BuildPath("HKEY_CURRENT_USER\\TestKey", "SubKey"), "HKEY_CURRENT_USER\\TestKey\\SubKey")
assign(R.BuildPath("HKEY_CURRENT_USER\\TestKey", "\\SubKey"), "HKEY_CURRENT_USER\\TestKey\\SubKey")
assign(R.BuildPath("HKEY_CURRENT_USER\\TestKey\\", "\\SubKey"), "HKEY_CURRENT_USER\\TestKey\\SubKey")
