RegWrite("HKCU\\TestKey\\Value", "Hello, World!")

R.CurrentKey = "HKCU\\TestKey"
assign(R.Values.Name(0), "Value")

RegDelete("HKCU\\TestKey\\Value")
RegDelete("HKCU\\TestKey\\")