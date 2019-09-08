RegWrite("HKCU\\TestKey\\Value", "Hello, World!")

R.CurrentKey = "HKCU\\TestKey"
assign(R.Values(0), "Hello, World!")
assign(R.Values("Value"), "Hello, World!")

var vals = R.Values
assign(vals.Item(0), "Hello, World!")
assign(vals.Item("Value"), "Hello, World!")

RegDelete("HKCU\\TestKey\\Value")
RegDelete("HKCU\\TestKey\\")