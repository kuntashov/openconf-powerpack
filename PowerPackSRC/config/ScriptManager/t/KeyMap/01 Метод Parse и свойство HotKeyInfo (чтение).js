var src = 'hk.CtrlQ = new Array(9);'
		+ 'hk.CtrlQ[0] = { script : "jsSMTest", macros	: "NoWindow"		};'
		+ 'hk.CtrlQ[1] = { script : "jsSMTest", macros	: "Metadata"		};'
		+ 'hk.CtrlQ[2] = { script : "jsSMTest", macros	: "Interfaces"		};'
		+ 'hk.CtrlQ[3] = { script : "jsSMTest", macros	: "Rights"			};'
		+ 'hk.CtrlQ[4] = { script : "jsSMTest", macros	: "OpenConfMDTab"	};'
		+ 'hk.CtrlQ[5] = { script : "jsSMTest", macros	: "TextEditor" 		};'
		+ 'hk.CtrlQ[6] = { script : "jsSMTest", macros	: "DialogEditor"	};'
		+ 'hk.CtrlQ[7] = { script : "jsSMTest", macros	: "TableEditor"		};'
		+ 'hk.CtrlQ[8] = { script : "jsSMTest", macros	: "UnknownWindow"	};';

assign(KM.Parse(src), true)

assign(KM.Keys().length, 1);
assign(KM.HotKeyInfo("CtrlQ").length, 9)
assign(KM.HotKeyInfo("CtrlQ", 0).macros, "NoWindow")
// в принципе, мы не JScript Engine тестируем, 
// но все же для успокоения совести 
assign(KM.HotKeyInfo("CtrlQ", 1).macros, "Metadata")
assign(KM.HotKeyInfo("CtrlQ", 2).macros, "Interfaces")
assign(KM.HotKeyInfo("CtrlQ", 3).macros, "Rights")
assign(KM.HotKeyInfo("CtrlQ", 4).macros, "OpenConfMDTab")
assign(KM.HotKeyInfo("CtrlQ", 5).macros, "TextEditor")
assign(KM.HotKeyInfo("CtrlQ", 6).macros, "DialogEditor")
assign(KM.HotKeyInfo("CtrlQ", 7).macros, "TableEditor")
assign(KM.HotKeyInfo("CtrlQ", 8).macros, "UnknownWindow")