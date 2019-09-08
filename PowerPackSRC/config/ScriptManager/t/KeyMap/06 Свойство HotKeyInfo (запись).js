var CtrlQ = new Array(9)
CtrlQ[0] = { script	: "jsSMTest", macros : "NoWindow"	}
CtrlQ[4] = { script	: "jsSMTest", macros : "OpenConfMDTab"	}
CtrlQ[8] = { script	: "jsSMTest", macros : "UnknownWindow" }

var CrLf = "\r\n"
var res = '//keymap file, version=1.0' + CrLf
		+ 'hk.CtrlQ = new Array(9)' + CrLf
		+ 'hk.CtrlQ[0] = {' + CrLf
		+ '	"script" : "jsSMTest",' + CrLf 
		+ '	"macros" : "NoWindow"' + CrLf
		+ '}' + CrLf
		+ 'hk.CtrlQ[4] = {' + CrLf
		+ '	"script" : "jsSMTest",' + CrLf 
		+ '	"macros" : "OpenConfMDTab"' + CrLf
		+'}' + CrLf
		+ 'hk.CtrlQ[8] = {' + CrLf
		+ '	"script" : "jsSMTest",' + CrLf 
		+ '	"macros" : "UnknownWindow"' + CrLf
		+'}' + CrLf;


KM.HotKeyInfo("CtrlQ") = CtrlQ
assign(KM.Stringify(), res)
