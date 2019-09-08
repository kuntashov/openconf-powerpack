var msg = function (str) { WScript.Echo(str) }

var assign = function (a, b) { aNo++; if (a!=b) throw(new Error(0, "Assignment failed: " + a + " != " + b))  }

var wsh = new ActiveXObject("WScript.Shell");
var RegRead 	= function (strName) 					{ return wsh.RegRead(strName) }
var RegWrite	= function (strName, strValue, Type)	{ wsh.RegWrite(strName, strValue, Type?Type:"REG_SZ") }
var WriteDWORD	= function (strName, strValue)			{ wsh.RegWrite(strName, strValue, "REG_DWORD") }
var RegDelete	= function (strName)					{ wsh.RegDelete(strName) }

var aNo = 0;

msg('--------------------------------------------------------');
msg('    Testing ScriptManager.KeyMapsContainer component    ');
msg('--------------------------------------------------------');

var Counter = 0;
var Failed = 0;

var mask = null;
if (WScript.Arguments.Count() > 0) {
	mask = new RegExp(WScript.Arguments(0));
}

var fso = new ActiveXObject("Scripting.FileSystemObject");
var dir = fso.GetFolder(".\\KeyMapsCnt");
var tests = new Enumerator(dir.Files);

for(; !tests.atEnd(); tests.moveNext()) {
	var test = tests.item();
	if (mask instanceof RegExp) {
		if (!mask.test(test.Name)) {
			continue;
		}
	}
	Counter++
	msg("Running test " + test.Name);
	try {
		var KMC = new ActiveXObject("ScriptManager.KeyMapsContainer");
		aNo = 0;

		eval(test.OpenAsTextStream().ReadAll());

		msg("OK.");
	}
	catch (e) {
		Failed++;
		msg('FAILED (' + aNo + '). Reason: '+ e.description);
	}
}

msg('--------------------------------------------------------');
msg('Tests passed: ' + (Counter - Failed) + "/" + Counter);
msg('Tests failed: ' + Failed + "/" + Counter);
