var msg = function (str) { WScript.Echo(str) }

var assign = function (a, b) { aNo++; if (a!=b) throw(new Error(0, "Assignment failed: " + a + " != " + b))  }

var aNo = 0;

msg('----------------------------------------------');
msg('    Testing ScriptManager.KeyMap component    ');
msg('----------------------------------------------');

var Counter = 0;
var Failed = 0;

var mask = null;
if (WScript.Arguments.Count() > 0) {
	mask = new RegExp(WScript.Arguments(0));
}

var fso = new ActiveXObject("Scripting.FileSystemObject");
var dir = fso.GetFolder(".\\KeyMap");
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
		var KM = new ActiveXObject("ScriptManager.KeyMap");
		aNo = 0;

		eval(test.OpenAsTextStream().ReadAll());

		msg("OK.");
	}
	catch (e) {
		Failed++;
		msg('FAILED (' + aNo + '). Reason: '+ e.description);
	}
}

msg('----------------------------------------------');
msg('Tests passed: ' + (Counter - Failed) + "/" + Counter);
msg('Tests failed: ' + Failed + "/" + Counter);
