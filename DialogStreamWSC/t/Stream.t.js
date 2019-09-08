eval(include("../DialogStream.js"));

function include(fileName)
{
    with ( new ActiveXObject("Scripting.FileSystemObject") )
        return OpenTextFile(fileName).ReadAll();    
}

function  print(str) {
    WScript.Echo(str);
}

function assign(result) {
    return result ? 'Ok' : 'Failed';
}

try {
    var s = new Stream();

    s.read('{"Test1"}');  
    print(assign(s.item(0).toString()==='Test1'));
  

    s.read('{"Test2","1","2","3"}');
    print(assign(s.length()===4));

    var a = ['Test2', 1, 2, 3];
    for(var i=0; i<s.length(); i++) {
        print(assign(s.item(i).toString()===a[i].toString()));
    }           

    s.read('{"Main stream", {"Child stream", "Data"},"End"}');
    print(assign(s.item(0).toString()==='Main stream'));
    print(assign(s.item(1).item(0).toString() === 'Child stream'));
    print(assign(s.item(1).item(1).toString() === 'Data'));
    print(assign(s.item(2).toString() === 'End'));

    s.read('{}');
    print(assign(s.length()===0));

    s.read('{{}}');
    print(assign(s.length()===1));
    print(assign(s.item(0).length()===0));
}
catch (e) {
    print (e);
}

