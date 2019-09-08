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

var a = new Stream();

print('Test 1. Initial length ' + assign(a.length() === 0));

print('Test 2. ArrayEx::pushItem() method');
print("\t pushItem return: " + assign(a.pushItem(12345) === 12345));
print("\t new length: " + assign(a.length() === 1));
print("\t pushItem return: " + assign(a.pushItem('qwerty') === 'qwerty'));
print("\t new length: " + assign(a.length() === 2));

print('Test 3. ArrayEx::item() method, part 1');
print("\t item(index) return: " + assign(a.item(0) === 12345));
print("\t item(index) return: " + assign(a.item(1) === 'qwerty'));
print("\t item(index) - index out of the bounds (1): " + assign(a.item(-10) === undefined));
print("\t item(index) - index out of the bounds (2): " + assign(a.item(100) === undefined));

print('Test 3. ArrayEx::item() method, part 2');
print("\t item(index, newValue) -- set new value: " + assign(a.item(0, 666) === 12345));
print("\t item(index) -- check new value: " + assign(a.item(0) === 666));


