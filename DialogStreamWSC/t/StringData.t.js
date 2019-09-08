function  print(str) {
    WScript.Echo(str);
}

function assign(result) {
    return result ? 'Ok' : 'Failed';
}

function StringData()
{
    this.src = null;
    this.data = '';
}

StringData.prototype.assignIndex = function (index)
{
    return ((this.src)&&(-1 < index)&&(index<this.src.length));    
}

StringData.prototype.charAt = function (index)
{  
    if (this.assignIndex(index)) {
        return this.src.charAt(index);
    }
    return undefined;
}

StringData.prototype.read = function (rawData)
{
    var index = 0;
    this.src = rawData;
    this.data = '';

    if (this.charAt(index++) == '"') {
        while (ch = this.charAt(index)) {
            if (ch == '"') {
                index++;
                if (this.charAt(index) == '"') {
                    this.data += '"';
                    index++;
                } else {
                    this.src = this.src.substr(0, index);
                    return (index);
                }
            } 
            else {
                this.data += ch;
                index++;    
            }
        }
    }
    else {
        throw("StringData::read(): Синтаксическая ошибка (" + index + ")!");
    }
    throw("StringData::read(): Неожиданный конец входной строки!");
}

try {
    var s = new StringData();

    print(assign(s.read('"test string"') == 13));
    print(assign(s.data == 'test string'));

    print(assign(s.read('""""') == 4));
    print(assign(s.data == '"'));
         
    print(assign(s.read('"""Test"""') == 10));
    print(assign(s.data == '"Test"'));

    print(assign(s.read('"""first line""\r\n""second line"""') == 33));
    print(assign(s.data == '"first line"\r\n"second line"'));

}
catch (e) {
    print (e);
}




















