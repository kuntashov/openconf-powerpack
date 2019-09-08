function StringData()
{
    this.src = null;
    this.data = '';
}

StringData.prototype.toString = function ()
{
    return this.data.toString();
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
        throw("StringData::read(): Syntax error (" + index + ")!");
    }
    throw("StringData::read(): Unexpected end of input data!");
}

function Stream() 
{
    this.src = null;
    
    // private
    var items = new Array();
    
    // public 
    this.item = function (index, newValue) {                  
        if ((-1 < index)&&(index < items.length)) {
            var ret = items[index];
            if ((newValue !== undefined)) {
                items[index] = newValue;
            }            
            return ret;
        }
        return undefined;
    }

    this.pushItem = function (item) {        
        return (items[items.length] = item);
    }

    this.length = function () {
        return items.length;
    }

    this.clearItems = function () {
        items = new Array();
    }
}

Stream.prototype.assignIndex = function (index)
{
    return ((this.src)&&(-1 < index)&&(index<this.src.length));    
}

Stream.prototype.charAt = function (index)
{  
    if (this.assignIndex(index)) {
        return this.src.charAt(index);
    }
    return undefined;
}

Stream.prototype.skipSpaces = function (index)
{        
    var count = 0;
    while (this.assignIndex(index)) {
        if (!(this.src.charAt(index++).match(new RegExp("\\s")))) {
            break;
        }        
        count++;
    } 
    return count;
}

Stream.prototype.read = function (rawData, ignoreTail)
{ 
    var index = 0;    
    this.src = rawData;
    this.clearItems();

    index += this.skipSpaces(index);
    
    if (this.charAt(index) == '{') {
        index++;        
        index = this.parse(index);
        if (!ignoreTail) {                  
            index += this.skipSpaces(index);   
            if (index !== this.src.length) {
                throw("Stream::read(): Syntax error (" + index + ")!");
            }
        }
        this.src = this.src.substr(0, index);
        return (index);
    } 
    else {    
        throw("Stream::read(): Unexpected char '" 
                + this.charAt(index) + "' (" + index + ")!");
    }   
}

Stream.prototype.parse = function (index)
{
    var ch, item;

    while (ch = this.charAt(index)) {        
        switch (ch) {
            case '"':
                item = this.pushItem(new StringData());                
                index += item.read(this.src.substr(index));                                
                break;
            case '{':
                item = this.pushItem(new Stream());
                index += item.read(this.src.substr(index), true);                
                break;                
            case '}':
                return (index+1);
            case ',':
                index++;
            default:                
                index += this.skipSpaces(index);            
                ch = this.charAt(index);
                if (ch === undefined) {
                    throw("Stream::parse(): Unexpexted end of input data!");
                } 
                else if ((ch != '{')&&(ch != '"')) {
                    throw("Stream::parse(): Unexpexted char '" 
                        + this.charAt(index) + "' (" + index + ")!" + ch);                                        
                }
                break;                
        }
    }
    throw("Stream::parse(): Unexpexted end of input data!");
}


