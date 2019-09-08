перед использованием не забудте зарегистрировать
regsvr32 Enumerator.wsc

Пример использования

     fc = СоздатьОбъект("Scripting.Enumerator");

     fso = СоздатьОбъект("Scripting.FileSystemObject");

     f = fso.GetFolder("D:\\Temp\\");

     fc.enumCollection(f.files);     // этот по сути и есть new Enumerator

     Пока (fc.atEnd()=0) Цикл
        Сообщить();(fc.item().name);
        fc.moveNext();
     КонецЦикла
