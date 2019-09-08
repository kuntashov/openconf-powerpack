// Загружаем наш "псевдо-обработчик" нажатий клавиш-хоткеев
Scripts.Load(BinDir + "config\\system\\ScriptManager\\HotKeysHandler.js")
// Скрипт для редактирования хоткеев (потом сделаем загрузку по команде)
Scripts.Load(BinDir + "config\\system\\ScriptManager\\HotKeysEditor.js")
// ...и выгружаем себя за ненадобностью
Scripts.UnLoad(SelfScript.Name)