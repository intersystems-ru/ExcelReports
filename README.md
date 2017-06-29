# ExcelReports
ExcelReports from Caché
## Инструкция

1. Заполните поля в таблице excel английской буквой **c** и цифрой от 1 до n. Это колонки в sql запросе.
![Image of Yaktocat](http://savepic.ru/14518749.jpg)
2. Если у вас больше одной таблицы, то в строках между ними необходимо поставить символ, например пробел. Это строка с симвалом будет служить ограничителем между таблицами.
3. Сохраните отчет в формате **Таблица XML 2003**.
4. Импортируйте класс **Excel.XSL.xml** в нужную облость, например **Sample**.После, вызовите метод **XSLtoFile(Query,InputFileName,OutputFileName)**:
  ```
  set Query = "SELECT Name, Age, FavoriteColors, DOB,  SSN, Home_City, Home_State, Home_Street 
               FROM Sample.Person 
               WHERE Age >18 AND Age < 40 
               ORDER BY Age"
  set InputFileName = "C:\temp\input.xml"
  set OutputFileName = "C:\temp\output.xml"
  do ##class(Excel.XSL).XSLtoFile(Query,InputFileName,OutputFileName)
  ```
  6. Результат:
  ![Image of Yaktocat](http://savepic.ru/14494195.jpg)
