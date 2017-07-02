# ExcelReports
ExcelReports from Caché
## Инструкция

1. Заполните поля в таблице excel английской буквой **c** и цифрой от 1 до n. Это колонки в sql запросе.
![Image of Yaktocat](http://savepic.ru/14518749.jpg)
2. Если у вас после результата выполнения SQL запроса идут другие данные (футер), то в строках между ними необходимо поставить любой символ, например пробел или изменить форматирование (не должно быть полностью пустых строк). 
3. Сохраните отчет в формате **Таблица XML 2003**.
4. Импортируйте класс **Excel.XSL.xml** в нужную облость, например **Sample**.После, вызовите метод **XSLtoFile(Query,InputFileName,OutputFileName)**:
  ```
  set Query = "SELECT Name, Age, FavoriteColors, DOB,  SSN, Home_City, Home_State, Home_Street 
               FROM Sample.Person 
               WHERE Age > ? AND Age < ? 
               ORDER BY Age"
  set InputFileName = "C:\temp\input.xml"
  set OutputFileName = "C:\temp\output.xml"
  set Status = ##class(Excel.XSL).XSLtoFile(InputFileName, OutputFileName, Query, 26, 49)
  write $System.Status.GetErrorText(Status)
  ```
  6. Результат:
  ![Image of Yaktocat](http://savepic.ru/14494195.jpg)
