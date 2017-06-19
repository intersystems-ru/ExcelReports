# ExcelReports
ExcelReports from Caché
## Инструкция

1. Заполните поля в таблице excel английской буквой **c** и цифрой от 1 до n. Это колонки в sql запросе.
![Image of Yaktocat](http://savepic.ru/14518749.jpg)

2. Сохраните отчет в формате **Таблица XML 2003**.
3. Откройте сохраненный файл в любом редакторе и вставьте теги 
```<xsl:for-each select="//row">```
```</xsl:for-each>``` 
до и после нужного нам тега **Row**.
![Image of Yaktocat](http://savepic.ru/14524895.jpg)
4. Если после закрывающего тега ```</xsl:for-each>``` есть теги **Row** или **Cell** с аттрибутом **ss:Index**, то необходимо удалить атрибут и в теле тега прописать следующее: 
```
<xsl:attribute name="ss:Index"> 
   <xsl:value-of select="M+$CountParam"/>
</xsl:attribute>
  ```
  Где `M` это значение, которое было в атрибуте **ss:Index**. После внесения изменений сохраните файл и закройте его.
  5. Импортируйте класс **Excel.XSL.xml** в нужную облость, например **Sample**.После вызовите метод **XSLtoFile(Query,InputFileName,OutputFileName)**:
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
