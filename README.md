# ExcelReports
ExcelReports from Caché
## Инструкция

1. В таблице excel, заполните поля, значения которых вы хотите заменить на данные английскими буквами **qNcM**, где N - номер набора данных, а M - номер колонки в наборе данных. Либо **qNaM**, где N - номер набора данных, а M - номер аргумента в наборе данных.
![Image of Yaktocat](http://savepic.ru/14822127.jpg)
2. В пустых строках между необходимо поставить любой символ, например пробел или изменить форматирование (не должно быть полностью пустых строк). 
3. Сохраните отчет в формате **Таблица XML 2003** (не XML-Данные).
4. Импортируйте класс **Excel.XSL.xml** в нужную облость, например **Samples**. 
5. Вызовите метод **XSLtoFile(InputFileName, OutputFileName, Queries** класса Excel.XSL, пример вызова доступен в Excel.Test:
  ```
ClassMethod Test()
{
	// Fill global with test data
  kill ^ExcelTest
  set ^ExcelTest(1)    = $ListBuild("47", "49")
  set ^ExcelTest(1,1) = $ListBuild("47", "RoboSoft Media Inc.", "69926933")
  set ^ExcelTest(1,2) = $ListBuild("48", "OctoTech Corp.", "679240524")
  set ^ExcelTest(1,3) = $ListBuild("49", "HyperGlomerate Gmbh.", "530112065")
  
  // Specify datasources
  set Queries(1) = "SELECT ID, COMPANY, ""DATE"", NAME, NUMBER, PHONE, SSN FROM Tests.Data WHERE ID>=? AND ID<=?"
  set Queries(1,1) = 47
  set Queries(1,2) = 49
  set Queries(2) = "^ExcelTest(1)"
  
  // Specify input and output files
  set XMLFile = "C:\temp\ExcelReports\input.xml"
  set XLSFile = "C:\temp\ExcelReports\Out.xls"
  
  quit ##class(Excel.XSL).XSLtoFile(XMLFile,XLSFile, .Queries)
}
  ```
  
  6. Результат:
  
  ![Image of Yaktocat](http://savepic.ru/14824175.jpg)
