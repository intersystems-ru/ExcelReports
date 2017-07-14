# ExcelReports
ExcelReports from Caché
## Инструкция

1. В таблице excel, заполните поля, значения которых вы хотите заменить на данные английскими буквами **qNcM**, где N - номер набора данных, а M - номер колонки в наборе данных. Либо **qNaM**, где N - номер набора данных, а M - номер аргумента в наборе данных.
![Image of Yaktocat](http://savepic.ru/14518749.jpg)
2. В пустых строках между необходимо поставить любой символ, например пробел или изменить форматирование (не должно быть полностью пустых строк). 
3. Сохраните отчет в формате **Таблица XML 2003** (не XML-Данные).
4. Импортируйте класс **Excel.XSL.xml** в нужную облость, например **Samples**. 
5. Вызовите метод **XSLtoFile(InputFileName, OutputFileName, Queries** класса Excel.XSL, пример вызова доступен в Excel.Test:
  ```
ClassMethod Test()
{
	// Fill global with test data
	kill ^ExcelTest
	set ^ExcelTest(1)    = $ListBuild("Argument1", "Argument2")
	set ^ExcelTest(1,1) = $ListBuild("Value11", "Value12")
	set ^ExcelTest(1,2) = $ListBuild("Value21", "Value22")
	set ^ExcelTest(1,3) = $ListBuild("Value31", "Value32")
	
	set ^ExcelTest(2)    = $ListBuild("Argument1", "Argument2")
	set ^ExcelTest(2,1) = $ListBuild("Value11", "Value12")
	set ^ExcelTest(2,2) = $ListBuild("Value21", "Value22")
 	set ^ExcelTest(2,3) = $ListBuild("Value31", "Value32")
	
	// Specify datasources
	set Queries(1) = "SELECT Id FROM Sample.Person WHERE Id>? AND Id<?"
	set Queries(1,1) = 1
	set Queries(1,2) = 10
	set Queries(2) = "^ExcelTest(1)"
	set Queries(3) = "^ExcelTest(2)"
	
	// Specify input and output files
	set XMLFile = "D:\Cache\ExcelReports\Test\Source.xml"
	set XLSFile = "D:\Cache\ExcelReports\Test\Out.xls"
	
	quit ##class(Excel.XSL).XSLtoFile(XMLFile, XLSFile, .Queries)
}
  ```
  
  6. Результат:
  ![Image of Yaktocat](http://savepic.ru/14494195.jpg)
