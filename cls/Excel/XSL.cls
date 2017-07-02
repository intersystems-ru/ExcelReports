Class Excel.XSL
{

/// Входной файл
Parameter XMLFILE = "C:\temp\input.xml";

/// Выходной файл
Parameter XLSFILE = "C:\temp\output.xls";

/// Преобразование Excel таблицы в формате XML 2003 в XSL шаблон и наполнение его данными
/// На вход: 
/// XMLFile - Полный путь до исходного файла шаблона Excel таблицы в формате XML 2003 
/// XMLFile - Полный путь до выходного файла Excel таблицы с расширением xls
/// Query - SQL запрос
/// SQLArgs... - параметры SQL запроса
/// w $System.Status.GetErrorText(##class(Excel.XSL).XSLtoFile(,,"SELECT Id FROM Sample.Person WHERE Id>?", 50))
ClassMethod XSLtoFile(XMLFile As %String = {..#XMLFILE}, XLSFile As %String = {..#XLSFILE}, Query As %String = "", SQLArgs...) As %Status
{
	#dim Status As %Status = $$$OK
	set XML = ##class(%FileCharacterStream).%New()
	set XML.Filename = XMLFile
	
	set XSL = ..XMLToXSL(XML, .Status)
	quit:$$$ISERR(Status) Status  
	
	set Status = ..GenXMLData(.XMLData, Query, SQLArgs...)
	quit:$$$ISERR(Status) Status  
	
	set OutputFile = ##class(%FileCharacterStream).%New()
	set OutputFile.TranslateTable = "UTF8"
	set OutputFile.Filename = XLSFile
	set Status = ##class(%XML.XSLT.Transformer).TransformStream(XMLData, XSL, .OutputFile)
	quit:$$$ISERR(Status) Status  
	
	quit OutputFile.%Save()
}

/// Генерация XML с данными
ClassMethod GenXMLData(Stream As %Stream.Object = {##class(%GlobalCharacterStream).%New()}, Query As %String, SQLArgs...) As %Status
{
	#dim Status As %Status = $$$OK
	
	set Writer = ##class(%XML.Writer).%New()
	set Writer.Indent = 1
	set Status = Writer.OutputToStream(Stream)
	quit:$$$ISERR(Status) Status
	do Writer.RootElement("SQLResult")
	set Status = ..AddSQLTags(.Writer, Query, SQLArgs...)	
 	do Writer.EndRootElement()
	quit Status
}

ClassMethod AddSQLTags(Writer As %XML.Writer, Query As %String, SQLArgs...) As %Status
{
	#dim Status As %Status = $$$OK
	
	#dim RS As %SQL.StatementResult = ##class(%SQL.Statement).%ExecDirect(,Query, SQLArgs...)
	quit:RS.%SQLCODE'=0 $$$ERROR($$$SQLError, RS.%SQLCODE, RS.%Message)

	set ColCount = RS.%ResultColumnCount
	
	#dim Metadata As %SQL.StatementMetadata = RS.%GetMetadata() 
	for i = 1:1:ColCount {
		set ColTypes(i) = Metadata.columns.GetAt(i).clientType
	}
	
	while RS.%Next(.Status)
	{ 
		quit:$$$ISERR(Status)
		do Writer.Element("row")
		for i = 1:1:ColCount {
			if (ColTypes(i) = 2) {
				do ..AddTagElement(.Writer, "c"_i, $zd(RS.%GetData(i), 3))
			} else {
				do ..AddTagElement(.Writer, "c"_i, RS.%GetData(i))
			}
		}
		do Writer.EndElement()	
	}
	quit Status
}

ClassMethod AddTagElement(Writer As %XML.Writer, TagName As %String, TagValue As %String) As %Status
{
	#dim Status As %Status = $$$OK
	do Writer.Element(TagName)  
	do Writer.Write(TagValue)  
	do Writer.EndElement()
	quit Status
}

/// Преобразование исходного XML файла в XSL шаблон, на вход которому подаются данные.
ClassMethod XMLToXSL(XML As %Stream.FileCharacter, ByRef Status As %Status) As %GlobalBinaryStream
{
	set Status = $$$OK
	set XSL = ##class(%Dictionary.CompiledXData).IDKEYOpen($classname(), "XSLData").Data
	set Status = ##class(%XML.XSLT.Transformer).TransformStream(XML, XSL, .ResultXSL)
	return ResultXSL
}

ClassMethod Test() As %Status
{
	#dim Status As %Status = $$$OK
	set XML = ##class(%FileCharacterStream).%New()
	set XML.Filename = ..#XMLFILE
	
	write $zt($p($h,",",2)), " Конвертирование XML в XSL: ...", !
	set XSL = ..XMLToXSL(XML,.Status) 
	quit:$$$ISERR(Status) Status
	write $zt($p($h,",",2)), " Конвертирование XML в XSL: OK!", !
	write XSL.%Save()
   	quit Status
}

XData XSLData
{
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:x="urn:schemas-microsoft-com:office:excel"
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
	xmlns:out="http://xmlportfolio.com/dummy"
	version="1.0">
<xsl:namespace-alias stylesheet-prefix="out" result-prefix="xsl"/>
<xsl:output method="xml" omit-xml-declaration="no" indent="yes" />
<xsl:variable name="vDigits" select="'0123456789'"/>
<xsl:template match="/">
<out:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<out:output method="xml" omit-xml-declaration="no" indent="yes" />		
<out:template match="/">
<out:processing-instruction name="mso-application">
<out:text>progid="Excel.Sheet"</out:text>
</out:processing-instruction>
<xsl:copy>
<xsl:apply-templates select="@*"/>
<xsl:apply-templates select="node()"/>
</xsl:copy>
</out:template>
</out:stylesheet>
</xsl:template>
<xsl:template match="//@* | //node()">
<xsl:copy>
<xsl:apply-templates select="@*"/>
<xsl:apply-templates select="node()"/>
</xsl:copy>
</xsl:template>
<xsl:template match="//*[local-name() = 'Table']/@ss:ExpandedColumnCount"/>
<xsl:template match="//*[local-name() = 'Table']/@ss:ExpandedRowCount"/>
<xsl:template match="//*[local-name() = 'Row']/@ss:Index"/>
<xsl:template match="//*[local-name() = 'Row']">
<xsl:choose>
<xsl:when test="(.//*[local-name() = 'Data' and (starts-with(.,'c') = 'true') and (translate(substring-after(.,'c'),$vDigits,'') = '')])">
<out:for-each select="//row">
<xsl:copy>
<xsl:apply-templates select="node()|@*"/>
</xsl:copy>
</out:for-each>
</xsl:when>
<xsl:otherwise>
<xsl:copy>
<xsl:apply-templates select="node()|@*"/>
</xsl:copy>
</xsl:otherwise>
</xsl:choose>
</xsl:template>
<xsl:template match="//*[local-name() = 'Data' and count(*) = 0]">
<xsl:element name="{local-name()}" namespace="{namespace-uri()}">
<xsl:copy-of select="./namespace::*"/>
<xsl:apply-templates select="@*"/>
<xsl:choose>
<xsl:when test="(starts-with(.,'c') = 'true') and (translate(substring-after(.,'c'),$vDigits,'') = '')">
<xsl:element name="xsl:value-of">
<xsl:attribute name="select">
<xsl:value-of select="."/>
</xsl:attribute>
</xsl:element>
</xsl:when>
<xsl:otherwise>
<xsl:value-of select="."/>
</xsl:otherwise>
</xsl:choose>
</xsl:element>
</xsl:template>
</xsl:stylesheet>
}

}

