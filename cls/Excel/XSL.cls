Class Excel.XSL
{

/// Входной файл
Parameter XMLFILE = "C:\temp\input.xml";

/// Выходной файл
Parameter XLSFILE = "C:\temp\output.xls";

/// Преобразование Excel таблицы в формате XML 2003 в XSL шаблон и наполнение его данными
/// На вход: 
/// XMLFile - Полный путь до исходного файла шаблона Excel таблицы в формате XML 2003 
/// XLSFile - Полный путь до выходного файла Excel таблицы с расширением xls
/// Queries - SQL запросы и глобалы в формате:
/// Queries(n) = <SQL Запрос> или <Глобал> номер n
/// Queries(n, m) = параметр SQL запроса номер m
/// w $System.Status.GetErrorText(##class(Excel.XSL).XSLtoFile(,,.Queries))
ClassMethod XSLtoFile(XMLFile As %String = {..#XMLFILE}, XLSFile As %String = {..#XLSFILE}, ByRef Queries As %String) As %Status
{
	#dim Status As %Status = $$$OK
	
	set Status = ##class(%File).Exists(XMLFile)
	quit:$$$ISERR(Status) $$$ERROR($$$FileDoesNotExist, XMLFile)	
	
	set XML = ##class(%FileBinaryStream).%New()
	set XML.Filename = XMLFile
	
	set XSL = ..XMLToXSL(XML, .Status)
	quit:$$$ISERR(Status) Status  
	
	set Status = ..GenXMLData(.XMLData, .Queries)
	quit:$$$ISERR(Status) Status  

	set OutputFile = ##class(%FileCharacterStream).%New()
	set OutputFile.TranslateTable = "UTF8"
	set OutputFile.Filename = XLSFile
	set Status = ##class(%XML.XSLT.Transformer).TransformStream(XMLData, XSL, .OutputFile)
	quit:$$$ISERR(Status) Status  
	
	quit OutputFile.%Save()
}

ClassMethod GenXMLData(Stream As %Stream.Object = {##class(%GlobalCharacterStream).%New()}, Queries As %String) As %Status
{
	#dim Status As %Status = $$$OK
	
	set Writer = ##class(%XML.Writer).%New()
	set Writer.Indent = 1
	set Status = Writer.OutputToStream(Stream)
	quit:$$$ISERR(Status) Status
	
	do Writer.RootElement("data")
	
	set Iter = $Order(Queries(""))
	while (Iter '= "")
	{
		if $Length($ZConvert(Queries(Iter), "U"), "SELECT") > 1
		{
			merge QueryNode =  Queries(Iter)
			do ..GenXMLDataBySQL(.Writer, .QueryNode, Iter)
			kill QueryNode
			
		}elseif $ListValid(@Queries(Iter)){
			merge GlobalNode =  @Queries(Iter)
			do ..GenXMLDataByGlobal(.Writer, .GlobalNode, Iter)
			kill GlobalNode
		}else{
			return $$$ERROR($$$GeneralError, "Unexpected query: " _ Queries(Iter))
		}
		set Iter = $Order(Queries(Iter))
	}
	do Writer.EndRootElement()
 	
	quit Status
}

ClassMethod GenXMLDataByGlobal(Writer As %XML.Writer, GlobalNode As %String, NumGlobalNode As %String) As %Status
{
	#dim Status As %Status = $$$OK
	
	do Writer.Element("query")
	do Writer.WriteAttribute("id", NumGlobalNode)
	do Writer.WriteAttribute("type", "Global")
	
	if $ListValid(GlobalNode){
		set Status = ..AddGlobalArgument(.Writer, GlobalNode, "q" _ NumGlobalNode)
	}
	set Status = ..AddGlobalTags(.Writer, .GlobalNode, "q" _ NumGlobalNode)
	
	do Writer.EndElement()
	
	quit Status
}

ClassMethod AddGlobalArgument(Writer As %XML.Writer, GlobalArgss As %String, NumGlobalNode As %String) As %Status
{
	#dim Status As %Status = $$$OK
	
	do Writer.Element("args")
	for i=1:1:$ListLength(GlobalArgss)
	{
		set Status = ..AddTagElement(.Writer, NumGlobalNode _ "a" _ i, $ListGet(GlobalArgss, i))
	}
	do Writer.EndElement()
	
	return Status
}

ClassMethod AddGlobalTags(Writer As %XML.Writer, GlobalNodes As %String, NumGlobalNode As %String) As %Status
{
	#dim Status As %Status = $$$OK
	
	set Node = $Order(GlobalNodes(""))
	while (Node '= "")
	{
		do Writer.Element("row")
		for NodeList = 1:1:$ListLength(GlobalNodes(Node))
		{
			set Status = ..AddTagElement(.Writer, NumGlobalNode _ "c" _ NodeList, $ListGet(GlobalNodes(Node), NodeList))
		}
		set Node = $Order(GlobalNodes(Node))
		do Writer.EndElement()
	}
	
	return Status
}

/// Генерация XML с данными по SQL запрсам
ClassMethod GenXMLDataBySQL(Writer As %XML.Writer, Query As %String, NumQuery As %String) As %Status
{
	#dim Status As %Status = $$$OK
	
	do Writer.Element("query")
	do Writer.WriteAttribute("id", NumQuery)
	do Writer.WriteAttribute("type","SQL")
	
	set ArgsCount = ..GetSQLArgsCount(.Query)
	
	if ArgsCount > 0 {
		merge SQLArgs = Query
		set SQLArgs = ArgsCount
		set Status = ..AddSQLAttributes(.Writer, .SQLArgs, "q" _ NumQuery)
	}
	set Status = ..AddSQLTags(.Writer, Query,"q"_NumQuery, SQLArgs...)
	
	do Writer.EndElement()
	
	kill SQLArgs
	quit Status
}

ClassMethod AddSQLAttributes(Writer As %XML.Writer, SQLArgs, NumQuery As %String) As %Status
{
	set SC = $$$OK
	set SC = Writer.Element("args")
	set Iter = $Order(SQLArgs(""))
	while (Iter '= "")
	{
		set SC = Writer.Element(NumQuery_"a"_Iter)
		set SC = Writer.Write(SQLArgs(Iter))
		set SC = Writer.EndElement()
		set Iter = $Order(SQLArgs(Iter))
	}
	set SC = Writer.EndElement()
	return $$$ISERR(SC)
}

ClassMethod AddSQLTags(Writer As %XML.Writer, Query As %String, NumQuery As %String, SQLArgs...) As %Status
{
	#dim Status As %Status = $$$OK
	#dim RS As %SQL.StatementResult = ##class(%SQL.Statement).%ExecDirect(,Query, SQLArgs...)
	return:RS.%SQLCODE'=0 $$$ERROR($$$SQLError, RS.%SQLCODE, RS.%Message)

	set ColCount = RS.%ResultColumnCount
	
	#dim Metadata As %SQL.StatementMetadata = RS.%GetMetadata() 
	for i = 1:1:ColCount {
		set ColTypes(i) = Metadata.columns.GetAt(i).clientType
	}
	
	while RS.%Next(.Status)
	{ 
		quit:$$$ISERR(Status)
		do Writer.Element("row")
		for i = 1:1:ColCount 
		{
			if (ColTypes(i) = 2) {
				do ..AddTagElement(.Writer, NumQuery _ "c" _ i, $ZDate(RS.%GetData(i), 3))
			} else {
				do ..AddTagElement(.Writer, NumQuery _ "c" _ i, RS.%GetData(i))
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
ClassMethod XMLToXSL(XML As %FileBinaryStream, ByRef Status As %Status) As %GlobalBinaryStream
{
	set Status = $$$OK
	set XSL = ##class(%Dictionary.CompiledXData).IDKEYOpen($ClassName(), "XSLData").Data
	set Status = ##class(%XML.XSLT.Transformer).TransformStream(XML, XSL, .ResultXSL)
	return ResultXSL
}

ClassMethod GetSQLArgsCount(ByRef Args) As %Integer
{
	set Iter = $Order(Args(""))
	set Count = 0
	while (Iter '= "")
	{	
		set Count = Count + 1
		set Iter = $Order(Args(Iter))
	}
 	return Count
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
			<xsl:when test="(.//*[local-name() = 'Data' and ((starts-with(substring-before(.,'c'),'q') = 'true') 
														and (translate(substring-after(substring-before(.,'c'),'q'),$vDigits,'') = '')) 
														and ((starts-with(substring-after(.,substring-before(.,'c')),'c') = 'true') 
														and (translate(substring-after(substring-after(.,substring-before(.,'c')),'c'),$vDigits,'') = '')) ])"> 
					<xsl:element name="xsl:for-each">
						<xsl:attribute name="select">
							<xsl:text>//query[@id = '</xsl:text><xsl:value-of select="substring-after(substring-before(.,'c'),'q')"/><xsl:text>']/row</xsl:text>
						</xsl:attribute>
						<xsl:copy>
							<xsl:apply-templates select="node()|@*"/>
						</xsl:copy>
					</xsl:element>
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
				<xsl:when test="((starts-with(substring-before(.,'c'),'q') = 'true') 
														and (translate(substring-after(substring-before(.,'c'),'q'),$vDigits,'') = '')) 
														and ((starts-with(substring-after(.,substring-before(.,'c')),'c') = 'true') 
														and (translate(substring-after(substring-after(.,substring-before(.,'c')),'c'),$vDigits,'') = ''))">
					<xsl:element name="xsl:value-of">
						<xsl:attribute name="select">
							<xsl:value-of select="."/>
						</xsl:attribute>
					</xsl:element>
				</xsl:when>
				<xsl:when test="((starts-with(substring-before(.,'a'),'q') = 'true') 
														and (translate(substring-after(substring-before(.,'a'),'q'),$vDigits,'') = '')) 
														and ((starts-with(substring-after(.,substring-before(.,'a')),'a') = 'true') 
														and (translate(substring-after(substring-after(.,substring-before(.,'a')),'a'),$vDigits,'') = ''))">
					<xsl:element name="xsl:value-of">
						<xsl:attribute name="select">
							<xsl:text>//query[@id = '</xsl:text><xsl:value-of select="substring-after(substring-before(.,'a'),'q')"/><xsl:text>']/args/</xsl:text><xsl:value-of select="."/>
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
