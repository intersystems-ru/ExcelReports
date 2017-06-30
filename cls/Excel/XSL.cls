Class Excel.XSL Extends %CSP.Page
{

/// Входной файл
Parameter XMLFILE = "C:\temp\input.xml";

/// Выходной файл
Parameter XLSFILE = "C:\temp\output.xls";

/// XSL трансформация в файл
/// на вход: запрос, имя входного файла, имя выходного файла
ClassMethod XSLtoFile(Query As %String = "", XMLFile As %String = {..#XMLFILE}, XLSFile As %String = {..#XLSFILE}) As %Status
{
	set sc = $$$OK

	set XML=##class(%FileCharacterStream).%New()
	set XML.Filename = XMLFile
	
	set XSL = ..XMLInXSL(XML,.sc) if $$$ISERR(sc) quit sc

	set XMLData = ##class(%GlobalCharacterStream).%New()
	
	set sc=..genXMLData(XMLData, Query) if $$$ISERR(sc) quit sc

	set OutputFile = ##class(%FileCharacterStream).%New()
	set OutputFile.TranslateTable = "UTF8"
	set OutputFile.Filename = XLSFile
	set sc = ##class(%XML.XSLT.Transformer).TransformStream(XMLData, XSL, .OutputFile)
	if $$$ISERR(sc) quit sc  
	
	quit OutputFile.%Save()
}

/// генерация XML с данными
ClassMethod genXMLData(stream As %Stream.Object, Query As %String) As %Status
{
	set tSC=$$$OK
	if '$IsObject($g(stream)) s stream=##class(%GlobalCharacterStream).%New()
	
	set fWriter=##class(%XML.Writer).%New()
	set fWriter.Indent=1
	set tSC=fWriter.OutputToStream(stream)

	do fWriter.RootElement("SQLResult")
	
	do ..addSQLTags(.fWriter, Query)

	if $$$ISERR(tSC) q tSC	
 	do fWriter.EndRootElement()
	q tSC
}

ClassMethod addSQLTags(writer, Query As %String)
{
	set ^Rows = 0
	Set rs=##class(%ResultSet).%New("%DynamicQuery:SQL")
	Set sc=rs.Prepare(Query)
	do rs.Execute()

	set colCount = rs.%ResultColumnCountGet()
	while rs.Next(.sc)
	{ 
	set ^Rows = ^Rows + 1
		If $$$ISERR(sc) return 0
		d writer.Element("row")
		for i=1:1:colCount 
		{ if (rs.GetColumnType(i) = 2)
			{
				d ..addTagElement(.writer,"c"_i,$zd(rs.Data(rs.GetColumnName(i)),3))
			}else{
				d ..addTagElement(.writer,"c"_i,rs.Data(rs.GetColumnName(i)))
			}
		}
		d writer.EndElement()	
	}
}

ClassMethod addTagElement(writer As %XML.Writer, tagName As %String, tagValue As %String) As %Status
{
	s tSC=$$$OK
	d writer.Element(tagName)  
	d writer.Write(tagValue)  
	d writer.EndElement()
	q tSC
}

ClassMethod XMLInXSL(XML As %Stream.FileCharacter, ByRef sc As %Status) As %FileCharacterStream
{
	set sc = $$$OK
	
	Set tStream = ##class(%Dictionary.CompiledXData).%OpenId(..%ClassName(1)_"||XSLData").Data
   	set sc = ##class(%XML.TextReader).ParseStream(tStream,.reader)
   	
	set XSL = ##class(%FileCharacterStream).%New()
	set XSL.Filename = $EXTRACT(XML.Filename,1,*-3)_"xsl"
	
	set sc = reader.MoveToContent()
	set sc = XSL.WriteLine(reader.Value)
	
	set gbs = ##class(%GlobalBinaryStream).%New()
	set sc = ##class(%XML.XSLT.Transformer).TransformStream(XML, XSL, .gbs)
	
	return gbs
}

ClassMethod test() As %Status
{
	set sc = $$$OK

	set XML=##class(%FileCharacterStream).%New()
	set XML.Filename = ..#XMLFILE
	
	w $zt($p($h,",",2))," Конвертирование XML в XSL: ...",!
	set XSL = ..XMLInXSL(XML,.sc) if $$$ISERR(sc) quit sc
	w $zt($p($h,",",2))," Конвертирование XML в XSL: OK!",!
	w XSL.%Save()
   	q sc
}

XData XSLData
{
<XSL>
<![CDATA[
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


</xsl:stylesheet>]]>
</XSL>
}

}

