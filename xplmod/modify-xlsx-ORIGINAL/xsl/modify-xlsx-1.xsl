<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions" xmlns:xtpxplib="http://www.xatapult.nl/namespaces/common/xproc/library"
  xmlns:xtpxmsolib="http://www.xatapult.nl/namespaces/common/xslt/mso/library" xmlns:xtpxlib="http://www.xatapult.nl/namespaces/common/xslt/library"
  xmlns:mso-wb="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:mso-rels="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:bcmconv-lib="http://www.malmberg.nl/namespaces/bcmconverters/library"
  xmlns:local="http://www.xatapult.nl/namespaces/common/xslt/mso/library/local" exclude-result-prefixes="#all">
  <!-- ================================================================== -->
  <!--*	
    Performs a "raw" insertion of the modification data into the worksheets. After this, the worksheet data must be sorted and made unique!	
	-->
  <!-- ================================================================== -->
  <!-- SETUP: -->
  <!-- -->
  <xsl:output method="xml" indent="no" encoding="UTF-8"/>
  <!-- -->
  <xsl:include href="../../../libxsl/office-lib.xsl"/>
  <!-- -->
  <xsl:variable name="container" as="element(bcmconv-lib:conversionContainer)" select="/*"/>
  <!-- -->
  <!-- Find the modification specification: -->
  <xsl:variable name="modification-spec" as="element(xtpxmsolib:ExcelModifications)"
    select="/*/bcmconv-lib:containerFile[@type eq 'xlsx-modification']/*"/>
  <xsl:variable name="global-force-string" as="xs:boolean" select="xtpxlib:string2boolean($modification-spec/@force-string, false())"/>
  <!-- -->
  <!-- Get the main workbook: -->
  <xsl:variable name="main-workbook" as="element(mso-wb:workbook)"
    select="bcmconv-lib:get-file-root-from-relationship-type(/*, '', $bcmconv-lib:office-relationship-type-main-document, true())"/>
  <!-- -->
  <!-- Some data for the cell index computation: -->
  <xsl:variable name="codepoint-A" as="xs:integer" select="string-to-codepoints('A')"/>
  <xsl:variable name="char-index-max" as="xs:integer" select="26"/>
  <!-- -->
  <!-- ================================================================== -->
  <!-- -->
  <xsl:template match="@* | node()">
    <xsl:copy>
      <xsl:apply-templates select="@* | node()"/>
    </xsl:copy>
  </xsl:template>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:template match="bcmconv-lib:containerFile/mso-wb:worksheet/mso-wb:sheetData" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <!-- -->
    <!-- Get the name of this worksheet and find any modifications for this worksheet: -->
    <xsl:variable name="worksheet-name" as="xs:string" select="lower-case(local:worksheet-name(..))"/>
    <xsl:variable name="modifications" as="element(xtpxmsolib:Worksheet)?"
      select="$modification-spec/xtpxmsolib:Worksheet[lower-case(@name) eq $worksheet-name]"/>
    <!-- -->
    <xsl:choose>
      <!-- -->
      <!-- When there are any modifications, merge them in: -->
      <xsl:when test="exists($modifications/xtpxmsolib:Row/xtpxmsolib:Cell)">
        <xsl:copy>
          <xsl:copy-of select="@*"/>
          <!-- Remark: We mark this sheet data section as modified down here so we can identify modified section in the next pass. 
            Make sure to remove the MODIFIED attribute there! -->
          <xsl:attribute name="MODIFIED" select="true()"/>
          <xsl:comment> == WORKSHEET MODIFIED - PHASE 1 == </xsl:comment>
          <!-- -->
          <!-- Remark: Make sure the new data comes first... This ensures that in sorting (and removing doubles) the new data gets precedence. -->
          <xsl:for-each select="$modifications/xtpxmsolib:Row">
            <xsl:variable name="row" as="xs:integer" select="xs:integer(@index)"/>
            <row r="{@index}">
              <!-- -->
              <xsl:for-each select="xtpxmsolib:Cell">
                <xsl:variable name="col" as="xs:integer" select="xs:integer(@index)"/>
                <xsl:variable name="value" as="xs:string" select="xtpxmsolib:Value"/>
                <xsl:variable name="force-string" as="xs:boolean?"
                  select="if (exists(xtpxmsolib:Value/@force-string)) then xtpxlib:string2boolean(xtpxmsolib:Value/@force-string, false()) else ()"/>
                <c r="{local:cell-index($row, $col)}">
                  <xsl:choose>
                    <xsl:when
                      test="(exists($force-string) and $force-string) or (empty($force-string) and $global-force-string) or not($value castable as xs:double)">
                      <xsl:attribute name="t" select="'inlineStr'"/>
                      <is>
                        <t>
                          <xsl:value-of select="$value"/>
                        </t>
                      </is>
                    </xsl:when>
                    <xsl:otherwise>
                      <v>
                        <xsl:value-of select="$value"/>
                      </v>
                    </xsl:otherwise>
                  </xsl:choose>
                </c>
              </xsl:for-each>
              <!-- -->
            </row>
          </xsl:for-each>
          <!-- -->
          <!-- Existing row/cell information: -->
          <xsl:apply-templates/>
          <!-- -->
        </xsl:copy>
      </xsl:when>
      <!-- -->
      <!-- No modifications for this worksheet: -->
      <xsl:otherwise>
        <xsl:copy-of select="."/>
      </xsl:otherwise>
      <!-- -->
    </xsl:choose>
    <!-- -->
  </xsl:template>
  <!-- -->
  <!-- ================================================================== -->
  <!-- SUPPORT: -->
  <!-- -->
  <xsl:function name="local:worksheet-name" as="xs:string">
    <xsl:param name="worksheet" as="element(mso-wb:worksheet)"/>
    <!-- -->
    <xsl:sequence select="$main-workbook/mso-wb:sheets/mso-wb:sheet[local:workbook-worksheet-ref-to-worksheet(.) is $worksheet]/@name"/>
  </xsl:function>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:function name="local:workbook-worksheet-ref-to-worksheet" as="element(mso-wb:worksheet)">
    <xsl:param name="workbook-worksheet-ref" as="element(mso-wb:sheet)"/>
    <!-- -->
    <xsl:sequence
      select="bcmconv-lib:get-file-root-from-relationship-id($container, bcmconv-lib:get-href($main-workbook), $workbook-worksheet-ref/@mso-rels:id, true())"
    />
  </xsl:function>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:function name="local:cell-index" as="xs:string">
    <xsl:param name="row" as="xs:integer"/>
    <xsl:param name="col" as="xs:integer"/>
    <!-- -->
    <xsl:sequence select="concat(local:number2alpha-index($col), string($row))"/>
  </xsl:function>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:function name="local:number2alpha-index" as="xs:string">
    <xsl:param name="number" as="xs:integer"/>
    <!-- -->
    <xsl:choose>
      <xsl:when test="$number le $char-index-max">
        <xsl:sequence select="local:number2alpha($number)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="current" as="xs:integer" select="$number mod $char-index-max"/>
        <xsl:variable name="remainder" as="xs:integer" select="xs:integer(floor($number div $char-index-max))"/>
        <xsl:sequence select="concat(local:number2alpha-index($remainder), local:number2alpha($current))"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:function name="local:number2alpha" as="xs:string">
    <xsl:param name="number" as="xs:integer"/>
    <!-- $number is supposed to be in-between 1 and $char-index-max! -->
    <!-- -->
    <xsl:sequence select="codepoints-to-string($number - 1 + $codepoint-A)"/>
  </xsl:function>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:template match="bcmconv-lib:warning-prevention-dummy-template"/>
  <!-- -->
</xsl:stylesheet>
