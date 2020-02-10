<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions" xmlns:xtpxplib="http://www.xatapult.nl/namespaces/common/xproc/library"
  xmlns:xtpxmsolib="http://www.xatapult.nl/namespaces/common/xslt/mso/library" xmlns="http://www.xatapult.nl/namespaces/common/xslt/mso/library"
  xmlns:xtpxlib="http://www.xatapult.nl/namespaces/common/xslt/library" xmlns:mso-wb="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:mso-rels="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:bcmconv-lib="http://www.malmberg.nl/namespaces/bcmconverters/library"
  xmlns:local="http://www.xatapult.nl/namespaces/common/xslt/mso/library/local" exclude-result-prefixes="#all">
  <!-- ================================================================== -->
  <!--*	
    /TBD: Description/		
	-->
  <!-- ================================================================== -->
  <!-- SETUP: -->
  <!-- -->
  <xsl:output method="xml" indent="no" encoding="UTF-8"/>
  <!-- -->
  <xsl:include href="../../../libxsl/office-lib.xsl"/>
  <!-- -->
  <!-- ================================================================== -->
  <!-- -->
  <xsl:template match="@* | node()">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@* | node()"/>
    </xsl:copy>
  </xsl:template>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:template match="/*">
    <Workbook href="{/*/@source-zip-ref}" timestamp="{current-dateTime()}">
      <!-- -->
      <!-- Find the workbook first: -->
      <xsl:variable name="workbook" as="element(mso-wb:workbook)"
        select="bcmconv-lib:get-file-root-from-relationship-type(/*, '', $bcmconv-lib:office-relationship-type-main-document, true())"/>
      <xsl:variable name="workbook-href" as="xs:string" select="bcmconv-lib:get-href($workbook)"/>
      <!-- -->
      <!-- Get the (optional) shared string table: -->
      <xsl:variable name="shared-strings" as="element(mso-wb:sst)?"
        select="bcmconv-lib:get-file-root-from-relationship-type(/*, $workbook-href, $bcmconv-lib:office-relationship-type-shared-strings, false())"/>
      <xsl:variable name="styles" as="element(mso-wb:styleSheet)?"
        select="bcmconv-lib:get-file-root-from-relationship-type(/*, $workbook-href, $bcmconv-lib:office-relationship-type-styles, false())"/>
      <!-- -->
      <!-- Process the worksheets: -->
      <xsl:for-each select="$workbook/mso-wb:sheets/mso-wb:sheet">
        <xsl:variable name="sheet" as="element(mso-wb:worksheet)"
          select="bcmconv-lib:get-file-root-from-relationship-id(/*, $workbook-href, @mso-rels:id, true())"/>
        <xsl:call-template name="process-worksheet">
          <xsl:with-param name="worksheet" select="$sheet"/>
          <xsl:with-param name="name" select="string(@name)"/>
          <xsl:with-param name="shared-strings" select="$shared-strings"/>
          <xsl:with-param name="styles" select="$styles"/>
        </xsl:call-template>
      </xsl:for-each>
      <!-- -->
    </Workbook>
  </xsl:template>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:template name="process-worksheet">
    <xsl:param name="worksheet" as="element(mso-wb:worksheet)" required="yes"/>
    <xsl:param name="shared-strings" as="element(mso-wb:sst)?" required="yes"/>
    <xsl:param name="styles" as="element(mso-wb:styleSheet)?" required="yes"/>
    <xsl:param name="name" as="xs:string" required="yes"/>
    <!-- -->
    <Worksheet name="{$name}">
      <xsl:for-each select="$worksheet/mso-wb:sheetData/mso-wb:row">
        <Row index="{@r}">
          <xsl:for-each select="mso-wb:c">
            <Cell index="{local:excelref-to-index(@r)}">
              <!-- -->
              <!-- Get style information for the cell: -->
              <xsl:variable name="main-style-classes" as="xs:string*" select="local:get-style-classes($styles, @s)"/>
              <xsl:if test="exists($main-style-classes)">
                <xsl:attribute name="class" select="string-join($main-style-classes, ' ')"/>
              </xsl:if>
              <!-- -->
              <!-- Get the cell contents: -->
              <xsl:variable name="celltype" as="xs:string" select="string(@t)"/>
              <xsl:choose>
                <!-- -->
                <!-- Shared string (with optional markup): -->
                <xsl:when test="exists(mso-wb:v) and ($celltype eq 's')">
                  <xsl:variable name="shared-string-index" as="xs:integer" select="xs:integer(mso-wb:v) + 1"/>
                  <xsl:variable name="shared-string-elm" as="element(mso-wb:si)?" select="$shared-strings/mso-wb:si[$shared-string-index]"/>
                  <Value>
                    <xsl:choose>
                      <xsl:when test="exists($shared-string-elm/mso-wb:r)">
                        <xsl:for-each select="$shared-string-elm/mso-wb:r">
                          <xsl:variable name="style-info" as="xs:string*" select="local:get-relevant-font-info(mso-wb:rPr/*)"/>
                          <xsl:choose>
                            <xsl:when test="exists($style-info)">
                              <span class="{string-join($style-info, ' ')}">
                                <xsl:value-of select="mso-wb:t"/>
                              </span>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="mso-wb:t"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:for-each>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="string-join($shared-string-elm//text()/string(), '')"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </Value>
                  <!-- TBD: Process markup? -->
                </xsl:when>
                <!-- -->
                <!-- Direct vaalue: -->
                <xsl:otherwise>
                  <Value>
                    <xsl:value-of select="mso-wb:v"/>
                  </Value>
                </xsl:otherwise>
                <!-- -->
              </xsl:choose>
              <!-- -->
              <!-- See if there is a formula, if so add it: -->
              <xsl:if test="exists(mso-wb:f)">
                <Formula>
                  <xsl:value-of select="mso-wb:f"/>
                </Formula>
              </xsl:if>
              <!-- -->
            </Cell>
          </xsl:for-each>
        </Row>
      </xsl:for-each>
    </Worksheet>
  </xsl:template>
  <!-- -->
  <!-- ================================================================== -->
  <!-- STYLE HANDLING -->
  <!-- Remark: We only take a minimum of style information from the Excel sheet, just b, i, u. -->
  <!-- -->
  <xsl:function name="local:get-style-classes" as="xs:string*">
    <xsl:param name="styles" as="element(mso-wb:styleSheet)?"/>
    <xsl:param name="style-index" as="xs:string?" required="yes"/>
    <!-- -->
    <!-- First, get the main cellXfs/xf record: -->
    <xsl:variable name="cell-xf" as="element(mso-wb:xf)?" select="local:get-list-entry($styles/mso-wb:cellXfs/mso-wb:xf, $style-index)"/>
    <!-- -->
    <!-- Get from this the font information from both the direct ref and from the cellStyleXfs/xf record: -->
    <xsl:variable name="font-entries" as="element(mso-wb:font)*" select="$styles/mso-wb:fonts/mso-wb:font"/>
    <!-- -->
    <xsl:variable name="direct-font-entry" as="element(mso-wb:font)?" select="local:get-list-entry($font-entries, $cell-xf/@fontId)"/>
    <!-- -->
    <xsl:variable name="cell-style-xf" as="element(mso-wb:xf)?" select="local:get-list-entry($styles/mso-wb:cellStyleXfs/mso-wb:xf, $cell-xf/@xfId)"/>
    <xsl:variable name="style-font-entry" as="element(mso-wb:font)?" select="local:get-list-entry($font-entries, $cell-style-xf/@fontId)"/>
    <!-- -->
    <!-- make this into a coherent lost of classes: -->
    <xsl:sequence select="distinct-values((local:get-relevant-font-info($direct-font-entry/*), local:get-relevant-font-info($style-font-entry/*)))"/>
  </xsl:function>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:function name="local:get-list-entry" as="element()?">
    <xsl:param name="list" as="element()*"/>
    <xsl:param name="index" as="xs:string?"/>
    <!-- -->
    <xsl:choose>
      <xsl:when test="exists($list) and exists($index) and ($index castable as xs:integer)">
        <xsl:variable name="list-index" as="xs:integer" select="xs:integer($index) + 1"/>
        <xsl:sequence select="$list[$list-index]"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="()"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:function name="local:get-relevant-font-info" as="xs:string*">
    <xsl:param name="font-style-elements" as="element()*"/>
    <!-- -->
    <xsl:variable name="relevant-element-names" as="xs:string+" select="('b', 'u', 'i')"/>
    <xsl:sequence select="distinct-values($font-style-elements[local-name(.) = $relevant-element-names]/local-name(.))"/>
  </xsl:function>
  <!-- -->
  <!-- ================================================================== -->
  <!-- HELPERS -->
  <!-- -->
  <xsl:function name="local:excelref-to-index" as="xs:integer">
    <xsl:param name="excelref" as="xs:string"/>
    <!-- -->
    <!-- Use the alphabetic part only: -->
    <xsl:variable name="excelref-alpha" as="xs:string" select="upper-case(replace($excelref, '^([A-Za-z]+).+', '$1'))"/>
    <!-- -->
    <!-- Compute the index: -->
    <xsl:variable name="charindex-A" as="xs:integer" select="string-to-codepoints('A')"/>
    <xsl:variable name="excelref-charindexes" as="xs:integer+" select="string-to-codepoints($excelref-alpha)"/>
    <xsl:variable name="excelref-charindexes-normalized" as="xs:integer+" select="for $ci in $excelref-charindexes return ($ci - $charindex-A + 1)"/>
    <xsl:sequence select="local:excelref-to-index-helper(0, $excelref-charindexes-normalized)"/>
  </xsl:function>
  <!-- -->
  <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
  <!-- -->
  <xsl:function name="local:excelref-to-index-helper" as="xs:integer">
    <xsl:param name="prev" as="xs:integer"/>
    <xsl:param name="charindexes-normalized" as="xs:integer+"/>
    <!-- -->
    <xsl:choose>
      <xsl:when test="empty($charindexes-normalized)">
        <xsl:sequence select="$prev"/>
      </xsl:when>
      <xsl:when test="count($charindexes-normalized) eq 1">
        <xsl:sequence select="($prev * 26) + $charindexes-normalized[1]"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="index-value" as="xs:integer" select="($prev * 26) + $charindexes-normalized[1]"/>
        <xsl:sequence select="local:excelref-to-index-helper($index-value, subsequence($charindexes-normalized, 2))"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  <!-- -->
</xsl:stylesheet>
