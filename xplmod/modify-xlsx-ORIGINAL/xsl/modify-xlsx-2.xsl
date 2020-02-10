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
    Sorts the worksheet data and removes doubles.	
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
  <xsl:template match="bcmconv-lib:containerFile/mso-wb:worksheet/mso-wb:sheetData[xtpxlib:string2boolean(@MODIFIED, false())]"
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <!-- -->
    <xsl:copy>
      <xsl:copy-of select="@* except @MODIFIED"/>
      <xsl:comment> == WORKSHEET MODIFIED - PHASE 2 == </xsl:comment>
      <!-- -->
      <!-- Take all rows with the same row number together: -->
      <xsl:for-each-group select="mso-wb:row" group-by="xs:integer(@r)">
        <xsl:sort select="xs:integer(@r)"/>
        <row>
          <!-- Remark: The next instruction copies *all* attributes of all current rows (this removes attribute doubles automatically). 
            The effect is that when we are merging any existing rows (some created by Excel), we automatically get all those non-understood, 
            Excel generated, maybe relevant, attributes in. -->
          <xsl:copy-of select="current-group()/@*"/>
          <!-- -->
          <!-- Within these rows, sort the cells and remove doubles: -->
          <xsl:for-each-group select="current-group()/mso-wb:c" group-by="@r">
            <xsl:sort select="@r"/>
            <xsl:copy-of select="current-group()[1]"/>
          </xsl:for-each-group>
          <!-- -->
        </row>
      </xsl:for-each-group>
    </xsl:copy>
    <!-- -->
  </xsl:template>
  <!-- -->
  <!-- ================================================================== -->
  <!-- SUPPORT: -->
  <!-- -->
  <xsl:template match="bcmconv-lib:warning-prevention-dummy-template"/>
  <!-- -->
</xsl:stylesheet>
