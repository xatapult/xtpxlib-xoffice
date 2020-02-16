<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="3.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema"
   xmlns:xtpxplib="http://www.xatapult.nl/namespaces/common/xproc/library"
  xmlns:mso-wb="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:mso-rels="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xtlcon="http://www.xtpxlib.nl/ns/container"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" xmlns:local="#local-998hy5" exclude-result-prefixes="#all">
  <!-- ================================================================== -->
  <!--*	
    Sorts the worksheet data and removes doubles.	
	-->
  <!-- ================================================================== -->
  <!-- SETUP: -->

  <xsl:output method="xml" indent="no" encoding="UTF-8"/>

  <xsl:mode on-no-match="shallow-copy"/>
  
  <!-- ================================================================== -->

  <xsl:template match="xtlcon:document/mso-wb:worksheet/mso-wb:sheetData[xs:boolean(@xtlxo:MODIFIED)]"
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <!-- We only need to handle modified sections (identified with @xtlxo:MODIFIED="true"). We take this attribute off again. -->

    <xsl:copy>
      <xsl:copy-of select="@* except @xtlxo:MODIFIED"/>

      <!-- Take all rows with the same row number together: -->
      <xsl:for-each-group select="mso-wb:row" group-by="xs:integer(@r)">
        <xsl:sort select="xs:integer(@r)"/>
        <row>
          <!-- Remark: The next instruction copies *all* attributes of all current rows (this removes attribute doubles automatically). 
            The effect is that when we are merging any existing rows (some created by Excel), we automatically get all those non-understood, 
            Excel generated, maybe relevant, attributes in. -->
          <xsl:copy-of select="current-group()/@*"/>

          <!-- Within these rows, sort the cells and remove doubles: -->
          <xsl:for-each-group select="current-group()/mso-wb:c" group-by="@r">
            <xsl:sort select="@r"/>
            <c>
              <!-- Copy again all attributes of the cells for this coordinate but retain the @t value (type of the cell) of 
                the one we're actually going to use: -->
              <xsl:copy-of select="current-group()/@* except current-group()/@t"/>
              <xsl:copy-of select="current-group()[1]/@t"/>
              <xsl:copy-of select="current-group()[1]/*"/>
            </c>
          </xsl:for-each-group>

        </row>
      </xsl:for-each-group>
    </xsl:copy>

  </xsl:template>

  <!-- ================================================================== -->
  <!-- SUPPORT: -->

  <xsl:template match="xtlcon:warning-prevention-dummy-template | xtlxo:warning-prevention-dummy-template"/>

</xsl:stylesheet>
