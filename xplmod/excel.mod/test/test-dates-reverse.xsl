<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="3.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" xmlns:local="#local.lrd_tgv_skb" xmlns="http://www.xtpxlib.nl/ns/xoffice"
  exclude-result-prefixes="#all" expand-text="true">
  <!-- ================================================================== -->
  <!--	
       Converts the Excel derived dates to xs:date
	-->
  <!-- ================================================================== -->

  <xsl:include href="../../../xslmod/excel-conversions.mod.xsl"/>

  <xsl:template match="/">
    <xsl:variable name="worksheet-name" as="xs:string" select="/*/@worksheet-name"/>

    <xlsx-modifications>
      <worksheet name="{$worksheet-name}">
        <xsl:for-each select="/*/result">
          <row index="{@row}">
            <column index="2">
              <!--<number>
                <xsl:value-of select="xtlxo:xs-date-to-excel-date(xs:date(.))"/>
              </number>-->
              <date>
                <xsl:value-of select="."/>
              </date>
            </column>
          </row>
        </xsl:for-each>
      </worksheet>
    </xlsx-modifications>

  </xsl:template>

</xsl:stylesheet>
