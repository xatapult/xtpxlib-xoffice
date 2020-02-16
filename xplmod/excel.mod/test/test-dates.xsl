<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="3.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" xmlns:local="#local.lrd_tgv_skb" exclude-result-prefixes="#all" expand-text="true">
  <!-- ================================================================== -->
  <!--	
       Converts the Excel derived dates to xs:date
	-->
  <!-- ================================================================== -->

  <xsl:include href="../../../xslmod/excel-conversions.mod.xsl"/>

  <xsl:template match="/">
    <results worksheet-name="{/*/xtlxo:worksheet[1]/@name}">
      <xsl:for-each select="/*/xtlxo:worksheet[1]//xtlxo:cell/xtlxo:value[string(.) castable as xs:integer]">
        <result value="{.}" row="{../../@index}">
          <xsl:value-of select="xtlxo:excel-date-to-xs-date(xs:integer(.))"/>
        </result>
      </xsl:for-each>
    </results>
  </xsl:template>

</xsl:stylesheet>
