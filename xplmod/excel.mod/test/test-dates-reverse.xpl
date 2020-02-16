<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:local="#local.c1j_xs3_bz"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" version="1.0" xpath-version="2.0" exclude-inline-prefixes="#all">

  <p:documentation>
    Test driver for testing the xs:date to Excel date conversion.
    
    Takes the values date in the first row of the input Excel sheet and puts them in the second row, attempted as a date. 
  </p:documentation>

  <!-- ================================================================== -->
  <!-- SETUP: -->

  <p:option name="xlsx-href" required="false" select="resolve-uri('test-dates.xlsx', static-base-uri())"/>
  <p:option name="xlsx-href-out" required="false" select="resolve-uri('../../../tmp/test-dates-reverse-result.xlsx', static-base-uri())"/>

  <p:output port="result" primary="true" sequence="false"/>
  <p:serialization port="result" method="xml" encoding="UTF-8" indent="true" omit-xml-declaration="false"/>

  <p:import href="../excel.mod.xpl"/>

  <!-- ================================================================== -->

  <xtlxo:extract-xlsx>
    <p:with-option name="xlsx-href" select="$xlsx-href"/>
  </xtlxo:extract-xlsx>
  
  <p:xslt>
    <p:input port="stylesheet">
      <p:document href="test-dates.xsl"/>
    </p:input>
    <p:with-param name="null" select="()"/>
  </p:xslt>
  
  <p:xslt>
    <p:input port="stylesheet">
      <p:document href="test-dates-reverse.xsl"/>
    </p:input>
    <p:with-param name="null" select="()"/>
  </p:xslt>
  
  <xtlxo:modify-xlsx>
    <p:with-option name="xlsx-href-in" select="$xlsx-href"/> 
    <p:with-option name="xlsx-href-out" select="$xlsx-href-out"/> 
  </xtlxo:modify-xlsx>
  
</p:declare-step>
