<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:xtpxplib="http://www.xatapult.nl/namespaces/common/xproc/library" version="1.0" xpath-version="2.0" exclude-inline-prefixes="#all">
  <!-- -->
  <p:documentation>
    Testbed for the xtpxplib:modify-xlsx step
  </p:documentation>
  <!-- -->
  <!-- ================================================================== -->
  <!-- SETUP: -->
  <!-- -->
  <p:option name="debug" required="false" select="string(false())"/>
  <!-- -->
  <p:option name="href" required="false" select="resolve-uri('test.xlsx', static-base-uri())"/>
  <p:option name="output-href" required="false" select="'file:/C:/xdata2/score-xml-TMPOUT/modified-excel.xlsx'"/>
  <p:option name="modification-spec-file" required="false" select="resolve-uri('test-modify-xlsx-input.xml', static-base-uri())"/>
  <!-- -->
  <p:output port="result"/>
  <p:serialization port="result" method="xml" encoding="UTF-8" indent="true"/>
  <!-- -->
  <p:import href="../xtp-office-lib.xpl"/>
  <!-- -->
  <!-- ================================================================== -->
  <!-- -->
  <p:load dtd-validate="false">
    <p:with-option name="href" select="$modification-spec-file"/>
  </p:load>
  <!-- -->
  <xtpxplib:modify-xlsx>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="input-xlsx-href" select="$href"/>
    <p:with-option name="output-xlsx-href" select="$output-href"/>
  </xtpxplib:modify-xlsx>
  <!-- -->
  
  
</p:declare-step>
