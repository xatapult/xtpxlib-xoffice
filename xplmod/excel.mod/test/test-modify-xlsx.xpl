<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice"
  version="1.0" xpath-version="2.0" exclude-inline-prefixes="#all">

  <p:documentation>
    Test driver for the xtlxo:modify-xlsx step
  </p:documentation>

  <!-- ================================================================== -->
  <!-- SETUP: -->

  <p:option name="specification-href" required="false" select="resolve-uri('test-modify-xlsx-input.xml', static-base-uri())"/>
  <p:option name="xlsx-href-in" required="false" select="resolve-uri('test-modify.xlsx', static-base-uri())"/>
  <p:option name="xlsx-href-out" required="false" select="resolve-uri('../../../tmp/modify-xlsx-result.xlsx', static-base-uri())"/>

  <p:output port="result"/>
  <p:serialization port="result" method="xml" encoding="UTF-8" indent="true"/>

  <p:import href="../excel.mod.xpl"/>

  <!-- ================================================================== -->

  <p:load dtd-validate="false">
    <p:with-option name="href" select="$specification-href"/>
  </p:load>

  <xtlxo:modify-xlsx>
    <p:with-option name="xlsx-href-in" select="$xlsx-href-in"/>
    <p:with-option name="xlsx-href-out" select="$xlsx-href-out"/>
  </xtlxo:modify-xlsx>

</p:declare-step>
