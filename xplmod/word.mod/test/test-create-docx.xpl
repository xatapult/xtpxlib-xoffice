<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:local="#local.c1j_xs3_bz"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" version="1.0" xpath-version="2.0" exclude-inline-prefixes="#all">

  <p:documentation>
    Test driver for the create-docx step
  </p:documentation>

  <!-- ================================================================== -->
  <!-- SETUP: -->

  <p:option name="template-docx-href" required="false" select="resolve-uri('test-xtp-template.docx', static-base-uri())"/>
  <p:option name="word-xml-href" required="false" select="resolve-uri('example-docx-result.xml', static-base-uri())"/>
  <p:option name="result-href" required="false" select="resolve-uri('../../../tmp/test-create-docx-result.docx', static-base-uri())"/>

  <p:output port="result" primary="true" sequence="false"/>
  <p:serialization port="result" method="xml" encoding="UTF-8" indent="true" omit-xml-declaration="false"/>

  <p:import href="../word.mod.xpl"/>

  <!-- ================================================================== -->

  <p:load dtd-validate="false">
    <p:with-option name="href" select="$word-xml-href"/> 
  </p:load>

  <xtlxo:create-docx>
    <p:with-option name="template-docx-href" select="$template-docx-href"/>
    <p:with-option name="result-docx-href" select="$result-href"/> 
  </xtlxo:create-docx>

</p:declare-step>
