<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:local="#local.vk3_rdn_cyb"
  version="3.0" xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" exclude-inline-prefixes="#all" name="xlsx-container-to-xml"
  type="xtlxo:xlsx-container-to-xml">

  <p:documentation>
    Local step that turns a (direct) zip container made from an .xlsx file into the extract format (see ../../xsd/xlsx-extract.xsd).
  </p:documentation>

  <!-- ======================================================================= -->
  <!-- DEVELOPMENT SETTINGS: -->

  <p:option name="develop" as="xs:boolean" static="true" select="false()"/>

  <!-- ======================================================================= -->
  <!-- PORTS: -->

  <p:input port="source" primary="true" sequence="false" content-types="xml">
    <p:documentation>The .xlsx file zip container</p:documentation>
  </p:input>

  <p:output port="result" primary="true" sequence="false" content-types="xml" serialization="map{'method': 'xml', 'indent': true()}">
    <p:documentation>The .xlsx file in XML format </p:documentation>
  </p:output>

  <!-- ================================================================== -->
  <!-- MAIN: -->

  <!-- Transform the contents into something manageable: -->
  <p:xslt>
    <p:with-input port="stylesheet" href="xsl-xlsx-container-to-xml/extract-xlsx-1.xsl"/>
  </p:xslt>

  <!-- Remove any empty rows and cells: -->
  <p:xslt>
    <p:with-input port="stylesheet" href="xsl-xlsx-container-to-xml/extract-xlsx-2.xsl"/>
  </p:xslt>

</p:declare-step>
