<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:local="#local.fk4_fbn_cyb"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" xmlns:xtlcon="http://www.xtpxlib.nl/ns/container" xmlns:xtlc="http://www.xtpxlib.nl/ns/common"
  version="3.0" exclude-inline-prefixes="#all" name="xlsx-to-xml" type="xtlxo:xlsx-to-xml">

  <p:documentation>
    Extracts the contents of an Excel (`.xlsx`) file in a more useable [XML format](%xlsx-extract.xsd).
  </p:documentation>

  <!-- ======================================================================= -->
  <!-- IMPORTS: -->

  <p:import href="../../xtpxlib-container/xpl3mod/zip-to-container/zip-to-container.xpl"/>
  <p:import href="shared/xlsx-container-to-xml.xpl"></p:import>

  <!-- ======================================================================= -->
  <!-- DEVELOPMENT SETTINGS: -->

  <p:option name="develop" as="xs:boolean" static="true" select="false()"/>

  <!-- ======================================================================= -->
  <!-- PORTS: -->
 
  <p:output port="result" primary="true" sequence="false" content-types="xml" serialization="map{'method': 'xml', 'indent': true()}">
    <p:documentation>The resulting XML document.</p:documentation>
  </p:output>

  <!-- ======================================================================= -->
  <!-- OPTIONS: -->

  <p:option name="xlsx-href" as="xs:string" required="true" use-when="not($develop)">
    <p:documentation>Document reference of the `.xlsx` file to process (must have `file://` in front).</p:documentation>
  </p:option>
  <p:option name="xlsx-href" as="xs:string" required="false" use-when="$develop" select="resolve-uri('../test/test.xlsx', static-base-uri())"/>
  
  <!-- ================================================================== -->
  <!-- MAIN: -->

  <!-- Extract the Excel document into a container (also treat these weird .rels files as XML): -->
  <xtlcon:zip-to-container href-source-zip="{$xlsx-href}">
    <p:with-option name="override-content-types" select="[ ['\.rels$', 'text/xml'] ]"></p:with-option>
  </xtlcon:zip-to-container>
    
  <xtlxo:xlsx-container-to-xml/>

</p:declare-step>
