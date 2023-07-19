<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:local="#local.fk4_fbn_cybx"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" version="3.0" exclude-inline-prefixes="#all" name="docx-to-xml" type="xtlxo:docx-to-xml">

  <p:documentation>
    Extracts the contents of a Word (`.docx`) file in a more useable XML format (unspecified). Somewhat experimental and unfinished!
  </p:documentation>

  <!-- ======================================================================= -->
  <!-- IMPORTS: -->

  <p:import href="../xpl3mod/xoffice.mod.xpl"/>

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
    <p:documentation>Document reference of the `.docx` file to process (must have `file://` in front).</p:documentation>
  </p:option>
  <p:option name="xlsx-href" as="xs:string" required="false" use-when="$develop" select="resolve-uri('../test/test.docx', static-base-uri())"/>

  <!-- ================================================================== -->
  <!-- MAIN: -->

  <!-- Get an xo container and convert it: -->
  <xtlxo:get-xo-container href="{$xlsx-href}"/>
  <xtlxo:docx-xo-container-to-xml/>

</p:declare-step>
