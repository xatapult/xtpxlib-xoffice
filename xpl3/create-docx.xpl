<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:local="#local.bmz_5hp_dyb"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" xmlns:xtlcon="http://www.xtpxlib.nl/ns/container" version="3.0" exclude-inline-prefixes="#all"
  name="create-docx" type="xtlxo:create-docx">

  <p:documentation>
      Takes as input the same kind of (unspecified) XML as create by `docx-to-xml.xpl` and tries to turn this into a Word file. 
      Unfinished and experimental (for instance: tables are not (yet) supported)! 
    </p:documentation>

  <!-- ======================================================================= -->
  <!-- IMPORTS: -->

  <p:import href="../xpl3mod/xoffice.mod.xpl"/>
  <p:import href="../../xtpxlib-container/xpl3mod/container-to-zip/container-to-zip.xpl"/>

  <!-- ======================================================================= -->
  <!-- DEVELOPMENT SETTINGS: -->

  <p:option name="develop" as="xs:boolean" static="true" select="false()"/>

  <!-- ======================================================================= -->
  <!-- PORTS: -->

  <p:input port="source" primary="true" sequence="false" content-types="xml">
    <p:documentation>The XML to convert into `.docx`.</p:documentation>
    <p:document href="../test/test-docx-xml.xml" use-when="$develop"/>
  </p:input>

  <p:output port="result" primary="true" sequence="false" content-types="xml">
    <p:documentation>The output is identical to the input but with `@timestamp`, `@docx-href-in` and `@docx-href-out` added to 
        the root element.</p:documentation>
  </p:output>

  <!-- ======================================================================= -->
  <!-- OPTIONS: -->

  <p:option name="docx-href-in" as="xs:string" required="true" use-when="not($develop)">
    <p:documentation>URI of the input (template) `.docx` file to process</p:documentation>
  </p:option>
  <p:option name="docx-href-in" as="xs:string" required="false" select="resolve-uri('../test/test-xtp-template.docx', static-base-uri())"
    use-when="$develop"/>

  <p:option name="docx-href-out" as="xs:string" required="true" use-when="not($develop)">
    <p:documentation>URI of the output `.docx` file.</p:documentation>
  </p:option>
  <p:option name="docx-href-out" as="xs:string" required="false" select="resolve-uri('tmp/create-docx-result.docx', static-base-uri())"
    use-when="$develop"/>

  <!-- ================================================================== -->
  <!-- MAIN: -->
  
  <!-- Create a container for the template input: -->
  <xtlxo:get-xo-container href="{$docx-href-in}" add-document-target-paths="true" name="docx-container"/>
  
  
  <!-- Add the Word XML to the container (as a normal container document): -->
  <p:insert match="/*" position="first-child">
    <p:with-input port="insertion">
        <xtlcon:document document-type="xtpxlib-word-xml">
          <WORDXMLHERE/>
        </xtlcon:document>
    </p:with-input>
  </p:insert>
  <p:viewport match="WORDXMLHERE">
    <p:identity>
      <p:with-input port="source" pipe="source@create-docx"/>
    </p:identity>
  </p:viewport>
  
  <!-- Merge the Word XML into the real .docx document XML: -->
  <p:xslt>
    <p:with-input port="stylesheet" href="xsl-modify-xlsx/create-docx-1.xsl"/>
  </p:xslt>

  <!-- Write the container back to disk: -->
  <xtlcon:container-to-zip href-target-zip="{$docx-href-out}"/>

  <!-- Create the output xml: -->
  <p:add-attribute attribute-name="docx-href-in" match="/*">
    <p:with-input port="source" pipe="source@create-docx"/>
    <p:with-option name="attribute-value" select="$docx-href-in"/>
  </p:add-attribute>
  <p:add-attribute attribute-name="docx-href-out" match="/*">
    <p:with-option name="attribute-value" select="$docx-href-out"/>
  </p:add-attribute>
  <p:add-attribute attribute-name="timestamp" match="/*">
    <p:with-option name="attribute-value" select="string(current-dateTime())"/>
  </p:add-attribute>

</p:declare-step>
