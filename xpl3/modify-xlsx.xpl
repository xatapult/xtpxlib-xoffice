<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:array="http://www.w3.org/2005/xpath-functions/array" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:local="#local.bmz_5hp_dyb"
  xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice" xmlns:xtlcon="http://www.xtpxlib.nl/ns/container" version="3.0" exclude-inline-prefixes="#all"
  name="modify-xlsx" type="xtlxo:modify-xlsx">

  <p:documentation>
      Takes an input/template Excel (`.xlsx`)  and a [modification specification](%xlsx-modify.xsd) and from this creates a 
      new modified Excel file that merges these two sources.
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
    <p:documentation>The [modification specification](%xlsx-modify.xsd).</p:documentation>
    <p:document href="../test/test-modify-xlsx-input.xml" use-when="$develop"/>
  </p:input>

  <p:output port="result" primary="true" sequence="false" content-types="xml">
    <p:documentation>The output is identical to the input but with `@timestamp`, `@xlsx-href-in` and `@xlsx-href-out` added to 
        the root element.</p:documentation>
  </p:output>

  <!-- ======================================================================= -->
  <!-- OPTIONS: -->

  <p:option name="xlsx-href-in" as="xs:string" required="true" use-when="not($develop)">
    <p:documentation>URI of the input (template) `.xlsx` file to process</p:documentation>
  </p:option>
  <p:option name="xlsx-href-in" as="xs:string" required="false" select="resolve-uri('../test/test-modify.xlsx', static-base-uri())"
    use-when="$develop"/>

  <p:option name="xlsx-href-out" as="xs:string" required="true" use-when="not($develop)">
    <p:documentation>URI of the output `.xlsx` file.</p:documentation>
  </p:option>
  <p:option name="xlsx-href-out" as="xs:string" required="false" select="resolve-uri('tmp/modify-xlsx-result.xlsx', static-base-uri())"
    use-when="$develop"/>

  <!-- ================================================================== -->
  <!-- MAIN: -->

  <p:identity name="modify-xlsx-original-input"/>

  <!-- Wrap the modification document in in an appropriate <xtlcon:document> element so we can insert it in the container: -->
  <p:wrap match="/*" wrapper="xtlcon:document"/>
  <p:add-attribute match="/*" attribute-name="type" attribute-value="modification-specification"/>
  <p:identity name="modification-specification-xtlcon-document"/>

  <!-- Create a container for the Excel input: -->
  <xtlxo:get-xo-container href="{$xlsx-href-in}" add-document-target-paths="true"/>
  <p:identity name="xlsx-container"/>

  <!-- Get the .xlsx information in extract format so we can more easily lookup names. Out in an appropriate <xtlcon:document> element 
      so we can insert it in the container: -->
  <xtlxo:xlsx-xo-container-to-xml/>
  <p:wrap match="/*" wrapper="xtlcon:document"/>
  <p:add-attribute match="/*" attribute-name="type" attribute-value="xlsx-in-extract-format"/>
  <p:identity name="xlsx-in-extract-format-xtlcon-document"/>

  <!-- Put the extract format and the modification into the original zip file container so now we have everything in one document: -->
  <p:insert match="/*" position="first-child">
    <p:with-input port="source" pipe="@xlsx-container"/>
    <p:with-input port="insertion" pipe="@xlsx-in-extract-format-xtlcon-document"/>
  </p:insert>
  <p:insert match="/*" position="first-child">
    <p:with-input port="insertion" pipe="@modification-specification-xtlcon-document"/>
  </p:insert>

  <!-- Prepare the modification part (convert name references into index coordinates): -->
  <p:xslt>
    <p:with-input port="stylesheet" href="xsl-modify-xlsx/modify-xlsx-prepare.xsl"/>
  </p:xslt>

  <!-- First perform a rather raw merge of the modifications in the Excel contents: -->
  <p:xslt>
    <p:with-input port="stylesheet" href="xsl-modify-xlsx/modify-xlsx-1.xsl"/>
  </p:xslt>

  <!-- Now sort and remove doubles: -->
  <p:xslt>
    <p:with-input port="stylesheet" href="xsl-modify-xlsx/modify-xlsx-2.xsl"/>
  </p:xslt>

  <!-- Create the output xlsx: -->
  <xtlcon:container-to-zip href-target-zip="{$xlsx-href-out}"/>

  <!-- Create the output xml: -->
  <p:add-attribute attribute-name="xlsx-href-in" match="/*">
    <p:with-input port="source" pipe="@modify-xlsx-original-input"/>
    <p:with-option name="attribute-value" select="$xlsx-href-in"/>
  </p:add-attribute>
  <p:add-attribute attribute-name="xlsx-href-out" match="/*">
    <p:with-option name="attribute-value" select="$xlsx-href-out"/>
  </p:add-attribute>
  <p:add-attribute attribute-name="timestamp" match="/*">
    <p:with-option name="attribute-value" select="string(current-dateTime())"/>
  </p:add-attribute>

</p:declare-step>
