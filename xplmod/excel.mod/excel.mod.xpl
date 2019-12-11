<?xml version="1.0" encoding="UTF-8"?>
<p:library xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step" xmlns:xtlxo="http://www.xtpxlib.nl/ns/xoffice"
  xmlns:xtlcon="http://www.xtpxlib.nl/ns/container" xmlns:xtlc="http://www.xtpxlib.nl/ns/common" xmlns:pxp="http://exproc.org/proposed/steps"
  xmlns:local="#local.excel-to-xml.mod.xpl" version="1.0" xpath-version="2.0" exclude-inline-prefixes="#all">
 
  <p:documentation>
    Conversions for Excel (`.xlsx`) files.
  </p:documentation>
  
  <!-- ================================================================== -->
  
  <p:declare-step type="xtlxo:extract-xlsx">

    <p:documentation>
      Extracts the contents of an Excel (`.xlsx`) file in a more useable [XML format](%xlsx-extract.xsd).
    </p:documentation>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    <!-- SETUP: -->

    <p:option name="xlsx-href" required="true">
      <p:documentation>Document reference of the `.xlsx` file to process (must have `file://` in front).</p:documentation>
    </p:option>

    <p:output port="result" primary="true" sequence="false">
      <p:documentation>
        The resulting XML representation of the Excel file.
      </p:documentation>
    </p:output>

    <p:import href="../../../xtpxlib-container/xplmod/container.mod/container.mod.xpl"/>

    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->

    <!-- Extract all XML content: -->
    <xtlcon:zip-to-container>
      <p:with-option name="href-source-zip" select="$xlsx-href"/>
      <p:with-option name="add-document-target-paths" select="false()"/>
    </xtlcon:zip-to-container>

    <!-- Transform the contents into something manageable: -->
    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/extract-xlsx-1.xsl"/>
      </p:input>
      <p:with-param name="null" select="()"/>
    </p:xslt>
    
    <!-- Remove any empty rows and cells: -->
    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/extract-xlsx-2.xsl"/>
      </p:input>
      <p:with-param name="null" select="()"/>
    </p:xslt>
    
  </p:declare-step>

</p:library>
