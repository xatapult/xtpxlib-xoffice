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
  
 <!-- ================================================================== -->
  
  <p:declare-step type="xtlxo:modify-xlsx">
    
    <p:documentation>TBD</p:documentation>
    
    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    <!-- SETUP: -->
    
    <p:input port="source" primary="true" sequence="false">
      <p:documentation>The modification specification. See `../../xsd/xlsx-modify.xsd`.</p:documentation>
    </p:input>
    
    <p:option name="xlsx-href-in" required="true">
      <p:documentation>URI of the input xlsx file to process</p:documentation>
    </p:option>
    
    <p:option name="xlsx-href-out" required="true">
      <p:documentation>URI of the output xlsx file.</p:documentation>
    </p:option>
    
    <p:output port="result" primary="true" sequence="false">
      <p:documentation>The output is identical to the input but with `@timestamp`, `@`xlsx-href-in` and `@xlsx-href-out` added to 
        the root element.</p:documentation>
    </p:output>
    
    <p:import href="../../../xtpxlib-container/xplmod/container.mod/container.mod.xpl"/>
    
    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    
    <p:identity name="modify-xlsx-original-input"/>
    
    <!-- Create a container for the Excel input: -->
    <xtlcon:zip-to-container>
      <p:with-option name="href-source-zip" select="$xlsx-href-in"/>
      <p:with-option name="add-document-target-paths" select="true()"/>
    </xtlcon:zip-to-container>
    
    <xtlcon:container-to-zip>
      <p:with-option name="href-target-zip" select="$xlsx-href-out"/>
    </xtlcon:container-to-zip>
    
    <!--<!-\- Take the input (the modification XML specification) and wrap it in a <containerFile> element. We later on
      insert this into the container created from the input xlsx:-\->
    <p:wrap match="/*" wrapper="bcmconv-lib:containerFile"/>
    <p:add-attribute attribute-name="type" match="/*">
      <p:with-option name="attribute-value" select="'xlsx-modification'"/>
    </p:add-attribute>
    <p:identity name="modify-xlsx-input-wrapped"/>
    <p:sink/>
    
    <!-\- Get the input xlsx in a container: -\->
    <bcmconv-lib:zip2container>
      <p:with-option name="debug" select="$debug"/>
      <p:with-option name="zip-href" select="$input-xlsx-href"/>
    </bcmconv-lib:zip2container>
    
    <!-\- Add the modification XML to the container: -\->
    <p:insert position="first-child" match="/*">
      <p:input port="insertion">
        <p:pipe port="result" step="modify-xlsx-input-wrapped"/>
      </p:input>
    </p:insert>
    
    <!-\- First perform a rather raw merge of the modifications in the Excel contents: -\->
    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/modify-xlsx-1.xsl"/>
      </p:input>
      <p:with-param name="debug" select="$debug"/>
    </p:xslt>
    
    <!-\- Now sort and remove doubles: -\->
    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/modify-xlsx-2.xsl"/>
      </p:input>
      <p:with-param name="debug" select="$debug"/>
    </p:xslt>
    
    <!-\- Create the output xlsx: -\->
    <bcmconv-lib:container2zip>
      <p:with-option name="zip-result-href" select="$output-xlsx-href"/>
    </bcmconv-lib:container2zip>
    
    <!-\- Create the output xml: -\->
    <p:sink/>
    <p:identity>
      <p:input port="source">
        <p:pipe port="result" step="modify-xlsx-input"/>
      </p:input>
    </p:identity>
    <p:add-attribute attribute-name="input-xlsx" match="/*">
      <p:with-option name="attribute-value" select="$input-xlsx-href"/>
    </p:add-attribute>
    <p:add-attribute attribute-name="output-xlsx" match="/*">
      <p:with-option name="attribute-value" select="$output-xlsx-href"/>
    </p:add-attribute>-->
    
  </p:declare-step>
  

</p:library>
