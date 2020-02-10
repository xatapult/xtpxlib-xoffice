<?xml version="1.0" encoding="UTF-8"?>
<p:library xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:xtpxplib="http://www.xatapult.nl/namespaces/common/xproc/library" xmlns:bcmconv-lib="http://www.malmberg.nl/namespaces/bcmconverters/library"
  xmlns:pxp="http://exproc.org/proposed/steps" version="1.0" xpath-version="2.0" exclude-inline-prefixes="#all">
  <!-- ================================================================== -->
  <!--* 
    Library for common operations on MS Office files.
    Currently contains only a conversion for Excel files.
    
    Due to refactoring it is a bit mashed up: The conversion uses the bcmconv-lib facilities for containers but
    produces output in an xtp namespace... (and is itself in such a namespace).
    This is all not to affect existing code that extensively used the original output format. However we had to make the 
    office conversion more genric and open, therefore this strange setup.
  -->
  <!-- ================================================================== -->
  <!-- PROCESS XSLX FILE INTO SOMETHING MANAGEABLE: -->
  <!-- -->
  <p:declare-step type="xtpxplib:extract-xlsx">
    <!-- -->
    <p:documentation>
      Extracts the Excel into something more helpfull and maneagable.
    </p:documentation>
    <!-- -->
    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    <!-- SETUP: -->
    <!-- -->
    <p:option name="debug" required="false" select="string(false())"/>
    <!-- -->
    <p:option name="href" required="true">
      <p:documentation>URI of the xlsx file to process</p:documentation>
    </p:option>
    <!-- -->
    <p:output port="result" primary="true" sequence="false"/>
    <!-- -->
    <p:import href="../_zip2container/zip2container.xpl"/>
    <!-- -->
    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    <!-- PROCESS THE XLSX FILE: -->
    <!-- -->
    <!-- Extract all XML content: -->
    <bcmconv-lib:zip2container>
      <p:with-option name="zip-href" select="$href"/>
      <p:with-option name="debug" select="$debug"/>
    </bcmconv-lib:zip2container>
    <!-- -->
    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/extract-xlsx-1.xsl"/>
      </p:input>
      <p:with-param name="debug" select="$debug"/>
    </p:xslt>
    <!-- -->
  </p:declare-step>
  <!-- -->
  <!-- ================================================================== -->
  <!-- MODIFY XLSX FILE: -->
  <!-- -->
  <p:declare-step type="xtpxplib:modify-xlsx">
    <!-- -->
    <p:documentation>Modifies an Excel file.</p:documentation>
    <!-- -->
    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    <!-- SETUP: -->
    <!-- -->
    <p:option name="debug" required="false" select="string(false())"/>
    <!-- -->
    <p:input port="source" primary="true" sequence="false">
      <p:documentation>The modification specification. See xsd/excel-modifications.xsd</p:documentation>
    </p:input>
    <!-- -->
    <p:option name="input-xlsx-href" required="true">
      <p:documentation>URI of the input xlsx file to process</p:documentation>
    </p:option>
    <!-- -->
    <p:option name="output-xlsx-href" required="true">
      <p:documentation>URI of the output xlsx file.</p:documentation>
    </p:option>
    <!-- -->
    <p:output port="result" primary="true" sequence="false">
      <p:documentation>The output is equal to the input but with @input-xlsx and @output-xlsx added to the root element</p:documentation>
    </p:output>
    <!-- -->
    <p:import href="../_zip2container/zip2container.xpl"/>
    <!-- -->
    <!-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -->
    <!-- -->
    <p:identity name="modify-xlsx-input"/>
    <!-- -->
    <!-- Take the input (the modification XML specification) and wrap it in a <containerFile> element. We later on
      insert this into the container created from the input xlsx:-->
    <p:wrap match="/*" wrapper="bcmconv-lib:containerFile"/>
    <p:add-attribute attribute-name="type" match="/*">
      <p:with-option name="attribute-value" select="'xlsx-modification'"/>
    </p:add-attribute>
    <p:identity name="modify-xlsx-input-wrapped"/>
    <p:sink/>
    <!-- -->
    <!-- Get the input xlsx in a container: -->
    <bcmconv-lib:zip2container>
      <p:with-option name="debug" select="$debug"/>
      <p:with-option name="zip-href" select="$input-xlsx-href"/>
    </bcmconv-lib:zip2container>
    <!-- -->
    <!-- Add the modification XML to the container: -->
    <p:insert position="first-child" match="/*">
      <p:input port="insertion">
        <p:pipe port="result" step="modify-xlsx-input-wrapped"/>
      </p:input>
    </p:insert>
    <!-- -->
    <!-- First perform a rather raw merge of the modifications in the Excel contents: -->
    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/modify-xlsx-1.xsl"/>
      </p:input>
      <p:with-param name="debug" select="$debug"/>
    </p:xslt>
    <!-- -->
    <!-- Now sort and remove doubles: -->
    <p:xslt>
      <p:input port="stylesheet">
        <p:document href="xsl/modify-xlsx-2.xsl"/>
      </p:input>
      <p:with-param name="debug" select="$debug"/>
    </p:xslt>
    <!-- -->
    <!-- Create the output xlsx: -->
    <bcmconv-lib:container2zip>
      <p:with-option name="zip-result-href" select="$output-xlsx-href"/>
    </bcmconv-lib:container2zip>
    <!-- -->
    <!-- Create the output xml: -->
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
    </p:add-attribute>
    <!-- -->
  </p:declare-step>
  <!-- -->
</p:library>
